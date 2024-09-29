package powerpoint

import (
	"image"
	"math"
	"path/filepath"
	"strings"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	"github.com/scjalliance/comshim"
	"github.com/sirupsen/logrus"
)

// Specifies a tri-state Boolean value
const (
	msoTrue  int = -1
	msoFalse int = 0
)

func filterName(outputFile string) string {
	if s := filepath.Ext(outputFile); len(s) > 1 {
		return strings.ToUpper(s[1:])
	}
	return "PNG"
}

func openPresentation(powerpoint *ole.IDispatch, inputFile string) (*ole.VARIANT, error) {
	presentations := oleutil.MustGetProperty(powerpoint, "Presentations").ToIDispatch()
	defer presentations.Release()

	return oleutil.CallMethod(presentations, "Open", inputFile, msoTrue, msoFalse, msoFalse)
}

func closePresentation(presentation *ole.VARIANT) error {
	if presentation.VT != ole.VT_EMPTY {
		oleutil.PutProperty(presentation.ToIDispatch(), "Saved", msoFalse)
		oleutil.CallMethod(presentation.ToIDispatch(), "Close")
	}
	return presentation.Clear()
}

func getSlide(presentation *ole.IDispatch, index int) (*ole.IDispatch, error) {
	slides := oleutil.MustGetProperty(presentation, "Slides").ToIDispatch()
	defer slides.Release()

	count := int(oleutil.MustGetProperty(slides, "Count").Val)
	index = int(math.Min(math.Max(float64(index), 1), float64(count)))

	return oleutil.MustCallMethod(slides, "Item", index).ToIDispatch(), nil
}

func getActiveObject() (*ole.IDispatch, error) {
	clsid, err := ole.CLSIDFromProgID("PowerPoint.Application")
	if err != nil {
		return nil, err
	}

	unknown, err := ole.CreateInstance(clsid, nil)
	if err != nil {
		return nil, err
	}
	defer unknown.Release()

	return unknown.QueryInterface(ole.IID_IDispatch)
}

// IsAvailable checks if PowerPoint is available and logs its version
func IsAvailable() bool {
	comshim.Add(1)
	defer comshim.Done()

	powerpoint, err := getActiveObject()
	if err != nil {
		return false
	}
	defer func() {
		oleutil.CallMethod(powerpoint, "Quit")
		powerpoint.Release()
	}()

	version := oleutil.MustGetProperty(powerpoint, "Version").ToString()
	build := oleutil.MustGetProperty(powerpoint, "Build").ToString()
	operatingSystem := oleutil.MustGetProperty(powerpoint, "OperatingSystem").ToString()

	logrus.Debugf("PowerPoint v%s build %s on %s", version, build, operatingSystem)
	return true
}

// GetDisplayBounds returns the dimensions of the PowerPoint slide
func GetDisplayBounds(inputFile string) (image.Rectangle, error) {
	comshim.Add(1)
	defer comshim.Done()

	powerpoint, err := getActiveObject()
	if err != nil {
		return image.Rectangle{}, err
	}
	defer func() {
		oleutil.CallMethod(powerpoint, "Quit")
		powerpoint.Release()
	}()

	presentation, err := openPresentation(powerpoint, inputFile)
	if err != nil {
		return image.Rectangle{}, err
	}
	defer closePresentation(presentation)

	slideMaster := oleutil.MustGetProperty(presentation.ToIDispatch(), "SlideMaster").ToIDispatch()
	defer slideMaster.Release()

	width := int(oleutil.MustGetProperty(slideMaster, "Width").Value().(float32))
	height := int(oleutil.MustGetProperty(slideMaster, "Height").Value().(float32))

	return image.Rect(0, 0, width, height), nil
}

// Export exports the slide, using the specified graphics filter, and saves
// the exported file under the specified outpt file name
func Export(inputFile, outputFile string, width, height, index int) error {
	comshim.Add(1)
	defer comshim.Done()

	powerpoint, err := getActiveObject()
	if err != nil {
		return err
	}
	defer func() {
		oleutil.CallMethod(powerpoint, "Quit")
		powerpoint.Release()
	}()

	presentation, err := openPresentation(powerpoint, inputFile)
	if err != nil {
		return err
	}
	defer closePresentation(presentation)

	slide, err := getSlide(presentation.ToIDispatch(), index)
	if err != nil {
		return err
	}
	defer slide.Release()

	filterName := filterName(outputFile)
	if _, err := oleutil.CallMethod(slide, "Export", outputFile, filterName, width, height); err != nil {
		return err
	}
	return nil
}
