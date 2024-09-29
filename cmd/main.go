package main

import (
	"flag"
	"fmt"
	"os"
	"path/filepath"

	"github.com/optoma-go/ppt-screenshot/pkg/powerpoint"
)

func main() {
	inputFile := flag.String("input", "", "input presentation filename (required)")
	outputFile := flag.String("output", "", "output screenshot image filename (required)")
	force := flag.Bool("force", false, "overwrite existing output file")
	width := flag.Int("width", 0, "output image scale width")
	height := flag.Int("height", 0, "output image scale height")
	index := flag.Int("index", 1, "use slide item of presentation to output")
	flag.Parse()

	// Check input file
	if *inputFile == "" {
		fmt.Fprintln(os.Stderr, "please enter input presentation filename (-input)")
		return
	}
	if _, err := os.Stat(*inputFile); os.IsNotExist(err) {
		fmt.Fprintln(os.Stderr, "the input presentation does not exist")
		return
	}
	if !filepath.IsAbs(*inputFile) {
		absPath, err := filepath.Abs(*inputFile)
		if err != nil {
			fmt.Fprintln(os.Stderr, "failed to get absolute path: %w", err)
			return
		}
		*inputFile = absPath
	}

	// Check output file
	if *outputFile == "" {
		fmt.Fprintln(os.Stderr, "please enter output image filename (-output)")
		return
	}
	if _, err := os.Stat(*outputFile); os.IsExist(err) && !*force {
		fmt.Fprintln(os.Stderr, "the output image file is exist")
		return
	}
	if !filepath.IsAbs(*outputFile) {
		absPath, err := filepath.Abs(*outputFile)
		if err != nil {
			fmt.Fprintln(os.Stderr, "failed to get absolute path: %w", err)
			return
		}
		*outputFile = absPath
	}

	if !powerpoint.IsAvailable() {
		fmt.Fprintln(os.Stderr, "the PowerPoint application is unavailable")
		return
	}

	bounds, err := powerpoint.GetDisplayBounds(*inputFile)
	if err != nil {
		fmt.Fprintln(os.Stderr, "failed to get presentation display bounds: %w", err)
		return
	}

	imgWidth := bounds.Dx()
	imgHeight := bounds.Dy()
	if *width != 0 {
		imgWidth = *width
	}
	if *height != 0 {
		imgHeight = *height
	}

	if err := powerpoint.Export(*inputFile, *outputFile, imgWidth, imgHeight, *index); err != nil {
		fmt.Fprintln(os.Stderr, "failed to export screenshot image: %w", err)
		return
	}
}
