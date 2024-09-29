package version

var (
	version    string = "v0.0.0-unset"
	buildDate  string = ""
	commitHash string = ""
)

func Version() string {
	return version
}

func BuildDate() string {
	return buildDate
}

func CommitHash() string {
	return commitHash
}
