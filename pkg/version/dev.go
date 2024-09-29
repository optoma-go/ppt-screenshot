//go:build !prod || !staging
// +build !prod !staging

package version

import "time"

func init() {
	if buildDate == "" {
		buildDate = time.Now().Format(time.RFC3339Nano)
	}
	if commitHash == "" {
		commitHash = "dev"
	}
}
