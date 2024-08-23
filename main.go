package main

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"log"
	"os"
	"os/user"
	"path/filepath"
	"runtime"
	"strings"

	"golang.org/x/net/html/charset"
)

// Helper function to read a file from within the zip archive
func readFileFromZip(zipReader *zip.ReadCloser, filename string) ([]byte, error) {
	for _, file := range zipReader.File {
		if file.Name == filename {
			rc, err := file.Open()
			if err != nil {
				return nil, err
			}
			defer rc.Close()

			buf := new(bytes.Buffer)
			_, err = io.Copy(buf, rc)
			if err != nil {
				return nil, err
			}
			return buf.Bytes(), nil
		}
	}
	return nil, fmt.Errorf("file %s not found in zip", filename)
}

func extractTextFromSlide(slideData []byte) (string, error) {
	decoder := xml.NewDecoder(bytes.NewReader(slideData))
	decoder.CharsetReader = charset.NewReaderLabel

	var slideText strings.Builder
	inTextElement := false

	for {
		tok, err := decoder.Token()
		if err != nil {
			if err == io.EOF {
				break
			}
			return "", err
		}

		switch element := tok.(type) {
		case xml.StartElement:
			if element.Name.Local == "t" {
				inTextElement = true
			}
		case xml.EndElement:
			if element.Name.Local == "t" {
				inTextElement = false
				// Add a space after each text block
				slideText.WriteString(" ")
			}
		case xml.CharData:
			if inTextElement {
				slideText.Write(element)
			}
		}
	}

	// Use strings.Fields to remove extra spaces
	trimmedText := strings.Join(strings.Fields(slideText.String()), " ")
	return trimmedText, nil
}

func getDesktopPath() (string, error) {
	usr, err := user.Current()
	if err != nil {
		return "", err
	}

	// Cross-platform desktop path determination
	desktopPath := ""

	switch os := runtime.GOOS; os {
	case "windows":
		desktopPath = filepath.Join(usr.HomeDir, "Desktop")
	case "darwin":
		desktopPath = filepath.Join(usr.HomeDir, "Desktop")
	case "linux":
		desktopPath = filepath.Join(usr.HomeDir, "Desktop")
	default:
		return "", fmt.Errorf("unsupported platform")
	}

	return desktopPath, nil
}

func main() {
	if len(os.Args) < 2 {
		log.Fatalf("Usage: %s <input.pptx>", os.Args[0])
	}

	pptPath := os.Args[1]
	outputFileName := strings.TrimSuffix(filepath.Base(pptPath), ".pptx") + "_output.txt"

	desktopPath, err := getDesktopPath()
	if err != nil {
		log.Fatalf("Error determining desktop path: %v", err)
	}

	outputFilePath := filepath.Join(desktopPath, outputFileName)

	// Open the PowerPoint file as a zip archive
	zipReader, err := zip.OpenReader(pptPath)
	if err != nil {
		log.Fatalf("Error opening PowerPoint file: %v", err)
	}
	defer zipReader.Close()

	var fullText strings.Builder

	// Loop over all slides
	for i := 1; ; i++ {
		slideName := fmt.Sprintf("ppt/slides/slide%d.xml", i)
		slideData, err := readFileFromZip(zipReader, slideName)
		if err != nil {
			// If the file is not found, assume we've processed all slides
			if strings.Contains(err.Error(), "not found") {
				break
			}
			log.Fatalf("Error reading slide file: %v", err)
		}

		slideText, err := extractTextFromSlide(slideData)
		if err != nil {
			log.Fatalf("Error extracting text from slide: %v", err)
		}

		fullText.WriteString(slideText)
		// Add two newlines between slides
		fullText.WriteString("\n\n")
	}

	// Write the extracted text to a file on the desktop
	err = os.WriteFile(outputFilePath, []byte(fullText.String()), 0644)
	if err != nil {
		log.Fatalf("Error writing to output file: %v", err)
	}

	fmt.Printf("Text extracted and written to %s\n", outputFilePath)
}
