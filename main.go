package main

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"log"
	"os"
	"strings"

	"golang.org/x/net/html/charset"
)

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
			}
		case xml.CharData:
			if inTextElement {
				slideText.Write(element)
			}
		}
	}

	trimmedText := strings.Join(strings.Fields(slideText.String()), " ")
	return trimmedText, nil
}

func main() {
	if len(os.Args) < 2 {
		log.Fatalf("Usage: %s <input.pptx>", os.Args[0])
	}

	pptPath := os.Args[1]
	outputFileName := strings.TrimSuffix(pptPath, ".pptx") + "_output.txt"

	zipReader, err := zip.OpenReader(pptPath)
	if err != nil {
		log.Fatalf("Error opening PowerPoint file: %v", err)
	}
	defer zipReader.Close()

	var fullText strings.Builder

	for i := 1; ; i++ {
		slideName := fmt.Sprintf("ppt/slides/slide%d.xml", i)
		slideData, err := readFileFromZip(zipReader, slideName)
		if err != nil {
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
		if i < len(zipReader.File) {
			fullText.WriteString("\n\n")
		}
	}

	err = os.WriteFile(outputFileName, []byte(fullText.String()), 0644)
	if err != nil {
		log.Fatalf("Error writing to output file: %v", err)
	}

	fmt.Printf("Text extracted and written to %s\n", outputFileName)
}
