package main

import (
	"flag"
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/lukasjarosch/go-docx"
)

var verbose bool

func parseMarkdown(markdownFile string) map[string]string {
	content, err := os.ReadFile(markdownFile)
	if err != nil {
		panic(err)
	}

	markdown := string(content)
	lines := strings.Split(markdown, "\n")

	data := make(map[string]string)
	currentPrefix := ""
	currentKey := ""
	currentValue := ""
	previousLine := ""

	for _, line := range lines {
		line = strings.TrimSpace(line)

		if strings.HasPrefix(line, "###") {
			// Third-level heading
			if verbose {
				fmt.Println("Found heading: " + line)
			}
			heading := strings.TrimPrefix(line, "###")
			key := strings.Map(func(r rune) rune {
				if r == ' ' || r == '_' || ('a' <= r && r <= 'z') || ('A' <= r && r <= 'Z') || ('0' <= r && r <= '9') || r == '-' {
					return r
				}
				return -1
			}, strings.ToLower(heading))
			if verbose {
				fmt.Println("Sanitized key: " + key)
			}
			key = strings.ToLower(strings.ReplaceAll(strings.ReplaceAll(strings.TrimSpace(key), " ", "-"), "_", "-"))
			if verbose {
				fmt.Println("key to kebab case: " + key)
			}
			if currentPrefix != "" {
				key = currentPrefix + "-" + key
			}

			if currentKey != "" {
				data[currentKey] = strings.TrimSpace(processValue(currentValue))
			}

			currentKey = key
			currentValue = ""
		} else if strings.HasPrefix(line, ":") {
			// Definition list item
			parts := strings.SplitN(line, ":", 2)
			if len(parts) == 2 {
				// key := strings.ToLower(strings.ReplaceAll(strings.ReplaceAll(strings.TrimSpace(previousLine), " ", "-"), "_", "-"))
				key := strings.Map(func(r rune) rune {
					if r == ' ' || r == '_' || ('a' <= r && r <= 'z') || ('A' <= r && r <= 'Z') || ('0' <= r && r <= '9') || r == '-' {
						return r
					}
					return -1
				}, strings.ToLower(previousLine))
				key = strings.ReplaceAll(strings.ReplaceAll(string(key), " ", "-"), "_", "-")
				value := strings.TrimSpace(parts[1])
				if currentPrefix != "" {
					key = currentPrefix + "-" + key
				}
				data[key] = value
			}
		} else if strings.HasPrefix(line, "##") {
			// Second-level heading
			if currentKey != "" {
				data[currentKey] = strings.TrimSpace(processValue(currentValue))
			}
			currentKey = ""
			currentValue = ""

			currentPrefix = strings.Map(func(r rune) rune {
				if r == ' ' || r == '_' || ('a' <= r && r <= 'z') || ('A' <= r && r <= 'Z') || ('0' <= r && r <= '9') || r == '-' {
					return r
				}
				return -1
			}, strings.ToLower(line))
			currentPrefix = strings.ToLower(strings.ReplaceAll(strings.ReplaceAll(strings.TrimSpace(strings.TrimPrefix(line, "##")), " ", "-"), "_", "-"))
		} else if currentKey != "" {
			// Append line to current value
			currentValue += line + "\n"
		}

		previousLine = line
	}

	// Handle the last heading or definition list item
	if currentKey != "" {
		data[currentKey] = strings.TrimSpace(processValue(currentValue))
	}
	if verbose {
		fmt.Printf("data length is %d\n", len(data))
		for key, value := range data {
			fmt.Printf("%s: %s\n", key, value)
		}
	}

	return data
}

func replaceMustacheTags(templateFile string, data map[string]string, outputFile string) {
	if verbose {
		fmt.Println("\nWill look for strings to replace now")
	}
	doc, err := docx.Open(templateFile)
	if err != nil {
		panic(err)
	}

	replaceMap := docx.PlaceholderMap{}
	for key, value := range data {
		replaceMap[key] = value
	}

	err = doc.ReplaceAll(replaceMap)
	if err != nil {
		panic(err)
	}

	err = doc.WriteToFile(outputFile)
	if err != nil {
		panic(err)
	}
}

func processValue(value string) string {
	listItems := strings.Split(value, "\n")
	var bulletPoints []string
	for _, item := range listItems {
		if strings.HasPrefix(item, "-") || strings.HasPrefix(item, "+") {
			item = strings.Replace(item, string(item[0]), "•", 1)
		}
		bulletPoints = append(bulletPoints, item)
	}
	return strings.Join(bulletPoints, "\n")
}

func main() {
	markdownFile := flag.String("markdown", "", "Path to the markdown file")
	templateFile := flag.String("template", "", "Path to the Word document template")
	outputFile := flag.String("output", "", "Path to the output Word document (optional)")
	flag.BoolVar(&verbose, "v", false, "Enable verbose output")
	flag.Parse()

	// Check if required arguments are provided
	if *markdownFile == "" {
		fmt.Println("Error: Markdown file path is required")
		return
	}
	if *templateFile == "" {
		fmt.Println("Error: Template file path is required")
		return
	}

	// Set default output file path if not provided
	if *outputFile == "" {
		*outputFile = strings.TrimSuffix(*markdownFile, filepath.Ext(*markdownFile)) + ".docx"
	}
	data := parseMarkdown(*markdownFile)
	replaceMustacheTags(*templateFile, data, *outputFile)
}
