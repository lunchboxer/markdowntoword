package main

import (
	"flag"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"unicode"

	"github.com/lukasjarosch/go-docx"
	"golang.org/x/text/cases"
)

var verbose bool

func sanitizeKey(s string) string {
	// Use Unicode-aware case folding
	caser := cases.Fold()
	s = caser.String(s)

	return strings.Map(func(r rune) rune {
		if unicode.IsLetter(r) || unicode.IsNumber(r) || r == ' ' || r == '_' || r == '-' {
			return r
		}
		return -1
	}, s)
}

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
			key := sanitizeKey(heading)
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
				key := sanitizeKey(previousLine)
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

			currentPrefix = sanitizeKey(line)
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
		// Remove markdown formatting characters
		cleanValue := strings.ReplaceAll(strings.ReplaceAll(value, "**", ""), "*", "")
		replaceMap[key] = cleanValue
	}

	for key, value := range replaceMap {
		fmt.Printf("%s: %s\n", key, value)
	}

	err = doc.ReplaceAll(replaceMap)
	if err != nil {
		fmt.Printf("Error replacing placeholders: %v\n", err)
	} else {
		if verbose {
			fmt.Println("Replacements completed successfully")
		}
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

func replaceMarkdownFormatting(value, delimiter, style string) string {
	parts := strings.Split(value, delimiter)
	var styledParts []string
	
	for i, part := range parts {
		if i%2 == 1 { // Only style the odd-numbered segments between delimiters
			styledParts = append(styledParts, part)
		} else {
			styledParts = append(styledParts, part)
		}
	}
	return strings.Join(styledParts, "")
}

func main() {
	flag.Usage = func() {
		fmt.Fprintf(flag.CommandLine.Output(), "Usage: %s -markdown INPUT.md -template TEMPLATE.docx [options]\n", os.Args[0])
		fmt.Fprintln(flag.CommandLine.Output(), "\nConverts markdown documentation to Word document using a template")
		fmt.Fprintln(flag.CommandLine.Output(), "\nRequired flags:")
		flag.PrintDefaults()
		fmt.Fprintln(flag.CommandLine.Output(), "\nExamples:")
		fmt.Fprintln(flag.CommandLine.Output(), "  Generate document with default output name:")
		fmt.Fprintf(flag.CommandLine.Output(), "  %s -markdown spec.md -template template.docx\n", os.Args[0])
		fmt.Fprintln(flag.CommandLine.Output(), "  Generate document with custom output name:")
		fmt.Fprintf(flag.CommandLine.Output(), "  %s -markdown spec.md -template template.docx -output final.docx\n", os.Args[0])
	}

	var helpFlag bool
	markdownFile := flag.String("markdown", "", "Input markdown file containing documentation content (required)")
	templateFile := flag.String("template", "", "Input Word template document with {{placeholders}} (required)")
	outputFile := flag.String("output", "", "Output Word document filename (default: input name with .docx extension)")
	flag.BoolVar(&verbose, "v", false, "Enable verbose debugging output")
	flag.BoolVar(&helpFlag, "h", false, "Show this help message")
	flag.Parse()

	if helpFlag {
		flag.Usage()
		os.Exit(0)
	}

	// Check required arguments or show help
	if *markdownFile == "" || *templateFile == "" {
		flag.Usage()
		os.Exit(1)
	}

	// Set default output file path if not provided
	if *outputFile == "" {
		*outputFile = strings.TrimSuffix(*markdownFile, filepath.Ext(*markdownFile)) + ".docx"
	}
	data := parseMarkdown(*markdownFile)
	replaceMustacheTags(*templateFile, data, *outputFile)
}
