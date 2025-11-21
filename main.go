package main

import (
	"bufio"
	"encoding/json"
	"fmt"
	"log"
	"os"
	"strings"

	"github.com/xuri/excelize/v2"
)

func main() {
	inputFileName := "文件.xlsx"
	outputFileName := "output.xlsx"

	// Open the Excel file
	f, err := excelize.OpenFile(inputFileName)
	if err != nil {
		log.Fatalf("Failed to open %s: %v\nMake sure the file exists in the current directory.", inputFileName, err)
	}
	defer func() {
		if err := f.Close(); err != nil {
			log.Printf("Error closing file: %v", err)
		}
	}()

	// Get the first sheet name
	sheetName := f.GetSheetName(0)
	if sheetName == "" {
		log.Fatal("No sheets found in the Excel file")
	}

	// Get all rows from the sheet

rows, err := f.GetRows(sheetName)
	if err != nil {
		log.Fatalf("Failed to get rows: %v", err)
	}

	fmt.Printf("Successfully read %d rows from %s.\n", len(rows), inputFileName)

	// Prompt user for key
	reader := bufio.NewReader(os.Stdin)
	fmt.Print("Enter the JSON key to extract: ")
	keyToExtract, _ := reader.ReadString('\n')
	keyToExtract = strings.TrimSpace(keyToExtract)

	if keyToExtract == "" {
		log.Fatal("Key cannot be empty")
	}

	// Create a new Excel file for output
	outputFile := excelize.NewFile()
	outputSheet := "Sheet1"
	// Create a new sheet if it doesn't exist (NewFile creates Sheet1 by default)
	index, err := outputFile.NewSheet(outputSheet)
    if err != nil {
        log.Fatalf("Failed to create new sheet: %v", err)
    }
	outputFile.SetActiveSheet(index)

	validRowCount := 0

	for _, row := range rows {
		if len(row) == 0 {
			continue
		}

		// Assume the JSON is in the first column (Column A)
		cellValue := strings.TrimSpace(row[0])

		// Basic check: Must start with '[' to be a JSON array
		if !strings.HasPrefix(cellValue, "[") {
			// Likely a header or empty or malformed
			fmt.Printf("Skipping row (not an array): %s\n", cellValue)
			continue
		}

		// Normalize: Replace single quotes with double quotes (common in Python exports)
		// Note: This is a simple fix and might break if strings strictly contain single quotes.
		cellValue = strings.ReplaceAll(cellValue, "'", "\"")

		// Try to parse as JSON array of objects
		var data []map[string]interface{}
		err := json.Unmarshal([]byte(cellValue), &data)
		if err != nil {
			// Debug: Print why it failed
			if len(cellValue) > 50 {
				fmt.Printf("Skipping invalid JSON (start: %s...): %v\n", cellValue[:50], err)
			} else {
				fmt.Printf("Skipping invalid JSON (%s): %v\n", cellValue, err)
			}
			continue
		}

		validRowCount++
		
		// Extract values
		var extractedValues []string
		for _, item := range data {
			if val, ok := item[keyToExtract]; ok {
				// Convert value to string representation
				extractedValues = append(extractedValues, fmt.Sprintf("%v", val))
			}
		}

		// Write to output file
		// 1. Write comma-joined string to the first column (Column A)
		joinedString := strings.Join(extractedValues, ",")
		cellName, _ := excelize.CoordinatesToCellName(1, validRowCount)
		outputFile.SetCellValue(outputSheet, cellName, joinedString)

		// 2. Write individual values starting from the second column (Column B)
		for cIdx, val := range extractedValues {
			cellName, _ := excelize.CoordinatesToCellName(cIdx+2, validRowCount)
			outputFile.SetCellValue(outputSheet, cellName, val)
		}
	}

	if err := outputFile.SaveAs(outputFileName); err != nil {
		log.Fatalf("Failed to save output file: %v", err)
	}

	fmt.Printf("Processed %d valid rows.\n", validRowCount)
	fmt.Printf("Output saved to %s\n", outputFileName)
}
