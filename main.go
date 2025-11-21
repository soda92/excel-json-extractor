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
		cellValue := row[0]

		// Try to parse as JSON array of objects
		var data []map[string]interface{}
		err := json.Unmarshal([]byte(cellValue), &data)
		if err != nil {
			// Skip invalid JSON rows
			// fmt.Printf("Skipping row %d: Invalid JSON\n", rIdx+1)
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
		// If extractedValues is empty, we leave the row empty or write nothing?
		// The prompt implies "split in excel", so each value in a separate column.
		for cIdx, val := range extractedValues {
			// Excel columns are 1-based, rows are 1-based.
			// Output row should match valid row count (packed) or original row index?
			// "remove invalid json rows" implies the output might be compacted.
			// Let's compact them (use validRowCount).
			
			cellName, _ := excelize.CoordinatesToCellName(cIdx+1, validRowCount)
			outputFile.SetCellValue(outputSheet, cellName, val)
		}
	}

	if err := outputFile.SaveAs(outputFileName); err != nil {
		log.Fatalf("Failed to save output file: %v", err)
	}

	fmt.Printf("Processed %d valid rows.\n", validRowCount)
	fmt.Printf("Output saved to %s\n", outputFileName)
}
