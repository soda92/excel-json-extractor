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
		log.Fatalf("无法打开 %s: %v\n请确保文件存在于当前目录下。", inputFileName, err)
	}
	defer func() {
		if err := f.Close(); err != nil {
			log.Printf("关闭文件时出错: %v", err)
		}
	}()

	// Get the first sheet name
	sheetName := f.GetSheetName(0)
	if sheetName == "" {
		log.Fatal("Excel 文件中未找到工作表")
	}

	// Get all rows from the sheet
	rows, err := f.GetRows(sheetName)
	if err != nil {
		log.Fatalf("获取行数据失败: %v", err)
	}

	fmt.Printf("成功从 %s 读取了 %d 行。\n", inputFileName, len(rows))

	// Prompt user for key
	reader := bufio.NewReader(os.Stdin)
	fmt.Print("请输入要提取的 JSON 键名: ")
	keyToExtract, _ := reader.ReadString('\n')
	keyToExtract = strings.TrimSpace(keyToExtract)

	if keyToExtract == "" {
		log.Fatal("键名不能为空")
	}

	// Create a new Excel file for output
	outputFile := excelize.NewFile()
	outputSheet := "Sheet1"
	// Create a new sheet if it doesn't exist (NewFile creates Sheet1 by default)
	index, err := outputFile.NewSheet(outputSheet)
	if err != nil {
		log.Fatalf("创建新工作表失败: %v", err)
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
			fmt.Printf("跳过该行（非数组）: %s\n", cellValue)
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
				fmt.Printf("跳过无效 JSON (开头: %s...): %v\n", cellValue[:50], err)
			} else {
				fmt.Printf("跳过无效 JSON (%s): %v\n", cellValue, err)
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
		log.Fatalf("保存输出文件失败: %v", err)
	}

	fmt.Printf("共处理 %d 行有效数据。\n", validRowCount)
	fmt.Printf("结果已保存至 %s\n", outputFileName)
}
