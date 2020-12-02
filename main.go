package main

import (
	"flag"
	"os"
	"path"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"github.com/google/uuid"
)

var (
	sourceFile, destinationFile, dataSheetName, sourceHeaderName, refSheet string
	newFile                                                                bool
)

func getReferenceData(file *excelize.File, sheetName string) map[string]string {
	cols, err := file.GetCols(sheetName)
	if err != nil {
		panic("Cannot fetch reference data!")
	}

	var mappedData map[string]string = make(map[string]string)

	for _, value := range cols[0] {
		if len(strings.TrimSpace(value)) > 0 {
			mappedData[value] = strings.TrimSpace(value)
		}
	}

	return mappedData
}

func main() {
	flag.StringVar(&sourceFile, "input", "", "Input file path")
	flag.StringVar(&destinationFile, "output", uuid.New().String()+".xlsx", "name of destination file, defaults to a unique identifier")
	flag.StringVar(&dataSheetName, "sheet", "", "Name of sheet containing data")
	flag.StringVar(&sourceHeaderName, "head", "", "Column header name for search target column")
	flag.StringVar(&refSheet, "reference", "", "Name of sheet containing reference data")
	flag.BoolVar(&newFile, "saveAsNew", false, "Whether to save search in new file")

	flag.Parse()

	if newFile && !path.IsAbs(destinationFile) {
		home, err := os.UserHomeDir()
		if err != nil {
			panic("Cannot resolve home directory!")
		}
		destinationFile = path.Join(home, destinationFile)
		println("destination file will be located at:", destinationFile)
	}

	if len(strings.TrimSpace(sourceFile)) < 1 {
		panic("Invalid source file entered!")
	}

	_, err := os.Stat(sourceFile)
	if os.IsNotExist(err) {
		panic("Source file does not exist!")
	}

	file, err := excelize.OpenFile(sourceFile)
	if err != nil {
		panic(err)
	}

	println("Sheet name:", dataSheetName)
	println("Header name:", sourceHeaderName)

	referenceData := getReferenceData(file, refSheet)
	columns, err := file.GetCols(dataSheetName)
	if err != nil {
		panic(err)
	}
	rows, err := file.GetRows(dataSheetName)
	if err != nil {
		panic(err)
	}

	var rowCount int = 1
	newSheetName := "Filter Result"
	_ = file.NewSheet(newSheetName)
	for y, col := range columns {
		if strings.TrimSpace(col[0]) != sourceHeaderName {
			continue
		}
		for x, rowVals := range rows {
			axis, err := excelize.CoordinatesToCellName(1, rowCount)
			if err != nil {
				panic(err)
			}
			// Skip header row
			if x == 0 {
				file.SetSheetRow(newSheetName, axis, &rowVals)
				continue
			}
			if _, exists := referenceData[rowVals[y]]; exists {
				_ = file.SetSheetRow(newSheetName, axis, &rowVals)
				rowCount++
			}
		}
	}

	if !newFile {
		err = file.Save()
		if err != nil {
			panic(err)
		}
	} else {
		err = file.SaveAs(destinationFile)
		if err != nil {
			panic(err)
		}
	}
}
