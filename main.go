package main

import (
	"encoding/csv"
	"flag"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"strings"

	"github.com/xuri/excelize/v2"
)

var csvFile string
var hasHeader bool
var headerStyleId int

func readCSV() ([][]string, error) {
	f, err := os.Open(csvFile)
	if err != nil {
		return nil, err
	}
	defer f.Close()

	csvReader := csv.NewReader(f)
	data, err := csvReader.ReadAll()
	if err != nil {
		return nil, err
	}

	if len(data) == 0 {
		return nil, fmt.Errorf("empty CSV file")
	}

	return data, nil
}

func main() {
	flag.Parse()

	rows, err := readCSV()
	if err != nil {
		log.Fatal(err)
	}

	sheet := "Sheet1"
	f := excelize.NewFile()

	if hasHeader {
		headerStyleId, err = f.NewStyle(&excelize.Style{
			Font: &excelize.Font{
				Bold: true,
				Size: 12,
			},
		})
		if err != nil {
			log.Fatal(err)
		}
	}

	for i, row := range rows {
		for j, col := range row {
			coordinates, err := excelize.CoordinatesToCellName(j+1, i+1)
			if err != nil {
				log.Fatal(err)
			}

			f.SetCellValue(sheet, coordinates, col)

			if i == 0 && hasHeader {
				err = f.SetCellStyle(sheet, coordinates, coordinates, headerStyleId)
				if err != nil {
					log.Fatal(err)
				}
			}
		}
	}

	if hasHeader {
		endCoord, err := excelize.CoordinatesToCellName(len(rows[0]), len(rows))
		if err != nil {
			log.Fatal(err)
		}

		err = f.AddTable(sheet, "A1", endCoord, ``)
		if err != nil {
			log.Fatal(err)
		}
	}

	filename := filepath.Base(csvFile)
	fileWithoutExt := strings.TrimSuffix(filename, filepath.Ext(filename))

	if err := f.SaveAs(fmt.Sprintf("%s.xlsx", fileWithoutExt)); err != nil {
		fmt.Println(err)
	}
}

func init() {
	flag.StringVar(&csvFile, "file", "", "The path to the csv file")
	flag.BoolVar(&hasHeader, "header", false, "The csv file includes a header row")
}
