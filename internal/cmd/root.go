package cmd

import (
	"encoding/csv"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"strings"

	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"
)

// rootCmd represents the base command when called without any subcommands
var rootCmd = &cobra.Command{
	Use:   "csv2xlsx",
	Short: "Convert CSV files to XLSX",
	Run: func(cmd *cobra.Command, args []string) {
		csvFile, _ := cmd.Flags().GetString("file")
		hasHeader, _ := cmd.Flags().GetBool("header")
		output, _ := cmd.Flags().GetString("output")

		rows, err := readCSV(csvFile)
		if err != nil {
			log.Fatal(err)
		}

		sheet := "Sheet1"
		f := excelize.NewFile()

		var headerStyleId int
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

		if output == "" {
			filename := filepath.Base(csvFile)
			fileWithoutExt := strings.TrimSuffix(filename, filepath.Ext(filename))
			output = fmt.Sprintf("%s.xlsx", fileWithoutExt)
		}

		if err := f.SaveAs(output); err != nil {
			fmt.Println(err)
		}
	},
}

func readCSV(file string) ([][]string, error) {
	f, err := os.Open(file)
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

// Execute adds all child commands to the root command and sets flags appropriately.
// This is called by main.main(). It only needs to happen once to the rootCmd.
func Execute() {
	err := rootCmd.Execute()
	if err != nil {
		os.Exit(1)
	}
}

func init() {
	rootCmd.Flags().StringP("file", "f", "", "The path to the csv file")
	rootCmd.Flags().StringP("output", "o", "", "Output file name")
	rootCmd.Flags().Bool("header", false, "The csv file includes a header row")
}
