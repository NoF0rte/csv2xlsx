package csv2xlsx

import (
	"github.com/xuri/excelize/v2"
)

func ToXLSX(rows [][]string, hasHeader bool, output string) error {
	sheet := "Sheet1"
	f := excelize.NewFile()

	var err error
	var headerStyleId int
	if hasHeader {
		headerStyleId, err = f.NewStyle(&excelize.Style{
			Font: &excelize.Font{
				Bold: true,
				Size: 12,
			},
		})
		if err != nil {
			return err
		}
	}

	for i, row := range rows {
		for j, col := range row {
			coordinates, err := excelize.CoordinatesToCellName(j+1, i+1)
			if err != nil {
				return err
			}

			f.SetCellValue(sheet, coordinates, col)

			if i == 0 && hasHeader {
				err = f.SetCellStyle(sheet, coordinates, coordinates, headerStyleId)
				if err != nil {
					return err
				}
			}
		}
	}

	if hasHeader {
		endCoord, err := excelize.CoordinatesToCellName(len(rows[0]), len(rows))
		if err != nil {
			return err
		}

		err = f.AddTable(sheet, "A1", endCoord, ``)
		if err != nil {
			return err
		}
	}

	return f.SaveAs(output)
}
