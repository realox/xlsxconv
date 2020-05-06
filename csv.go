package xlsxconv

import (
	"encoding/csv"
	"fmt"
	"io"

	"github.com/plandem/xlsx"
)

// ToCSV converts an XLSX to a CSV.
func ToCSV(out io.Writer, f interface{}, sheetIndex int) (err error) {
	xl, err := xlsx.Open(f)
	if err != nil {
		return err
	}
	defer xl.Close()

	sheet := xl.Sheet(sheetIndex, xlsx.SheetModeStream)
	defer sheet.Close()

	cw := csv.NewWriter(out)
	for rows := sheet.Rows(); rows.HasNext(); {
		lines := []string{}
		if _, row := rows.Next(); row != nil {
			for cells := row.Cells(); cells.HasNext(); {
				_, _, cell := cells.Next()
				str := cell.Value()
				lines = append(lines, fmt.Sprintf("%s", str))
			}
		}
		cw.Write(lines)
	}
	cw.Flush()

	return cw.Error()
}
