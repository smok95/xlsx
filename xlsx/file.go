package xlsx

import (
	"fmt"
	"sync"

	"github.com/tealeg/xlsx"
)

type XlsxFile struct {
	file *xlsx.File
}

var (
	handleMap          = make(map[uint64]XlsxFile)
	handleMutex        = &sync.Mutex{}
	nextHandle  uint64 = 1
)

func (xl XlsxFile) SheetCount() int {
	return len(xl.file.Sheets)
}

func (xl *XlsxFile) GetSheet(sheetIndex int) (*xlsx.Sheet, error) {
	if sheetIndex < 0 || sheetIndex >= xl.SheetCount() {
		return nil, fmt.Errorf("sheet index out of range")
	}
	return xl.file.Sheets[sheetIndex], nil
}

func (xl XlsxFile) GetSheetName(sheetIndex int) (string, error) {
	if sheetIndex < 0 || sheetIndex >= xl.SheetCount() {
		return "", fmt.Errorf("sheet index out of range")
	}
	return xl.file.Sheets[sheetIndex].Name, nil
}

func (xl XlsxFile) GetCellString(sheetIndex, row, col int) (string, error) {
	sheet, err := xl.GetSheet(sheetIndex)
	if err != nil {
		return "", err
	}
	if row < 0 || row >= len(sheet.Rows) {
		return "", fmt.Errorf("row index out of range")
	}
	if col < 0 || col >= len(sheet.Rows[row].Cells) {
		return "", fmt.Errorf("column index out of range")
	}
	return sheet.Rows[row].Cells[col].String(), nil
}

func (xl XlsxFile) RowCount(sheetIndex int) (int, error) {
	sheet, err := xl.GetSheet(sheetIndex)
	if err != nil {
		return 0, err
	}
	return len(sheet.Rows), nil
}

func (xl XlsxFile) ColCount(sheetIndex, row int) (int, error) {
	sheet, err := xl.GetSheet(sheetIndex)
	if err != nil {
		return 0, err
	}
	if row < 0 || row >= len(sheet.Rows) {
		return 0, fmt.Errorf("row index out of range")
	}
	return len(sheet.Rows[row].Cells), nil
}
