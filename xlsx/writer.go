package xlsx

import (
	"fmt"

	"github.com/tealeg/xlsx"
)

// CreateXlsx 새로은 xlsx 파일 생성
func CreateXlsx() (uint64, error) {

	xlFile := xlsx.NewFile()
	if xlFile == nil {
		return 0, fmt.Errorf("failed to create new xlsx file")
	}

	handleMutex.Lock()
	defer handleMutex.Unlock()

	h := nextHandle
	nextHandle++
	handleMap[h] = XlsxFile{file: xlFile}
	return h, nil
}

// AddSheet 시트 추가
func (xl *XlsxFile) AddSheet(sheetName string) (int, error) {
	if sheetName == "" {
		sheetName = "Sheet1"
	}

	_, err := xl.file.AddSheet(sheetName)
	if err != nil {
		return -1, err
	}

	return xl.SheetCount() - 1, nil
}

// SetCellValue 셀 값 설정
func (xl *XlsxFile) SetCellValue(sheetIndex, row, col int, value string) error {
	sheet, err := xl.GetSheet(sheetIndex)
	if err != nil {
		return err
	}

	// row가 없으면 생성
	for len(sheet.Rows) <= row {
		sheet.AddRow()
	}

	// col이 없으면 생성
	for len(sheet.Rows[row].Cells) <= col {
		sheet.Rows[row].AddCell()
	}

	// 셀 값 설정
	sheet.Rows[row].Cells[col].SetValue(value)
	return nil
}

func (xl *XlsxFile) Save(filename string) error {
	if filename == "" {
		return fmt.Errorf("file name is empty")
	}

	err := xl.file.Save(filename)
	if err != nil {
		return err
	}
	return nil
}
