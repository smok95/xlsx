package main

import (
	"C"
	"fmt"
	"strings"
	"encoding/csv"
	"github.com/tealeg/xlsx"
)

func main() {	
}

//export text2xlsx
func text2xlsx(pszText *C.char, pszOutfile *C.char) C.long {
	
	text := C.GoString(pszText);
	outfile := C.GoString(pszOutfile);
	fmt.Println("filename=" + outfile);
	xlsxFile := xlsx.NewFile();
	sheet, err := xlsxFile.AddSheet("hello");
	if err != nil {
		//return err
		return -1
	}

	_ = sheet

	delimiter := "\t"
	r := csv.NewReader(strings.NewReader(text))
	r.Comma = rune(delimiter[0])

	fields, err := r.Read()
	for err==nil {
		row := sheet.AddRow()
		for _, field := range fields {
			cell := row.AddCell()
			cell.Value = field
		}
		fields, err = r.Read()
	}

	xlsxFile.Save(outfile)
	return 11
}

/*
go env 로 아래와 같이 설정되어 있는지확인
set GOARCH=386
set CGO_ENABLED=1
// 빌드 방법
go build -buildmode=c-shared -o xlsx.dll xlsx.go
*/