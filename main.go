package main

import (
	"C"

	xl "./smok95/xlsx"
)
import (
	"flag"
	"fmt"
	"io/ioutil"
	"strings"
)

// Parse 실행인자 해석
func initFlag(opt *xl.Options) {
	flag.StringVar(&opt.Outfile, "out", "", "Path to the xlsx output file.")
	flag.BoolVar(&opt.SkipEmptyLine, "skip-empty-line", false, "엑셀파일로 변환시 빈줄은 제거하고 변환한다.")
	flag.StringVar(&opt.SheetName, "sheet-name", "Sheet1", "시트명")
	flag.StringVar(&opt.FontName, "font-name", "맑은 고딕", "Default font name")
	flag.IntVar(&opt.FontSize, "font-size", 11, "Default font size")
	flag.StringVar(&opt.Delimiter, "delimiter", ",", "csv 구분기호")
}

func main() {
	var opt xl.Options
	var csvFile string

	flag.StringVar(&csvFile, "in", "", "Path to the csv input file.")
	initFlag(&opt)

	flag.Parse()

	buf, err := ioutil.ReadFile(csvFile)
	var csv string
	if err != nil {
		fmt.Printf(err.Error())
		return
	}

	csv = string(buf)
	xl.ConvertFromCSV(csv, opt)
}

//export csv2xlsx
func csv2xlsx(pszText *C.char, pszOptions *C.char) C.long {
	var opt xl.Options
	text := C.GoString(pszText)
	cmdline := C.GoString(pszOptions)
	initFlag(&opt)
	args := strings.Split(cmdline, "\n")
	flag.CommandLine.Parse(args)
	ret := xl.ConvertFromCSV(text, opt)
	return C.long(ret)
}

/*
go env 로 아래와 같이 설정되어 있는지확인
set GOARCH=386
set CGO_ENABLED=1
// 빌드 방법
	* dll
		go build -buildmode=c-shared -o xlsx.dll main.go
	* exe
		go build -o xlsx.exe main.go
*/
