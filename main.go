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

var (
	version string = "0.1.0"
)

// Parse 실행인자 해석
func initFlag(f *flag.FlagSet, opt *xl.Options) {

	f.StringVar(&opt.Outfile, "out", "", "Path to the xlsx output file.")
	f.BoolVar(&opt.SkipEmptyLine, "skip-empty-line", false, "엑셀파일로 변환시 빈줄은 제거하고 변환한다.")
	f.StringVar(&opt.SheetName, "sheet-name", "Sheet1", "시트명")
	f.StringVar(&opt.FontName, "font-name", "맑은 고딕", "Default font name")
	f.IntVar(&opt.FontSize, "font-size", 11, "Default font size")
	f.StringVar(&opt.Delimiter, "delimiter", ",", "csv 구분기호")
}

func main() {
	var opt xl.Options
	var csvFile string
	var showVersion bool
	flag.StringVar(&csvFile, "in", "", "Path to the csv input file.")
	flag.BoolVar(&showVersion, "version", false, "print xlsx version")

	initFlag(flag.CommandLine, &opt)
	flag.Parse()

	if showVersion {
		fmt.Println(version)
		return
	}

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
	fs := flag.NewFlagSet("csv2xlsxOptions", flag.ContinueOnError)
	initFlag(fs, &opt)
	args := strings.Split(cmdline, "\n")
	fs.Parse(args)
	ret := xl.ConvertFromCSV(text, opt)
	return C.long(ret)
}

/*
// 사용한 go version go1.14.1 windows/amd64
go env 로 아래와 같이 설정되어 있는지확인
set GOARCH=386
set CGO_ENABLED=1
// 빌드 방법
	- Reduce complied file size
		-ldflags "-s -w"
	* dll
		go build -ldflags "-s -w" -buildmode=c-shared -o xlsx.dll main.go
	* exe
		go build -ldflags "-s -w" -o xlsx.exe main.go
*/
