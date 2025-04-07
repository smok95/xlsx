package main

/*
#include <stdlib.h>
*/
import "C"

import (
	"flag"
	"fmt"
	"os"
	"strings"
	"unsafe"

	xl "github.com/smok95/xlsx/xlsx"
)

var (
	version string = "0.3"
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

	buf, err := os.ReadFile(csvFile)
	var csv string
	if err != nil {
		fmt.Printf(err.Error())
		return
	}

	csv = string(buf)
	xl.ConvertFromCSV(csv, opt, "")
}

//export csv2xlsx
func csv2xlsx(pszCsv *C.char, pszOptions *C.char) C.long {
	return csv2xlsx_with_style(pszCsv, pszOptions, nil)
}

//export csv2xlsx_with_style
func csv2xlsx_with_style(pszCsv *C.char, pszOptions *C.char, pszStyles *C.char) C.long {
	var opt xl.Options
	text := C.GoString(pszCsv)
	cmdline := C.GoString(pszOptions)
	styles := C.GoString(pszStyles)
	fs := flag.NewFlagSet("csv2xlsxOptions", flag.ContinueOnError)
	initFlag(fs, &opt)
	args := strings.Split(cmdline, "\n")
	fs.Parse(args)
	ret := xl.ConvertFromCSV(text, opt, styles)
	return C.long(ret)
}

//export xlsx_load_file
func xlsx_load_file(path *C.char) C.uintptr_t {
	handle, err := xl.LoadXlsx(C.GoString(path))
	if err != nil {
		fmt.Println(err.Error())
		return 0
	}
	return C.uintptr_t(handle)
}

//export xlsx_close_file
func xlsx_close_file(handle C.uintptr_t) {
	xl.FreeXlsx(uint64(handle))
}

//export xlsx_get_sheet_count
func xlsx_get_sheet_count(handle C.uintptr_t) C.int {
	xlFile, err := xl.GetXlsx(uint64(handle))
	if err != nil {
		fmt.Println(err.Error())
		return 0
	}
	return C.int(xlFile.SheetCount())
}

//export xlsx_get_sheet_name
func xlsx_get_sheet_name(handle C.uintptr_t, sheetIndex C.int) *C.char {
	xlFile, err := xl.GetXlsx(uint64(handle))
	if err != nil {
		fmt.Println(err.Error())
		return nil
	}

	name, err := xlFile.GetSheetName(int(sheetIndex))
	if err != nil {
		fmt.Println(err.Error())
		return nil

	}

	return C.CString(name)
}

//export xlsx_get_row_count
func xlsx_get_row_count(handle C.uintptr_t, sheetIndex C.int) C.int {
	xlFile, err := xl.GetXlsx(uint64(handle))
	if err != nil {
		fmt.Println(err.Error())
		return 0
	}

	count, err := xlFile.RowCount(int(sheetIndex))
	if err != nil {
		fmt.Println(err.Error())
		return 0
	}
	return C.int(count)
}

//export xlsx_get_col_count
func xlsx_get_col_count(handle C.uintptr_t, sheetIndex C.int, row C.int) C.int {
	xlFile, err := xl.GetXlsx(uint64(handle))
	if err != nil {
		fmt.Println(err.Error())
		return 0
	}

	count, err := xlFile.ColCount(int(sheetIndex), int(row))
	if err != nil {
		fmt.Println(err.Error())
		return 0
	}
	return C.int(count)
}

//export xlsx_get_cell_value
func xlsx_get_cell_value(handle C.uintptr_t, sheetIndex C.int, row C.int, col C.int) *C.char {
	xlFile, err := xl.GetXlsx(uint64(handle))
	if err != nil {
		fmt.Println(err.Error())
		return nil
	}

	value, err := xlFile.GetCellString(int(sheetIndex), int(row), int(col))
	if err != nil {
		fmt.Println(err.Error())
		return nil
	}
	return C.CString(value)
}

//export xlsx_free_string
func xlsx_free_string(str *C.char) {
	C.free(unsafe.Pointer(str))
}

//export xlsx_create_file
func xlsx_create_file() C.uintptr_t {
	handle, err := xl.CreateXlsx()
	if err != nil {
		fmt.Println(err.Error())
		return 0
	}
	return C.uintptr_t(handle)
}

//export xlsx_add_sheet
func xlsx_add_sheet(handle C.uintptr_t, sheetName *C.char) C.int {
	xlFile, err := xl.GetXlsx(uint64(handle))
	if err != nil {
		fmt.Println(err.Error())
		return -1
	}

	index, err := xlFile.AddSheet(C.GoString(sheetName))
	if err != nil {
		fmt.Println(err.Error())
		return -1
	}
	return C.int(index)
}

//export xlsx_set_cell_value
func xlsx_set_cell_value(handle C.uintptr_t, sheetIndex C.int, row C.int, col C.int, value *C.char) {
	xlFile, err := xl.GetXlsx(uint64(handle))
	if err != nil {
		fmt.Println(err.Error())
		return
	}

	err = xlFile.SetCellValue(int(sheetIndex), int(row), int(col), C.GoString(value))
	if err != nil {
		fmt.Println(err.Error())
		return
	}
}

//export xlsx_save_file
func xlsx_save_file(handle C.uintptr_t, filename *C.char) C.int {
	xlFile, err := xl.GetXlsx(uint64(handle))
	if err != nil {
		fmt.Println(err.Error())
		return -1
	}

	err = xlFile.Save(C.GoString(filename))
	if err != nil {
		fmt.Println(err.Error())
		return -1
	}

	return 0
}
