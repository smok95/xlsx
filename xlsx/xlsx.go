package xlsx

import (
	"bufio"
	"encoding/csv"
	"fmt"
	"io"
	"strconv"
	"strings"
	"time"

	"github.com/tealeg/xlsx"
)

// Options csv to xlsx 변환 옵션값
type Options struct {
	SkipEmptyLine bool   // 빈줄 무시
	SheetName     string // 시트명
	Outfile       string // 출력파일명
	FontName      string // 폰트명
	FontSize      int    // 폰트크기
	Delimiter     string // 구분기호
}

// ConvertFromCSV csv데이터 xlsx파일로 변환
func ConvertFromCSV(text string, opt Options, stylesText string) int {

	// set default font
	// 설정된 값이 없으면, 맑은고딕 11로 설정
	// 폰트명 정보는 https://en.wikipedia.org/wiki/List_of_typefaces_included_with_Microsoft_Windows
	if opt.FontSize <= 0 {
		opt.FontSize = 11
	}
	if len(opt.FontName) == 0 {
		opt.FontName = "맑은 고딕" // "Malgun Gothic" 여기에서 설정한 폰트명이 그대로 엑셀에 표시됨
	}

	xlsx.SetDefaultFont(opt.FontSize, opt.FontName)

	xlsxFile := xlsx.NewFile()
	sheet, err := xlsxFile.AddSheet(opt.SheetName)
	if err != nil {
		//return err
		return -1
	}

	_ = sheet

	delimiter := ","
	delimLen := len(opt.Delimiter)
	if delimLen > 0 {
		if delimLen > 1 {
			switch opt.Delimiter {
			case "\\t", "tab":
				delimiter = "\t"
			}
		} else {
			delimiter = opt.Delimiter
		}
	}

	if !opt.SkipEmptyLine {
		text = emptyLine2emptyColumn(text)
		stylesText = emptyLine2emptyColumn(stylesText)
	}

	// 출력파일명이 없는 경우 out.xlsx로 저장
	if len(opt.Outfile) == 0 {
		opt.Outfile = "out.xlsx"
	}

	r := csv.NewReader(strings.NewReader(text))
	r.Comma = rune(delimiter[0])
	r.LazyQuotes = true

	// cell style csv
	rStyle := csv.NewReader(strings.NewReader(stylesText))
	rStyle.Comma = rune('\t')
	rStyle.LazyQuotes = true
	rStyle.FieldsPerRecord = -1

	// 레코드별 필드갯수 체크기능 사용안함.
	r.FieldsPerRecord = -1
	for {
		fields, err := r.Read()
		styles, _ := rStyle.Read()
		if err == io.EOF {
			break
		}

		//fmt.Println("fields len:", len(fields), fields)

		if err != nil {
			fmt.Println("csv reader Read error:", err)
			break
		}

		styleCnt := len(styles)

		row := sheet.AddRow()
		for idx, field := range fields {
			cell := row.AddCell()

			setCellValue(cell, field)

			if styleCnt > idx {
				setCellStyle(cell, styles[idx])
			}
		}
	}

	err = xlsxFile.Save(opt.Outfile)

	if err != nil {
		fmt.Println("Failed to save file : ", err)
		return -1
	}
	return 0
}

func emptyLine2emptyColumn(v string) string {
	var lines string = ""
	scanner := bufio.NewScanner(strings.NewReader(v))
	for scanner.Scan() {
		line := scanner.Text()

		if len(line) == 0 {
			lines += "\t"
		} else {
			lines += line
		}
		lines += "\n"
	}

	return lines
}

func isDigit(c rune) bool {
	if c >= '0' && c <= '9' {
		return true
	}
	return false
}

// 숫자문자열값 float값으로로 셀에 설정
func setFloat(cell *xlsx.Cell, v string) bool {
	ret, err := strconv.ParseFloat(v, 64)
	if err != nil {
		return false
	}

	cell.SetFloat(ret)
	return true
}

// 숫자문자열 int64값으로 셀에 설정
func setInt(cell *xlsx.Cell, v string) bool {
	ret, err := strconv.ParseInt(v, 10, 64)
	if err != nil {
		return false
	}

	cell.SetInt64(ret)
	return true
}

func setPercent(cell *xlsx.Cell, v string) bool {
	ret, err := strconv.ParseFloat(v, 64)
	if err != nil {
		return false
	}

	cell.SetFloat(ret / 100) // 변환과정에서 %형식은 xlsx소스상에서 100을 곱하고 있음. 따라서 100으로 미리 나눈다.
	return true
}

type numberInfo struct {
	value      string // 숫자값
	hasComma   bool   // comma포함 여부
	isFloat    bool   // true면 실수, false면 정수
	hasPercent bool   // % 기호 포함여부
	format     string // 문자열 표시 형식
}

// 숫자 문자열인지 확인
// return
//
//	numberInfo	숫자해석 정보
//	문자열이 숫자면 true, 아니면 false
func analyzeNumberString(v string) (numberInfo, bool) {
	sLen := len(v)

	isNumber := false
	var ret numberInfo

	digitCnt := 0  // 숫자개수
	commaCnt := 0  // comma개수
	pointIdx := -1 // 소수점 위치

	commaPad := 0
	neg := false // 음수여부

	if sLen == 0 {
		return ret, isNumber
	}

	ch := v[0]      // 맨앞 1글자
	neg = ch == '-' // 음수여부

	// 문자열에서 +,-부호 제거
	if neg || ch == '+' {
		v = v[1:]

		// 좌우공백 제거
		v = strings.TrimSpace(v)
		sLen = len(v)
		if sLen == 0 {
			return ret, false
		}
	}

	// 문자열에 % 포함여부 확인
	if v[sLen-1] == '%' {
		v = v[0 : sLen-1]
		// 좌우공백 제거
		v = strings.TrimSpace(v)
		sLen = len(v)
		ret.hasPercent = true
	}

	if neg {
		ret.value += "-"
	}

	for idx, ch := range v {
		if isDigit(ch) { // 숫자
			digitCnt++
			ret.value += string(ch)
		} else if ch == ',' { // comma
			commaCnt++

			if commaCnt == 1 {
				// 1번째 콤마는 앞자리 숫자가 3개 이하일 수 있음
				if digitCnt < 3 {
					commaPad = 3 - digitCnt
				}
			}

			commaIdx := commaPad + idx

			if (commaIdx+1)%4 != 0 {
				// 콤마 위치가 4의 배수(4, 8, 12, 16 ... )가 아니면 1000단위 콤마가 아님.
				goto RETURN
			} else if idx == sLen-1 {
				// 콤마가 문자열의 맨마지막에 위치할 수 없음
				goto RETURN
			}

			if pointIdx >= 0 && idx > pointIdx {
				goto RETURN // 소수점 이후에는 콤마가 올 수 없음
			}

		} else if ch == '.' { // 소수점
			if pointIdx == -1 {
				pointIdx = idx
				ret.value += string(ch)
			} else {
				goto RETURN // 소수점이 2개 이상인 경우
			}
		} else { // 숫자, 콤마, 소수점이 아닌 경우
			goto RETURN
		}

	}

	if commaCnt > 0 {
		ret.hasComma = true
	}

	if pointIdx >= 0 {
		ret.isFloat = true
	}

	if ret.hasPercent {
		if ret.isFloat {
			ret.format = "0.00%"
		} else {
			ret.format = "0%"
		}
	} else {
		if ret.hasComma {
			if ret.isFloat {
				ret.format = "#,##0.00"
			} else {
				ret.format = "#,##0"
			}
		}
	}

	isNumber = true
RETURN:
	return ret, isNumber
}

// 문자열값 숫자타입으로 셀에 지정
// return : 문자열값이 숫자타입이 맞으면 true, 숫자타입이 아니면 false
func setNumberCellValue(cell *xlsx.Cell, v string) bool {
	num, ok := analyzeNumberString(v)

	if !ok {
		return false
	}

	isNumber := false

	if num.hasPercent {
		isNumber = setPercent(cell, num.value)
	} else {
		if num.isFloat {
			isNumber = setFloat(cell, num.value)
		} else {
			isNumber = setInt(cell, num.value)
			if !isNumber {
				isNumber = setFloat(cell, num.value)
			}
		}
	}

	if isNumber && len(num.format) > 0 {
		cell.SetFormat(num.format)
	}

	return isNumber
}

type timeInfo struct {
	hour       int
	minute     int
	second     int
	cellFormat string // 셀 표시형식
}

// 문자열에서 시간(날짜제외)값으로 변환
func parseTime(v string) (timeInfo, bool) {
	var ret timeInfo
	h, m, s := 0, 0, 0
	arr := strings.Split(v, ":")
	cnt := len(arr)
	if cnt < 2 || cnt > 3 {
		return ret, false
	}
	var err error

	// 시
	h, err = strconv.Atoi(arr[0])
	if err != nil {
		return ret, false
	}

	// 분
	m, err = strconv.Atoi(arr[1])
	if err != nil {
		return ret, false
	}

	// 초
	if cnt == 3 {
		s, err = strconv.Atoi(arr[2])
		if err != nil {
			return ret, false
		}
	}

	// 엑셀 csv변환 테스트결과 시간값이 9999를 초과하면 문자열값으로 처리함.
	if h < 0 || h > 9999 ||
		m < 0 || m > 59 ||
		s < 0 || s > 59 {
		return ret, false
	}

	if cnt == 3 { // 시,분,초값이 다 있을때
		ret.cellFormat = "h:mm:ss"
	} else { // 시분만 있을때
		// 엑셀에서 확인한결과 시분만 있더라도 시간 값이 24이상이면 [h]:mm:ss형식으로 표시됨.
		if h >= 24 {
			ret.cellFormat = "[h]:mm:ss"
		} else {
			ret.cellFormat = "h:mm"
		}
	}

	ret.hour = h
	ret.minute = m
	ret.second = s
	return ret, true
}

func isValidMonthDay(month int, day int) bool {
	return month > 0 && month <= 12 && day > 0 && day <= 31
}

// parseDate 문자열에서 날짜(시간제외)값으로 변환
func parseDate(v string) (time.Time, bool) {
	y, m, d := -1, 0, 0
	var ret time.Time

	v = strings.ReplaceAll(v, "/", "-")
	arr := strings.Split(v, "-")

	cnt := len(arr)
	if cnt < 2 || cnt > 3 {
		return ret, false
	}

	iYear, iMon, iDay := 0, 1, 2 // 배열에서 년월일의 인덱스값

	// 엑셀에서 숫자개수가 2개면 월일, 3개면 년월일로 처리하고 있음
	if cnt == 2 {
		iYear = -1
		iMon = 0
		iDay = 1
	}

	var err error
	// 년
	if iYear >= 0 {
		y, err = strconv.Atoi(arr[iYear])
		if err != nil {
			return ret, false
		}

		// 엑셀에서 확인한 결과
		// 년은 0~9999값까지 유효
		if y < 0 || y > 9999 {
			return ret, false
		}

		// 유효범위 숫자라도 문자열길이가 3자리 또는 5자리 이상인 경우 년도로 인식안함. (엑셀확인)
		slen := len(arr[iYear])
		if slen == 3 || slen > 4 {
			return ret, false
		}
	}

	// 월 (문자열 2자리 이하의 숫자)
	if len(arr[iMon]) > 2 {
		return ret, false
	}

	m, err = strconv.Atoi(arr[iMon])
	if err != nil {
		return ret, false
	}

	// 일 (문자열 2자리 이하의 숫자)
	if len(arr[iDay]) > 2 {
		return ret, false
	}

	d, err = strconv.Atoi(arr[iDay])
	if err != nil {
		return ret, false
	}

	// 연도가 없으면 올해로 설정
	if iYear == -1 {
		y = time.Now().Year()
	}

	if !isValidMonthDay(m, d) {
		// 월,일을 바꿔서 다시 검사 (단, 년도값이 없는 경우에만)
		if iYear == -1 {
			temp := m
			m = d
			d = temp
			if !isValidMonthDay(m, d) {
				return ret, false
			}
		} else {
			return ret, false
		}
	}

	// 엑셀확인결과 년도값이 100이하일때는 다음과 같이 처리됨.
	// 0 ~ 29 : 2000년대
	// 30 ~ 99 : 1900년대
	if y >= 0 && y <= 29 {
		y += 2000
	} else if y >= 30 && y <= 99 {
		y += 1900
	}

	ret = time.Date(y, time.Month(m), d, 0, 0, 0, 0, time.UTC)
	return ret, true
}

// 날짜시간 문자열인지 확인
// return
// 	문자열이 날짜시간이면 true

// setDateTimeCellValue 문자열값을 날짜(시간)값으로 셀에 지정
// return : 문자열값이 날짜타입이 맞으면 true, 아니면 false
func setDateTimeCellValue(cell *xlsx.Cell, v string) bool {
	var dt time.Time
	var format string
	var isDate = false
	//var year, month, day, hour, min, second int

	/*
		csv날짜시간값 자동변환시 사용되는 기본포맷
		아마도 국가 또는 언어설정에 따라 달라질 수 있으나,
		현재는 대한민국 기준으로 작성되었음.
		시간 => "h:mm:ss"
		날짜 => "yyyy-MM-dd"
		날짜시간 => "yyyy-MM-dd h:mm"
	*/
	// 좌우 공백 제거
	v = strings.TrimSpace(v)
	arr := strings.Split(v, " ")

	var ti timeInfo
	var d time.Time
	hasTime, hasDate := false, false
	parseOk := false

	for i := 0; i < len(arr); i++ {

		colonCnt := strings.Count(arr[i], ":")
		if colonCnt > 0 && colonCnt <= 2 && !hasTime {
			hasTime = true
			ti, parseOk = parseTime(strings.TrimSpace(arr[i]))
			if !parseOk {
				goto RETURN
			}
		} else if strings.ContainsAny(arr[i], "/-") && !hasDate {
			hasDate = true
			d, parseOk = parseDate(strings.TrimSpace(arr[i]))
			if !parseOk {
				goto RETURN
			}
		} else {
			goto RETURN
		}
	}

	if hasTime && hasDate {
		dt = time.Date(d.Year(), d.Month(), d.Day(), ti.hour, ti.minute, ti.second, 0, time.UTC)
		format = "yyyy-mm-dd h:mm"

		cell.SetDateTime(dt)
		cell.SetFormat(format)
	} else if hasDate {
		dt = d
		format = "yyyy-mm-dd"
		cell.SetDateTime(dt)
		cell.SetFormat(format)
	} else if hasTime {
		// 엑셀기준시간은 1904.1.1일로 관련 정보는 xlsx/date.go excel1900Epoc, excel1904Epoc 참고
		dt = time.Date(1904, 1, 1, ti.hour, ti.minute, ti.second, 0, time.UTC)
		format = ti.cellFormat

		cell.SetDateTimeWithFormat(xlsx.TimeToExcelTime(dt.UTC(), true), format)
	}
	isDate = true

RETURN:
	return isDate
}

// setCellValue set data in correct format.
func setCellValue(cell *xlsx.Cell, v string) {

	// 좌우 공백제거
	temp := strings.TrimSpace(v)

	// 첫 문자가 '='(equal sign)이며 최소 2글자 이상이면 공식(Formulas)으로 저장
	if len(temp) > 1 && temp[0] == '=' {
		cell.SetStringFormula(temp)
		return
	}

	// 숫자면 숫자포맷에 맞춰 셀에 저장
	if setNumberCellValue(cell, temp) {
		return
	}

	// 날짜시간이면 해당 포맷에 맞춰 셀에 저장
	if setDateTimeCellValue(cell, temp) {
		return
	}

	// 나머지는 문자열 그대로 셀에 저장
	cell.SetString(v)
}

// setCellStyle cell 스타일 설정
func setCellStyle(cell *xlsx.Cell, styleText string) {

	style := cell.GetStyle()

	// styleText format
	// name:value;name:value;name:value;...
	attrs := strings.Split(styleText, ";")
	for _, attr := range attrs {
		kv := strings.Split(attr, ":")
		if len(kv) != 2 {
			continue
		}

		name := strings.TrimSpace(kv[0])
		value := strings.TrimSpace(kv[1])

		// color: text color
		if strings.EqualFold("color", name) {
			style.Font.Color = value
			style.ApplyFont = true
		} else if strings.EqualFold("background-color", name) {
			style.Fill.FgColor = value
			style.Fill.PatternType = "solid"
			style.ApplyFill = true
		}
	}
}
