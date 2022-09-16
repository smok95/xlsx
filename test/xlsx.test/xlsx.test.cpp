// xlsx.test.cpp : 이 파일에는 'main' 함수가 포함됩니다. 거기서 프로그램 실행이 시작되고 종료됩니다.
//
#include <Windows.h>
#include <stdio.h>
#include <atlstr.h>
#include <string>
#include <iostream>
#include <fstream>

typedef int (*csv2xlsx_ptr)(const char* pszCsv, const char* pszOptions);
typedef int (*csv2xlsx_with_style_ptr)(const char* pszCsv, const char* pszOptions, const char* pszStyles);


std::string mbstr2utf8(const char* pszText) {
	std::wstring sUnicode = CA2W(pszText).m_psz;
	std::string sUtf8 = CW2A(sUnicode.c_str(), CP_UTF8).m_psz;
	return sUtf8;
}

std::string readfile(const char* filename) {
	std::string lines;
	std::ifstream file(filename);
	if (!file.is_open())
		return lines;
	std::string line;
	while (std::getline(file, line)) {
		//if (line.length() == 0)
		//	line = " ";	// empty line인 경우 공백추가, go encoding/csv에서 empty line은 skip처리되는 문제
		lines += line + "\n";
	}		
	file.close();
	return lines;
}

void test_csv2xlsx() {
	int result = -1;
	std::string sAnsiCsv = readfile("ansi.sample.csv");

	LPCSTR procName = "csv2xlsx";

	if (HMODULE hDll = LoadLibrary("xlsx.dll")) {

		csv2xlsx_ptr csv2xlsx = (csv2xlsx_ptr)GetProcAddress(hDll, procName);
		if (csv2xlsx) {
			std::string sOptions = mbstr2utf8("-out=ansi.sample.xlsx\n-delimiter=\\t\n-sheet-name=ansi sample");
			std::string sCsv = mbstr2utf8(sAnsiCsv.c_str());
			result = csv2xlsx(sCsv.c_str(), sOptions.c_str());
		}
		else {
			printf("Failed to GetProcAddress('%s')\n", procName);
		}

		//do not call FreeLibrary function. 
		// FreeLibrary(hDll);
	}
	else {
		printf("Failed to LoadLibrary('%s')", procName);
	}

	printf("엑셀파일 생성결과 : %d\n", result);
}

void test_csv2xlsx_with_style() {
	int result = -1;
	std::string sAnsiCsv = readfile("ansi.sample.csv");
	std::string sAnsiStyleCsv = readfile("ansi.sample_style.csv");
	
	LPCSTR procName = "csv2xlsx_with_style";

	if (HMODULE hDll = LoadLibrary("xlsx.dll")) {

		csv2xlsx_with_style_ptr csv2xlsx_with_style = (csv2xlsx_with_style_ptr)GetProcAddress(hDll, procName);
		if (csv2xlsx_with_style) {
			std::string sOptions = mbstr2utf8("-out=ansi.sampleWithStyle.xlsx\n-delimiter=\\t\n-sheet-name=ansi sample");
			std::string sCsv = mbstr2utf8(sAnsiCsv.c_str());
			std::string sStyle = mbstr2utf8(sAnsiStyleCsv.c_str());
			result = csv2xlsx_with_style(sCsv.c_str(), sOptions.c_str(), sStyle.c_str());
		}
		else {
			printf("Failed to GetProcAddress('%s')\n", procName);
		}

		//do not call FreeLibrary function. 
		// FreeLibrary(hDll);
	}
	else {
		printf("Failed to LoadLibrary('%s')", procName);
	}

	printf("엑셀파일 생성결과 : %d\n", result);
}

int main()
{
	int count = 2;
	for (int i = 0; i < count; i++) {
		test_csv2xlsx();
	}		

	for (int i = 0; i < count; i++) {
		test_csv2xlsx_with_style();
	}
	return 0;
}
