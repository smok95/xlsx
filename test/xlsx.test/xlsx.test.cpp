// xlsx.test.cpp : 이 파일에는 'main' 함수가 포함됩니다. 거기서 프로그램 실행이 시작되고 종료됩니다.
//
#include <Windows.h>
#include <stdio.h>
#include <atlstr.h>
#include <string>
#include <iostream>
#include <fstream>

typedef int (*fnCsv2xlsx)(const char* pszText, const char* pszOutFile);

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

void test() {
	int result = -1;
	std::string sSrc = readfile("ansi.sample.csv");

	if (HMODULE hDll = LoadLibrary("xlsx.dll")) {

		fnCsv2xlsx text2xlsx = (fnCsv2xlsx)GetProcAddress(hDll, "csv2xlsx");
		if (text2xlsx) {
			std::string sUtf8 = mbstr2utf8("-out=ansi.sample.xlsx\n-delimiter=\\t\n-sheet-name=ansi sample");
			std::string sText = mbstr2utf8(sSrc.c_str());
			result = text2xlsx(sText.c_str(), sUtf8.c_str());
		}
		else
			puts("Failed to GetProcAddress('csv2xlsx')");

		FreeLibrary(hDll);
	}
	else {
		puts("Failed to LoadLibrary('xlsx.dll')");
	}

	printf("엑셀파일 생성결과 : %d\n", result);
}

int main()
{
	test();
	return 0;
}
