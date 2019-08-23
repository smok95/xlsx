// xlsx.test.cpp : 이 파일에는 'main' 함수가 포함됩니다. 거기서 프로그램 실행이 시작되고 종료됩니다.
//
#include <Windows.h>
#include <stdio.h>
#include <atlstr.h>
#include <string>

typedef int (*fnText2xlsx)(const char* pszText, const char* pszOutFile);

std::string mbstr2utf8(const char* pszText) {
	std::wstring sUnicode = CA2W(pszText).m_psz;
	std::string sUtf8 = CW2A(sUnicode.c_str(), CP_UTF8).m_psz;
	return sUtf8;
}

int main()
{
	int result = -1;
	const char* pszText = "col1\tcol2\tcol3\nvalue1\tvalue2\tvalue3";
	if (HMODULE hDll = LoadLibrary("xlsx.dll")) {

		fnText2xlsx text2xlsx = (fnText2xlsx)GetProcAddress(hDll, "text2xlsx");
		if (text2xlsx) {
			std::string sUtf8 = mbstr2utf8("안녕.xlsx");
			std::string sText = mbstr2utf8(pszText);
			result = text2xlsx(sText.c_str(), sUtf8.c_str());
		}			
		else
			puts("Failed to GetProcAddress('text2xlsx')");

		FreeLibrary(hDll);
	}
	else {
		puts("Failed to LoadLibrary('xlsx.dll')");
	}

	printf("엑셀파일 생성결과 : %d\n", result);

	return 0;
}
