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

typedef uintptr_t(*xlsx_load_file_ptr)(const char* path);
typedef void(*xlsx_close_file_ptr)(uintptr_t handle);
typedef int(*xlsx_get_sheet_count_ptr)(uintptr_t handle);
typedef char* (*xlsx_get_sheet_name_ptr)(uintptr_t handle, int sheetIndex);
typedef void(*xlsx_free_string_ptr)(char* str);
typedef int(*xlsx_get_row_count_ptr)(uintptr_t handle, int sheetIndex);
typedef int(*xlsx_get_col_count_ptr)(uintptr_t handle, int sheetIndex, int row);
typedef char* (*xlsx_get_cell_value_ptr)(uintptr_t handle, int sheetIndex, int row, int col);

typedef uintptr_t(*xlsx_create_file_ptr)();
typedef  int(*xlsx_add_sheet_ptr)(uintptr_t handle, const char* sheetName);
typedef void(*xlsx_set_cell_value_ptr)(uintptr_t handle, int sheetIndex, int row, int col, const char* value);
typedef int(*xlsx_save_file_ptr)(uintptr_t handle, const char* filename);


#define LIBRARY_NAME "xlsx.dll"
static HMODULE loadDll() {
	HMODULE hDll = LoadLibrary(LIBRARY_NAME);
	if (!hDll) {
		printf("Failed to LoadLibrary('%s')\n", LIBRARY_NAME);
	}
	return hDll;
}

static FARPROC __stdcall getProcAddr(HMODULE hDll, LPCSTR procName) {
	auto* proc = ::GetProcAddress(hDll, procName);
	if (!proc) {
		printf("Failed to GetProcAddress('%s')\n", procName);
	}

	return proc;
}

std::string mbstr_to_utf8(const char* pszText) {
	std::wstring sUnicode = CA2W(pszText).m_psz;
	std::string sUtf8 = CW2A(sUnicode.c_str(), CP_UTF8).m_psz;
	return sUtf8;
}

std::string utf8_to_mbstr(const char* pszUtf8) {
	std::wstring sUnicode = CA2W(pszUtf8, CP_UTF8).m_psz;
	std::string sMbstr = CW2A(sUnicode.c_str()).m_psz;
	return sMbstr;
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
	HMODULE hDll = loadDll();
	if (!hDll) {
		return;
	}

	csv2xlsx_ptr csv2xlsx = (csv2xlsx_ptr)getProcAddr(hDll, procName);
	if (csv2xlsx) {
		std::string sOptions = mbstr_to_utf8("-out=ansi.sample.xlsx\n-delimiter=\\t\n-sheet-name=ansi sample");
		std::string sCsv = mbstr_to_utf8(sAnsiCsv.c_str());
		result = csv2xlsx(sCsv.c_str(), sOptions.c_str());
	}

	//do not call FreeLibrary function. 
	// FreeLibrary(hDll);

	printf("엑셀파일 생성결과 : %d\n", result);
}

void test_csv2xlsx_with_style() {
	int result = -1;
	std::string sAnsiCsv = readfile("ansi.sample.csv");
	std::string sAnsiStyleCsv = readfile("ansi.sample_style.csv");
	
	LPCSTR procName = "csv2xlsx_with_style";
	HMODULE hDll = loadDll();
	if (!hDll) {
		return;
	}

	csv2xlsx_with_style_ptr csv2xlsx_with_style = (csv2xlsx_with_style_ptr)getProcAddr(hDll, procName);
	if (csv2xlsx_with_style) {
		std::string sOptions = mbstr_to_utf8("-out=ansi.sampleWithStyle.xlsx\n-delimiter=\\t\n-sheet-name=ansi sample");
		std::string sCsv = mbstr_to_utf8(sAnsiCsv.c_str());
		std::string sStyle = mbstr_to_utf8(sAnsiStyleCsv.c_str());
		result = csv2xlsx_with_style(sCsv.c_str(), sOptions.c_str(), sStyle.c_str());
	}

	//do not call FreeLibrary function. 
	// FreeLibrary(hDll);

	printf("엑셀파일 생성결과 : %d\n", result);
}

void test_xlsx_read() {
	HMODULE hDll = loadDll();
	if (!hDll) {
		return;
	}

	LPCSTR procName = "xlsx_free_string";
	xlsx_free_string_ptr xlsx_free_string = (xlsx_free_string_ptr)getProcAddr(hDll, procName);
	if (!xlsx_free_string) {
		return;
	}

	procName = "xlsx_get_row_count";
	xlsx_get_row_count_ptr xlsx_get_row_count = (xlsx_get_row_count_ptr)getProcAddr(hDll, procName);
	if (!xlsx_get_row_count) {
		return;
	}

	procName = "xlsx_get_col_count";
	xlsx_get_col_count_ptr xlsx_get_col_count = (xlsx_get_col_count_ptr)getProcAddr(hDll, procName);
	if (!xlsx_get_col_count) {
		return;
	}

	procName = "xlsx_load_file";
	xlsx_load_file_ptr xlsx_load_file = (xlsx_load_file_ptr)getProcAddr(hDll, procName);
	if (!xlsx_load_file) {
		return;
	}

	procName = "xlsx_get_cell_value";
	xlsx_get_cell_value_ptr xlsx_get_cell_value = (xlsx_get_cell_value_ptr)getProcAddr(hDll, procName);
	if (!xlsx_get_cell_value) {
		return;
	}
		

	const uintptr_t h = xlsx_load_file("ansi.sample.xlsx");
	if (h == 0) {
		printf("Failed to xlsx_load_file()\n");
	}

	int sheetCount = 0;
	procName = "xlsx_get_sheet_count";
	xlsx_get_sheet_count_ptr xlsx_get_sheet_count = (xlsx_get_sheet_count_ptr)getProcAddr(hDll, procName);
	if (xlsx_get_sheet_count) {
		 sheetCount = xlsx_get_sheet_count(h);
		printf("sheet count = %d\n", sheetCount);
	}

	procName = "xlsx_get_sheet_name";
	xlsx_get_sheet_name_ptr xlsx_get_sheet_name = (xlsx_get_sheet_name_ptr)getProcAddr(hDll, procName);
	
	std::string name, cellValue;
	int rowCount, colCount = 0;
	for (int i = 0; i < sheetCount; i++) {
		rowCount = 0;
		colCount = 0;

		if (xlsx_get_sheet_name) {

			char* u8Name = xlsx_get_sheet_name(h, i);
			name = utf8_to_mbstr(u8Name);
			printf("[%d] sheet name = '%s'\n", i, name.c_str());

			if (u8Name) {
				xlsx_free_string(u8Name);
			}

			rowCount = xlsx_get_row_count(h, i);
			printf("rowCount=%d\n", rowCount);			
			for (int row = 0; row < rowCount; row++) {
				//printf("row[%d]\n", row);

				colCount = xlsx_get_col_count(h, i, row);
				
				for (int col = 0; col < colCount; col++) {
					if (char* u8Val = xlsx_get_cell_value(h, i, row, col)) {
						cellValue = utf8_to_mbstr(u8Val);
						xlsx_free_string(u8Val);

						printf("%s", cellValue.c_str());

						if (col < colCount - 1) {
							printf(", ");
						}
					}
				}
				printf("\n");
			}
		}
	}

	procName = "xlsx_close_file";
	xlsx_close_file_ptr xlsx_close_file = (xlsx_close_file_ptr)getProcAddr(hDll, procName);
	if (xlsx_close_file) {
		xlsx_close_file(h);
	}
}

void test_xlsx_write() {
	HMODULE hDll = loadDll();
	if (!hDll) {
		return;
	}

	LPCSTR procName = "xlsx_create_file";
	xlsx_create_file_ptr xlsx_create_file = (xlsx_create_file_ptr)getProcAddr(hDll, procName);
	if (!xlsx_create_file) {
		return;
	}

	procName = "xlsx_add_sheet";
	xlsx_add_sheet_ptr xlsx_add_sheet = (xlsx_add_sheet_ptr)getProcAddr(hDll, procName);
	if (!xlsx_add_sheet) {
		return;
	}

	procName = "xlsx_set_cell_value";
	xlsx_set_cell_value_ptr xlsx_set_cell_value = (xlsx_set_cell_value_ptr)getProcAddr(hDll, procName);
	if (!xlsx_set_cell_value) {
		return;
	}

	procName = "xlsx_save_file";
	xlsx_save_file_ptr xlsx_save_file = (xlsx_save_file_ptr)getProcAddr(hDll, procName);
	if (!xlsx_save_file) {
		return;
	}

	procName = "xlsx_close_file";
	xlsx_close_file_ptr xlsx_close_file = (xlsx_close_file_ptr)getProcAddr(hDll, procName);
	if (!xlsx_close_file) {
		return;
	}

	uintptr_t h = xlsx_create_file();
	if (!h) {
		printf("Failed to xlsx_create_file()\n");
		return;
	}

	std::string u8str = mbstr_to_utf8("test_sheet");

	int sheetIndex = xlsx_add_sheet(h, u8str.c_str());
	if (sheetIndex >= 0) {
		const int rowCount = 10;
		const int colCount = 5;
		char buf[256];
		for (int row = 0; row < rowCount; row++) {
			for (int col = 0; col < colCount; col++) {
				sprintf_s(buf, "value[%d][%d]", row, col);
				u8str = mbstr_to_utf8(buf);
				xlsx_set_cell_value(h, sheetIndex, row, col, u8str.c_str());
			}
		}
	}

	if (xlsx_save_file(h, "write_test.xlsx") == 0) {
		printf("saved\n");
	}
	else {
		printf("Failed to xlsx_save_file()\n");
	}

	xlsx_close_file(h);
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

	test_xlsx_read();

	test_xlsx_write();
	return 0;
}
