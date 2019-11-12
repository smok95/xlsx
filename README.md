## Intoduction
xlsx는 csv파일을 엑셀파일로 변환시켜주는 프로그램입니다.

## Usage
exe, dll 2가지 버전이 있으며, 사용방법은 아래와 같습니다.

### xlsx.exe

    xlsx.exe -in test.csv -out test.xlsx
간단한 사용방법은 위와 같고 명령어는 다음과 같습니다.

실행인자  | 내용 | default | R/O
------------- | -------| ------ | -------
-h  | 도움말 출력 | | - 
-in  | csv파일명, 현재 utf8형식만 지원 | | Required
-out | xlsx파일명 | out.xlsx | Optional
-font-name | 기본 폰트 | 맑은 고딕 | Optional
-font-size | 폰트 크기 | 11 | Optional
-sheet-name | 시트명 | Sheet1 | Optional
-delimiter | 컬럼 구분기호 (탭의 경우 "/t" 또는 "tab"으로 입력) | , | Optional
-skip-empty-line | 해당 옵션 설정시 csv파일내의 빈줄은 무시됩니다. | | Optional


### xlsx.dll

```cpp

int main() {
  /*
    int csv2xlsx(char* csv, char* options);
    성공시 0 실패시 0 이외의 값을 리턴합니다.
    csv   : xlsx파일로 변환할 csv데이터 (utf8포맷만 지원)
    options : 변환 옵션으로 옵션값은 위 xlsx.exe와 동일합니다. 단 실행인자 사이에는 구분자(\n)를 사용해야 합니다.
    마찬가지로 utf8문자열이어야 합니다.
    
    xlsx.test 예제를 참고해주세요
  */
  std::string csv = "column1  column2 column3\nvalue1 value2  value3";  
  std::string options = "-out=test.xlsx\n-sheet-name=TestSheetName\n-delimiter=tab"; 
  int result = csv2xlsx(csv.c_str(), options.c_str()); 
  return result;
}
```

