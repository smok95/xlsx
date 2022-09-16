# 개발환경 구축 방법

- go1.19.1 windows/amd64 [https://go.dev/dl/go1.19.1.windows-amd64.msi]
- tdm64-gcc-10.3.0-2.exe [https://jmeubank.github.io/tdm-gcc/articles/2021-05/10.3.0-release]


## 32bit 빌드

go env 로 아래와 같이 설정되어 있는지확인
cmd.exe에서 할것 (powersehll에서는 적용이 안됨)
set GOARCH=386
set CGO_ENABLED=1
set CC=mingw32-gcc

// 빌드 방법
- Reduce complied file size
    -ldflags "-s -w"
* dll
    go build -ldflags "-s -w" -buildmode=c-shared -o xlsx.dll main.go
* exe
    go build -ldflags "-s -w" -o xlsx.exe main.go


## 64bit 빌드

go env 로 아래와 같이 설정되어 있는지확인

set GOARCH=amd64
set CGO_ENABLED=1
set CC=x86_64-w64-mingw32-gcc

// 빌드 방법
- Reduce complied file size
    -ldflags "-s -w"
* dll
    go build -ldflags "-s -w" -buildmode=c-shared -o xlsx.dll main.go
* exe
    go build -ldflags "-s -w" -o xlsx.exe main.go
