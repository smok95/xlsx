set GOARCH=386
set CGO_ENABLED=1
set CC=gcc
go build -ldflags "-s -w" -buildmode=c-shared -o xlsx.dll main.go
go build -ldflags "-s -w" -o xlsx.exe main.go

set GOARCH=amd64
set CGO_ENABLED=1
set CC=x86_64-w64-mingw32-gcc

go build -ldflags "-s -w" -buildmode=c-shared -o xlsx-x64.dll main.go
go build -ldflags "-s -w" -o xlsx-x64.exe main.go
