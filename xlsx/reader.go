package xlsx

import (
	"fmt"

	"github.com/tealeg/xlsx"
)

func LoadXlsx(path string) (uint64, error) {
	if path == "" {
		return 0, fmt.Errorf("file path is empty")
	}

	xlFile, err := xlsx.OpenFile(path)
	if err != nil {
		return 0, err
	}

	handleMutex.Lock()
	defer handleMutex.Unlock()

	h := nextHandle
	nextHandle++
	handleMap[h] = XlsxFile{file: xlFile}
	return h, nil
}

func FreeXlsx(handle uint64) error {
	handleMutex.Lock()
	defer handleMutex.Unlock()

	if _, ok := handleMap[handle]; ok {
		delete(handleMap, handle)
		return nil
	}
	return fmt.Errorf("invalid handle")
}

func GetXlsx(handle uint64) (*XlsxFile, error) {
	handleMutex.Lock()
	defer handleMutex.Unlock()

	if xlFile, ok := handleMap[handle]; ok {
		return &xlFile, nil
	}
	return nil, fmt.Errorf("invalid handle")
}
