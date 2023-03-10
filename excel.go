package excel

import (
	"fmt"
	"github.com/xuri/excelize/v2"
)

type FileExcel struct {
	*excelize.File
	SheetName    string
	CellValueMap map[string]any
}

var Chars = []string{"", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}

func NewFile(sheetName string) *FileExcel {
	return &FileExcel{
		File:      excelize.NewFile(),
		SheetName: sheetName,
	}
}

func (f *FileExcel) NewCell(row int) func(values ...string) error {
	var setCellValue = func(values ...string) error {
		for i, v := range values {
			cell := fmt.Sprintf("%s%d", ConvertNumToChars(i+1), row)
			err := f.SetCellValue(f.SheetName, cell, v)
			if err != nil {
				return fmt.Errorf("set cell value erro :%v", err)
			}
		}
		row++
		return nil
	}
	return setCellValue
}
func ConvertNumToChars(num int) string {
	var cols string
	v := num
	for v > 0 {
		k := v % 26
		if k == 0 {
			k = 26
		}
		v = (v - k) / 26
		cols = Chars[k] + cols
	}
	return cols
}
