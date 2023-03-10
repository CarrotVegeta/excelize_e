package excel

import (
	"fmt"
	"testing"
)

func TestExcel(t *testing.T) {
	f := NewFile("sheet1")
	c := f.NewCell(1)
	if err := c("host", "api", "源IP"); err != nil {
		t.Errorf(err.Error())
		return
	}
	if err := c("kdjfk", "JKdkfs", "dskfkd"); err != nil {
		t.Errorf(err.Error())
		return
	}

	err := f.SaveAs("./123.xlsx")
	if err != nil {
		t.Errorf(err.Error())
		return
	}
}
func TestConvertNumToChar(t *testing.T) {
	fmt.Println(ConvertNumToChars(10))
	fmt.Println(ConvertNumToChars(20))
	fmt.Println(ConvertNumToChars(26))
	fmt.Println(ConvertNumToChars(27))
	fmt.Println(ConvertNumToChars(53))
}
