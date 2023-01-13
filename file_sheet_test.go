package excel

import (
	"fmt"
	"testing"

	"github.com/360EntSecGroup-Skylar/excelize"
)

func TestNewFileSheet(t *testing.T) {
	ef := NewFileSheet(excelize.NewFile(), "Sheet1")

	err := ef.SetCellStyle("A1", "H10", &FormatStyle{
		Border: []Border{
			{
				Type:  "left",
				Color: "000000",
				Style: 1,
			},
			{
				Type:  "top",
				Color: "000000",
				Style: 1,
			},
			{
				Type:  "bottom",
				Color: "000000",
				Style: 1,
			},
			{
				Type:  "right",
				Color: "000000",
				Style: 1,
			},
		},
		Alignment: &Alignment{
			Horizontal: "left",
			Vertical:   "top",
		},
	})
	if err != nil {
		t.Fatal(err)
	}
	// 对比
	//ss := `{"border":[{"type":"left","color":"000000","style":1},{"type":"top","color":"000000","style":1},{"type":"bottom","color":"000000","style":1},{"type":"right","color":"000000","style":1}],"alignment":{"horizontal":"left","vertical":"top"}}`
	//ii, err := ef.File.NewStyle(ss)
	//if err != nil {
	//	t.Fatal(err)
	//}
	//ef.File.SetCellStyle("Sheet1", "A1", "H10", ii)

	ef2 := NewFileSheet(ef.File, "Sheet2")
	ef2.SetCellValue("A1", "制程")
	ef2.SetCellValue("B1", "机种")
	ef2.SetCellValue("C1", "版本")
	ef2.SetCellValue("D1", "颜色")

	ef.SetCellValue("A1", "制程")
	ef.SetCellValue("B1", "机种")
	ef.SetCellValue("C1", "版本")
	ef.SetCellValue("D1", "颜色")

	days := []string{"0101", "0102", "0103"}
	for i, day := range days {
		ef.SetCellValue(ToCell(NumberToTitle(5+i), 1), day)
	}

	data := [][]any{
		{1, 2, 3, 4, 5, 6, 7},
		{1, 2, 3, 4, 5, 6, 7},
		{1, 2, 3, 4, 5, 6, 7},
		{1, 2, 3, 4, 5, 6, 7},
	}

	for r := 0; r < len(data); r++ {
		for j := 0; j < len(data[r]); j++ {
			ef.SetCellValue(ToCell(NumberToTitle(j+1), r+2), data[r][j])
		}
	}

	ef.SaveAs("E:\\cloudia-apps\\src\\goexcel\\1.xlsx")
}

func TestName(t *testing.T) {
	fmt.Println(
		excelize.ToAlphaString(27),
		excelize.TitleToNumber("AB"),

		NumberToTitle(27),
		TitleToNumber("AA"),
	)
}
