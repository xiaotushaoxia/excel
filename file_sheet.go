package excel

import (
	"encoding/json"

	"github.com/360EntSecGroup-Skylar/excelize"
)

type FileSheet struct {
	*excelize.File
	SheetName string
}

func NewFileSheet(f *excelize.File, sheet string) *FileSheet {
	_, ok := f.Sheet[sheet]
	if !ok {
		f.NewSheet(sheet)
	}
	return &FileSheet{
		File:      f,
		SheetName: sheet,
	}
}

func (e *FileSheet) SetCellValue(axis string, value interface{}) {
	e.File.SetCellValue(e.SheetName, axis, value)
}

func (e *FileSheet) SetColWidth(startcol, endcol string, width float64) {
	e.File.SetColWidth(e.SheetName, startcol, endcol, width)
}

func (e *FileSheet) SetCellValueAndMergeCell(v, hcell, vcell string) {
	e.File.MergeCell(e.SheetName, hcell, vcell)
	e.SetCellValue(hcell, v)
}

func (e *FileSheet) MergeCell(hcell, vcell string) {
	e.File.MergeCell(e.SheetName, hcell, vcell)
}

func (e *FileSheet) SetCellStyle(hcell, vcell string, style *FormatStyle) error {
	i, err := e.NewStyle(style)
	if err != nil {
		return err
	}
	e.File.SetCellStyle(e.SheetName, hcell, vcell, i)
	return nil
}

func (e *FileSheet) SetCellStyleByID(hcell, vcell string, styleID int) {
	e.File.SetCellStyle(e.SheetName, hcell, vcell, styleID)
}

func (e *FileSheet) NewStyle(style *FormatStyle) (int, error) {
	marshal, err := json.Marshal(style)
	if err != nil {
		return 0, err
	}
	return e.File.NewStyle(string(marshal))
}
