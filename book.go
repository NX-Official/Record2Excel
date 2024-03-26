package Record2Excel

import (
	"fmt"
	"github.com/xuri/excelize/v2"
)

type WorkBook interface {
	AddSheet(name string, model any) (Sheet, error)
	UseSheet(name string) (Sheet, error)
	Export() *excelize.File
}

type workBook struct {
	file   *excelize.File
	sheets map[string]Sheet
}

func NewWorkBook() WorkBook {
	return &workBook{
		file:   excelize.NewFile(),
		sheets: make(map[string]Sheet),
	}
}

func (w *workBook) AddSheet(name string, model any) (s Sheet, err error) {
	w.sheets[name], err = newSheet(name, model, w.file)
	return w.sheets[name], err
}

func (w *workBook) UseSheet(name string) (s Sheet, err error) {
	s, ok := w.sheets[name]
	if !ok {
		return nil, fmt.Errorf("sheet %s not found", name)
	}
	return s, nil
}

func (w *workBook) Export() *excelize.File {
	_ = w.file.DeleteSheet("Sheet1")
	return w.file
}
