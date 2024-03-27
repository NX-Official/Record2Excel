package Record2Excel

import (
	"fmt"
	"github.com/xuri/excelize/v2"
)

type WorkBook interface {
	AddSheet(name string, model any, options ...Option) (Sheet, error)
	UseSheet(name string) (Sheet, error)
	Export() (*excelize.File, error)
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

func (w *workBook) AddSheet(name string, model any, options ...Option) (s Sheet, err error) {
	w.sheets[name], err = newSheet(name, model, w.file, options...)
	return w.sheets[name], err
}

func (w *workBook) UseSheet(name string) (s Sheet, err error) {
	s, ok := w.sheets[name]
	if !ok {
		return nil, fmt.Errorf("sheet %s not found", name)
	}
	return s, nil
}

func (w *workBook) Export() (*excelize.File, error) {
	for _, s := range w.sheets {
		err := s.applyHeaderStyle()
		if err != nil {
			return nil, err
		}
		err = s.applyContentStyle()
		if err != nil {
			return nil, err
		}
	}

	_ = w.file.DeleteSheet("Sheet1")
	return w.file, nil
}
