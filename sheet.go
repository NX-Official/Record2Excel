package Record2Excel

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"reflect"
)

type Sheet interface {
	AddRecord(record any) error
}

type sheet struct {
	name     string
	template template
	file     *excelize.File
	offset   int
	index    map[string]int
}

func newSheet(name string, model any, file *excelize.File) (s Sheet, err error) {
	var e = &sheet{
		name:  name,
		index: make(map[string]int),
		file:  file,
	}
	e.template, err = newTemplate(model)
	if err != nil {
		return nil, err
	}
	_, err = file.NewSheet(name)
	if err != nil {
		return nil, err
	}
	err = e.buildHeader()
	if err != nil {
		return nil, err
	}
	return e, nil
}

func (s *sheet) buildHeader() (err error) {
	s.offset = s.template.depth()

	currentColumn := 1
	var mergeRanges [][2]string // 用于记录需要合并的单元格范围

	// 定义一个递归函数，用于构建表头并记录合并单元格的范围
	var buildHeaderForRow func(node *itemNode, row int, parentColStart *int) error
	buildHeaderForRow = func(node *itemNode, row int, parentColStart *int) error {
		colStart := currentColumn // 当前节点开始的列
		if parentColStart != nil {
			colStart = *parentColStart // 如果有父节点，从父节点的列开始
		}

		// 设置单元格的值
		cell, _ := excelize.CoordinatesToCellName(currentColumn, row)
		err := s.file.SetCellValue(s.name, cell, node.tagName)
		if err != nil {
			return err
		}
		s.index[node.fieldPath] = currentColumn

		if len(node.subItems) == 0 { // 如果是叶子节点
			if s.offset > row { // 需要跨行合并
				cellEnd, _ := excelize.CoordinatesToCellName(currentColumn, s.offset)
				mergeRanges = append(mergeRanges, [2]string{cell, cellEnd})
			}
			currentColumn++ // 移动到下一个列
		} else { // 如果有子节点
			for _, child := range node.subItems {
				err := buildHeaderForRow(child, row+1, &colStart) // 递归构建子节点表头
				if err != nil {
					return err
				}
			}
			if row == 1 { // 如果是第一层嵌套，记录合并的单元格范围
				cellEnd, _ := excelize.CoordinatesToCellName(currentColumn-1, row)
				if cell != cellEnd { // 避免单列合并
					mergeRanges = append(mergeRanges, [2]string{cell, cellEnd})
				}
			}
		}
		return nil
	}

	// 从根节点开始递归
	for _, item := range s.template.items.subItems {
		err = buildHeaderForRow(item, 1, nil)
		if err != nil {
			return
		}
	}

	// 执行合并单元格操作
	for _, mergeRange := range mergeRanges {
		err = s.file.MergeCell(s.name, mergeRange[0], mergeRange[1])
		if err != nil {
			return
		}
	}
	return
}

func (s *sheet) AddRecords(records []any) error {
	for _, record := range records {
		err := s.AddRecord(record)
		if err != nil {
			return err
		}
	}
	return nil
}

func (s *sheet) AddRecord(record any) error {
	if reflect.TypeOf(record) != s.template.t {
		return fmt.Errorf("record type mismatch")
	}

	v := reflect.ValueOf(record)
	startIdx := s.offset + 1

	var insert func(v reflect.Value, node *itemNode) error
	insert = func(v reflect.Value, node *itemNode) error {
		for _, item := range node.subItems {
			val := v.FieldByName(item.name)
			switch val.Kind() {
			case reflect.Struct:
				err := insert(val, item)
				if err != nil {
					return err
				}

			case reflect.Slice:
				currentIdx := s.offset + 1
				for i := 0; i < val.Len(); i++ {
					val := val.Index(i)
					if val.Kind() == reflect.Struct {
						tmpIdx := startIdx
						startIdx = currentIdx
						err := insert(val, item)
						if err != nil {
							return err
						}
						startIdx = tmpIdx
					} else {
						cell, _ := excelize.CoordinatesToCellName(s.index[item.fieldPath], currentIdx)
						err := s.file.SetCellValue(s.name, cell, val.Interface())
						if err != nil {
							return err
						}
					}
					currentIdx++
				}
				startIdx = max(startIdx, currentIdx)

			default:
				cell, _ := excelize.CoordinatesToCellName(s.index[item.fieldPath], startIdx)
				err := s.file.SetCellValue(s.name, cell, val.Interface())
				if err != nil {
					return err
				}
			}
		}
		return nil
	}

	err := insert(v, s.template.items)
	if err != nil {
		return err
	}
	s.offset = startIdx - 1
	return nil
}
