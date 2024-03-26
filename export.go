package Record2Excel

import (
	"github.com/xuri/excelize/v2"
	"reflect"
)

type Exporter[T any] interface {
	AddRecord(record T) error
	Export() *excelize.File
}

type exporter[T any] struct {
	template template[T]
	file     *excelize.File
	offset   int
	index    map[string]int
}

func NewExporter[T any](model T) (Exporter Exporter[T], err error) {
	var e = &exporter[T]{
		file:  excelize.NewFile(),
		index: make(map[string]int),
	}
	e.template, err = newTemplate(model)
	if err != nil {
		return nil, err
	}
	e.buildHeader()
	return e, nil
}

func (e *exporter[T]) buildHeader() {
	e.offset = e.template.items.depth() - 1

	// 初始化Excel文件，如果还没有初始化
	if e.file == nil {
		e.file = excelize.NewFile()
	}

	currentColumn := 1
	var mergeRanges [][2]string // 用于记录需要合并的单元格范围

	// 定义一个递归函数，用于构建表头并记录合并单元格的范围
	var buildHeaderForRow func(node *itemNode, row int, parentColStart *int)
	buildHeaderForRow = func(node *itemNode, row int, parentColStart *int) {
		colStart := currentColumn // 当前节点开始的列
		if parentColStart != nil {
			colStart = *parentColStart // 如果有父节点，从父节点的列开始
		}

		// 设置单元格的值
		cell, _ := excelize.CoordinatesToCellName(currentColumn, row)
		e.file.SetCellValue("Sheet1", cell, node.tagName)
		e.index[node.fieldPath] = currentColumn

		if len(node.subItems) == 0 { // 如果是叶子节点
			if e.offset > row { // 需要跨行合并
				cellEnd, _ := excelize.CoordinatesToCellName(currentColumn, e.offset)
				mergeRanges = append(mergeRanges, [2]string{cell, cellEnd})
			}
			currentColumn++ // 移动到下一个列
		} else { // 如果有子节点
			for _, child := range node.subItems {
				buildHeaderForRow(child, row+1, &colStart) // 递归构建子节点表头
			}
			if row == 1 { // 如果是第一层嵌套，记录合并的单元格范围
				cellEnd, _ := excelize.CoordinatesToCellName(currentColumn-1, row)
				if cell != cellEnd { // 避免单列合并
					mergeRanges = append(mergeRanges, [2]string{cell, cellEnd})
				}
			}
		}
	}

	// 从根节点开始递归
	for _, item := range e.template.items.subItems {
		buildHeaderForRow(item, 1, nil)
	}

	// 执行合并单元格操作
	for _, mergeRange := range mergeRanges {
		e.file.MergeCell("Sheet1", mergeRange[0], mergeRange[1])
	}

	return
}

func (e *exporter[T]) AddRecord(record T) error {
	v := reflect.ValueOf(record)
	startIdx := e.offset + 1

	var insert func(v reflect.Value, node *itemNode)
	insert = func(v reflect.Value, node *itemNode) {
		for _, item := range node.subItems {
			val := v.FieldByName(item.name)
			switch val.Kind() {
			case reflect.Struct:
				insert(val, item)

			case reflect.Slice:
				currentIdx := e.offset + 1
				for i := 0; i < val.Len(); i++ {
					val := val.Index(i)
					if val.Kind() == reflect.Struct {
						tmpIdx := startIdx
						startIdx = currentIdx
						insert(val, item)
						startIdx = tmpIdx
					} else {
						cell, _ := excelize.CoordinatesToCellName(e.index[item.fieldPath], currentIdx)
						e.file.SetCellValue("Sheet1", cell, val.Interface())
					}
					currentIdx++
				}
				startIdx = max(startIdx, currentIdx)

			default:
				cell, _ := excelize.CoordinatesToCellName(e.index[item.fieldPath], startIdx)
				e.file.SetCellValue("Sheet1", cell, val.Interface())
			}
		}
		return
	}

	insert(v, e.template.items)
	e.offset = startIdx - 1
	return nil
}

func (e *exporter[T]) Export() *excelize.File {
	return e.file
}
