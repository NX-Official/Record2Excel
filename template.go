package Record2Excel

import (
	"fmt"
	"reflect"
	"strings"
)

type template[T any] struct {
	t     reflect.Type
	items *itemNode
}

type itemNode struct {
	name      string
	tagName   string
	fieldPath string
	subItems  []*itemNode
}

func (t template[T]) depth() int {
	return t.depth()
}
func (n itemNode) depth() int {
	if len(n.subItems) == 0 {
		return 1
	}
	maxDepth := 0
	for _, child := range n.subItems {
		childDepth := child.depth()
		if childDepth > maxDepth {
			maxDepth = childDepth
		}
	}
	return maxDepth + 1
}

func newTemplate[T any](model T) (template[T], error) {
	return template[T]{
		t:     reflect.TypeOf(model),
		items: buildItemTree(reflect.TypeOf(model), ""),
	}, nil
}

func buildItemTree(t reflect.Type, parentPath string) *itemNode {
	// 处理指针类型，我们需要其指向的元素类型
	if t.Kind() == reflect.Ptr {
		t = t.Elem()
	}

	// 创建当前节点
	node := &itemNode{
		name:      t.Name(),
		fieldPath: parentPath,
		subItems:  []*itemNode{},
	}

	if t.Kind() == reflect.Slice {
		t = t.Elem()
	}

	// 只有结构体类型才有子字段需要遍历
	if t.Kind() == reflect.Struct {
		for i := 0; i < t.NumField(); i++ {
			field := t.Field(i)
			// 构建完整的字段路径
			fieldPath := field.Name
			if parentPath != "" {
				fieldPath = parentPath + "." + field.Name
			}

			// 为当前字段创建 itemNode，并递归处理嵌套结构体字段
			childNode := buildItemTree(field.Type, fieldPath)
			childNode.name = field.Name // 更新为实际的字段名
			childNode.tagName = field.Name
			if name, ok := field.Tag.Lookup("excel"); ok {
				childNode.tagName = name
			}
			childNode.fieldPath = fieldPath
			node.subItems = append(node.subItems, childNode)
		}
	}
	return node
}

func (t template[T]) GetField(path string) (string, error) {
	val, err := getValueByPath(reflect.New(t.t).Elem().Interface(), path)
	if err != nil {
		return "", err
	}
	return fmt.Sprintf("%v", val), nil
}

func getValueByPath(v any, path string) (any, error) {
	// 将路径分割成部分
	pathParts := strings.Split(path, ".")
	val := reflect.ValueOf(v)

	// 遍历路径的每一部分，逐步深入
	for _, part := range pathParts {
		// 确保当前值可以被遍历
		if val.Kind() == reflect.Ptr {
			val = val.Elem()
		}

		// 确保我们处理的是结构体
		if val.Kind() != reflect.Struct {
			return nil, fmt.Errorf("not a struct or has no field '%s'", part)
		}

		// 获取指定的字段
		val = val.FieldByName(part)
		if !val.IsValid() {
			return nil, fmt.Errorf("field not found: %s", part)
		}
	}

	// 返回找到的值
	return val.Interface(), nil
}
