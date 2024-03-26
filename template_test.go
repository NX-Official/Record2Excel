package Record2Excel

import (
	"fmt"
	"strings"
	"testing"
)

func TestNewTemplate(t *testing.T) {
	type MyStruct struct {
		Field1 string
		Nested struct {
			SubField1 int
		}
	}

	tml, _ := newTemplate(MyStruct{})
	printItemNodeTree(tml.items, 0)
}

func printItemNodeTree(node *itemNode, indentLevel int) {
	indent := strings.Repeat("  ", indentLevel) // 根据层级重复空格，创建缩进
	fmt.Printf("%sName: %s, Path: %s\n", indent, node.name, node.fieldPath)

	// 遍历子项并递归打印
	for _, child := range node.subItems {
		printItemNodeTree(child, indentLevel+1) // 增加缩进级别
	}
}
