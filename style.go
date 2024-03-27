package Record2Excel

import "github.com/xuri/excelize/v2"

var (
	StyleHorizontalCenter = &excelize.Style{Alignment: &excelize.Alignment{Horizontal: "center"}}
	StyleVerticalCenter   = &excelize.Style{Alignment: &excelize.Alignment{Vertical: "center"}}
	StyleCenter           = &excelize.Style{Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"}}
)
