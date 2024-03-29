package Record2Excel

import (
	"testing"
)

var overview = Overview{
	StaffId:     "21063217",
	StaffName:   "童衍鑫",
	ClassNo:     "21063112",
	Score:       99.28,
	Tag:         nil,
	Awards:      nil,
	Achievement: nil,
	Pin: []OverviewPin{
		{Name: "学生干部加分", Score: 6},
		{Name: "文体类竞赛", Score: 2},
		{Name: "德育分", Score: 99.91},
		{Name: "加权平均分", Score: 94.01},
		{Name: "竞赛绩点加成", Score: 5},
		{Name: "智育分", Score: 99.01},
		{Name: "绩点", Score: 4.861},
	},
}

var achievement = Achievement{
	StaffId:     "21063217",
	StaffName:   "童衍鑫",
	ClassNo:     "21063112",
	Achievement: nil,
}

var pin = Pin{
	StaffId:   "21063217",
	StaffName: "童衍鑫",
	ClassNo:   "21063112",
	Pin: map[string]float64{
		"学生干部加分": 6,
		"文体类竞赛":  2,
		"德育分":    99.91,
		"加权平均分":  94.01,
		"竞赛绩点加成": 5,
		"智育分":    99.01,
		"绩点":     4.861,
	},
}

func Test_exporter_buildHeader(t *testing.T) {
	wb := NewWorkBook()
	o, _ := wb.AddSheet("overview", overview)
	o.AddRecord(overview)
	a, _ := wb.AddSheet("achievement", achievement)
	a.AddRecords([]Achievement{achievement, achievement, achievement})
	p, _ := wb.AddSheet("pin", pin)
	p.AddRecord(pin)
	p.AddRecord(pin)
	p.AddRecord(pin)

	file, _ := wb.Export()
	if err := file.SaveAs("test.xlsx"); err != nil {
		t.Fatal(err)
	}
}
