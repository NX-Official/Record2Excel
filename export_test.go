package Record2Excel

import (
	"testing"
)

func Test_exporter_buildHeader(t *testing.T) {
	type Achievements struct {
		Name         string  `excel:"名称"`
		Score        float64 `excel:"分数"`
		ArchiveScore float64 `excel:"归档分数"`
		Achieved     bool    `excel:"是否达成"`
	}

	type Pins struct {
		Name  string  `excel:"名称"`
		Score float64 `excel:"分数"`
	}

	type Data struct {
		ID           string         `excel:"编号"`
		ProjectID    string         `excel:"项目编号"`
		StaffId      string         `excel:"员工编号"`
		StaffName    string         `excel:"员工姓名"`
		ClassNo      string         `excel:"班级编号"`
		Score        float64        `excel:"总分"`
		Tags         []string       `excel:"标签"`
		Achievements []Achievements `excel:"成就"`
		Pins         []Pins         `excel:"奖项"`
	}
	s := Data{
		ID:        "123",
		ProjectID: "123",
		StaffId:   "21",
		StaffName: "231",
		ClassNo:   "21",
		Score:     23,
		Tags:      []string{"@31", "321"},
		Achievements: []Achievements{
			{
				Name:         "123",
				Score:        23,
				ArchiveScore: 23,
				Achieved:     true,
			},
			{
				Name:         "123456",
				Score:        23,
				ArchiveScore: 23,
				Achieved:     true,
			},
			{
				ArchiveScore: 23,
				Achieved:     true,
			},
		},
		Pins: []Pins{
			{
				Name:  "123",
				Score: 23,
			},
		},
	}

	e, _ := NewExporter(s)
	e.AddRecord(s)
	e.AddRecord(s)
	e.AddRecord(s)
	file := e.Export()
	file.SaveAs("test.xlsx")
}
