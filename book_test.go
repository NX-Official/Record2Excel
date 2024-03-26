package Record2Excel

import (
	"testing"
)

type (
	Achievements struct {
		Name         string  `excel:"名称"`
		Score        float64 `excel:"分数"`
		ArchiveScore float64 `excel:"达标分数"`
		Achieved     bool    `excel:"已达标"`
	}
	Pins struct {
		Name  string  `excel:"名称"`
		Score float64 `excel:"分数"`
	}
	Overview struct {
		StaffId      string         `excel:"学号"`
		StaffName    string         `excel:"姓名"`
		ClassNo      string         `excel:"班级号"`
		Score        float64        `excel:"总分"`
		Tags         []string       `excel:"获奖情况"`
		Achievements []Achievements `excel:"达标情况"`
		Pins         []Pins         `excel:"重要成绩项目"`
	}
)

type (
	AchievementOverview struct {
		StaffId      string          `excel:"学号"`
		StaffName    string          `excel:"姓名"`
		ClassNo      string          `excel:"班级号"`
		Achievements map[string]bool `excel:"达标情况"`
	}
)

func Test_exporter_buildHeader(t *testing.T) {
	overview := Overview{
		StaffId:   "22050626",
		StaffName: "王文杰",
		ClassNo:   "22184111",
		Score:     100,
		Tags:      []string{"Tag1", "Tag2", "Tag3"},
		Achievements: []Achievements{
			{
				Name:         "第一届钱潮杯普法短视频创意大赛",
				Score:        0,
				ArchiveScore: 10,
				Achieved:     false,
			},
			{
				Name:         "卡尔·马克思杯2",
				Score:        100,
				ArchiveScore: 10,
				Achieved:     true,
			},
		},
		Pins: []Pins{
			{
				Name:  "学生工作奖",
				Score: 10,
			}, {
				Name:  "学习进步奖",
				Score: 10,
			},
			{
				Name:  "优秀学生干部",
				Score: 10,
			},
		},
	}

	wb := NewWorkBook()
	sheet, err := wb.AddSheet("总览", Overview{})
	if err != nil {
		t.Fatal(err)
	}
	err = sheet.AddRecord(overview)
	if err != nil {
		t.Fatal(err)
	}
	file := wb.Export()
	if err := file.SaveAs("test.xlsx"); err != nil {
		t.Fatal(err)
	}
}
