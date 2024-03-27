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

var overview = Overview{
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
var achievementOverview = AchievementOverview{
	StaffId:   "22050626",
	StaffName: "王文杰",
	ClassNo:   "22184111",
	Achievements: map[string]bool{
		"第一届钱潮杯普法短视频创意大赛": true,
		"卡尔·马克思杯2":        false,
	},
}

func Test_exporter_buildHeader(t *testing.T) {

	wb := NewWorkBook()
	_ = overview
	sheetOverview, err := wb.AddSheet("总览", Overview{})
	if err != nil {
		t.Fatal(err)
	}
	err = sheetOverview.AddRecord(overview)
	err = sheetOverview.AddRecord(overview)
	err = sheetOverview.AddRecord(overview)
	err = sheetOverview.AddRecords([]Overview{overview, overview, overview})
	if err != nil {
		t.Fatal(err)
	}
	sheetAchievementOverview, err := wb.AddSheet("达标情况", AchievementOverview{
		Achievements: map[string]bool{
			"第一届钱潮杯普法短视频创意大赛": false,
			"卡尔·马克思杯2":        false,
		},
	})
	if err != nil {
		t.Fatal(err)
	}
	err = sheetAchievementOverview.AddRecord(achievementOverview)
	err = sheetAchievementOverview.AddRecord(achievementOverview)
	err = sheetAchievementOverview.AddRecord(achievementOverview)
	file, _ := wb.Export()
	if err := file.SaveAs("test.xlsx"); err != nil {
		t.Fatal(err)
	}
}
