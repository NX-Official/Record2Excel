package Record2Excel

import (
	"testing"
)

type (
	// Overview 总览
	Overview struct {
		StaffId     string                `excel:"学号"`
		StaffName   string                `excel:"姓名"`
		ClassNo     string                `excel:"班级"`
		Score       float64               `excel:"总分"`
		Tag         []string              `excel:"标签"`
		Awards      []string              `excel:"奖项"`
		Achievement []OverviewAchievement `excel:"达标情况"`
		Pin         []OverviewPin         `excel:"重要成绩项"`
	}
	OverviewAchievement struct {
		Name         string  `excel:"奖项"`
		Score        float64 `excel:"分数"`
		AchieveScore float64 `excel:"达标分数"`
		Achieved     bool    `excel:"已达标"`
	}
	OverviewPin struct {
		Name  string  `excel:"名称"`
		Score float64 `excel:"分数"`
	}

	// Achievement 达标情况
	Achievement struct {
		StaffId     string          `excel:"学号"`
		StaffName   string          `excel:"姓名"`
		ClassNo     string          `excel:"班级"`
		Achievement map[string]bool `excel:"达标情况"`
	}

	// Pin 重要成绩项
	Pin struct {
		StaffId   string             `excel:"学号"`
		StaffName string             `excel:"姓名"`
		ClassNo   string             `excel:"班级"`
		Pin       map[string]float64 `excel:"重要成绩项"`
	}
)

var overview = Overview{
	StaffId:   "22050626",
	StaffName: "王文杰",
	ClassNo:   "22184111",
	Score:     100,
	Tag:       []string{"Tag1", "Tag2", "Tag3"},
	Awards:    []string{"奖项1", "奖项2"},
	Achievement: []OverviewAchievement{
		{
			Name:         "第一届钱潮杯普法短视频创意大赛",
			Score:        0,
			AchieveScore: 10,
			Achieved:     false,
		},
		{
			Name:         "卡尔·马克思杯2",
			Score:        100,
			AchieveScore: 10,
			Achieved:     true,
		},
	},
	Pin: []OverviewPin{
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

var achievement = Achievement{
	StaffId:   "22050626",
	StaffName: "王文杰",
	ClassNo:   "22184111",
	Achievement: map[string]bool{
		"第一届钱潮杯普法短视频创意大赛": false,
		"卡尔·马克思杯2":        true,
	},
}

var pin = Pin{
	StaffId:   "22050626",
	StaffName: "王文杰",
	ClassNo:   "22184111",
	Pin: map[string]float64{
		"学生工作奖": 10,
		"学习进步奖": 10,
	},
}

func Test_exporter_buildHeader(t *testing.T) {
	wb := NewWorkBook()
	o, _ := wb.AddSheet("overview", overview)
	o.AddRecord(overview)
	a, _ := wb.AddSheet("achievement", achievement)
	a.AddRecord(achievement)
	p, _ := wb.AddSheet("pin", pin)
	p.AddRecord(pin)

	file, _ := wb.Export()
	if err := file.SaveAs("test.xlsx"); err != nil {
		t.Fatal(err)
	}
}
