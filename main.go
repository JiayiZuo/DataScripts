package main

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"log"
	"os"
	"path/filepath"
)

// MemberUsage 记录成员使用情况
type MemberUsage struct {
	Name       string
	Username   string
	Department string
	UsedDays   int
	LastUsed   string
	Platform   string
}

func main() {
	// 获取当前目录
	dir, err := os.Getwd()
	if err != nil {
		log.Fatal(err)
	}

	// 查找Excel文件
	excelFile := findExcelFile(dir)
	if excelFile == "" {
		log.Fatal("未找到Excel文件")
	}

	fmt.Printf("找到文件: %s\n", excelFile)

	// 打开Excel文件
	f, err := excelize.OpenFile(excelFile)
	if err != nil {
		log.Fatal(err)
	}
	defer func() {
		if err := f.Close(); err != nil {
			log.Fatal(err)
		}
	}()

	// 获取工作表名称
	sheetName := "成员使用记录"
	rows, err := f.GetRows(sheetName)
	if err != nil {
		log.Fatal(err)
	}

	// 创建映射来跟踪每个用户的使用情况
	userUsage := make(map[string]*MemberUsage)

	// 处理数据行（跳过标题行）
	for i, row := range rows {
		if i == 0 {
			continue // 跳过标题行
		}

		if len(row) < 7 {
			continue // 跳过不完整的记录
		}

		date := row[0]
		name := row[1]
		username := row[2]
		department := row[3]
		status := row[4]
		lastUsed := row[5]
		platform := row[6]

		// 如果用户尚未在映射中，则添加
		if _, exists := userUsage[username]; !exists {
			userUsage[username] = &MemberUsage{
				Name:       name,
				Username:   username,
				Department: department,
				UsedDays:   0,
				LastUsed:   "",
				Platform:   platform,
			}
		}

		// 如果状态是"使用"，增加使用天数
		if status == "使用" && lastUsed != userUsage[username].LastUsed {
			userUsage[username].UsedDays++
			// 更新最后使用时间和平台
			if date > userUsage[username].LastUsed && lastUsed != "--" {
				userUsage[username].LastUsed = date
			}
			userUsage[username].Platform = platform
		}
	}

	// 统计从未使用过的用户
	var neverUsed []*MemberUsage
	for _, user := range userUsage {
		if user.UsedDays == 0 {
			neverUsed = append(neverUsed, user)
		}
	}

	// 输出结果到控制台
	fmt.Println("\n拔萃资本成员使用情况分析")
	fmt.Println("========================")
	fmt.Printf("总成员数: %d\n", len(userUsage))
	fmt.Printf("从未使用过的成员数: %d\n", len(neverUsed))

	// 生成Excel报告
	generateExcelReport(userUsage, neverUsed, excelFile)
}

// findExcelFile 在当前目录查找Excel文件
func findExcelFile(dir string) string {
	files, err := os.ReadDir(dir)
	if err != nil {
		log.Fatal(err)
	}

	for _, file := range files {
		if !file.IsDir() {
			ext := filepath.Ext(file.Name())
			if ext == ".xlsx" || ext == ".xls" {
				return file.Name()
			}
		}
	}
	return ""
}

// generateExcelReport 生成Excel使用报告
func generateExcelReport(userUsage map[string]*MemberUsage, neverUsed []*MemberUsage, sourceFile string) {
	// 创建新的Excel文件
	newFile := excelize.NewFile()
	defer func() {
		if err := newFile.Close(); err != nil {
			log.Fatal(err)
		}
	}()

	// 设置默认工作表名称
	index := newFile.GetActiveSheetIndex()
	defaultSheet := newFile.GetSheetName(index)
	newFile.SetSheetName(defaultSheet, "汇总统计")

	// 在汇总统计工作表中添加数据
	newFile.SetCellValue("汇总统计", "A1", "拔萃资本成员使用情况报告")
	newFile.SetCellValue("汇总统计", "A2", "生成时间")
	newFile.SetCellValue("汇总统计", "B2", "基于源文件: "+sourceFile)
	newFile.SetCellValue("汇总统计", "A3", "总成员数")
	newFile.SetCellValue("汇总统计", "B3", len(userUsage))
	newFile.SetCellValue("汇总统计", "A4", "从未使用过的成员数")
	newFile.SetCellValue("汇总统计", "B4", len(neverUsed))

	// 设置标题样式
	titleStyle, _ := newFile.NewStyle(&excelize.Style{
		Font: &excelize.Font{Bold: true, Size: 16},
	})
	newFile.SetCellStyle("汇总统计", "A1", "A1", titleStyle)

	headerStyle, _ := newFile.NewStyle(&excelize.Style{
		Font: &excelize.Font{Bold: true},
		Fill: excelize.Fill{Type: "pattern", Color: []string{"#DDEBF7"}, Pattern: 1},
	})

	// 创建"从未使用过的成员"工作表
	newFile.NewSheet("从未使用过的成员")
	newFile.SetCellValue("从未使用过的成员", "A1", "姓名")
	newFile.SetCellValue("从未使用过的成员", "B1", "账号")
	newFile.SetCellValue("从未使用过的成员", "C1", "部门")
	newFile.SetCellStyle("从未使用过的成员", "A1", "C1", headerStyle)

	// 填充从未使用过的成员数据
	for i, user := range neverUsed {
		row := i + 2
		newFile.SetCellValue("从未使用过的成员", fmt.Sprintf("A%d", row), user.Name)
		newFile.SetCellValue("从未使用过的成员", fmt.Sprintf("B%d", row), user.Username)
		newFile.SetCellValue("从未使用过的成员", fmt.Sprintf("C%d", row), user.Department)
	}

	// 创建"所有成员使用情况"工作表
	newFile.NewSheet("所有成员使用情况")
	headers := []string{"姓名", "账号", "部门", "使用天数", "最后使用时间", "平台"}
	for i, header := range headers {
		cell, _ := excelize.CoordinatesToCellName(i+1, 1)
		newFile.SetCellValue("所有成员使用情况", cell, header)
	}
	newFile.SetCellStyle("所有成员使用情况", "A1", fmt.Sprintf("%c1", 'A'+len(headers)-1), headerStyle)

	// 填充所有成员使用情况数据
	row := 2
	for _, user := range userUsage {
		lastUsed := user.LastUsed
		if lastUsed == "" {
			lastUsed = "从未使用"
		}

		values := []interface{}{user.Name, user.Username, user.Department, user.UsedDays, lastUsed, user.Platform}
		for i, value := range values {
			cell, _ := excelize.CoordinatesToCellName(i+1, row)
			newFile.SetCellValue("所有成员使用情况", cell, value)
		}
		row++
	}

	// 自动调整列宽
	adjustColumnWidth(newFile, "从未使用过的成员", 3)
	adjustColumnWidth(newFile, "所有成员使用情况", 6)

	// 保存文件
	outputFile := "成员使用分析报告.xlsx"
	if err := newFile.SaveAs(outputFile); err != nil {
		log.Fatal(err)
	}

	fmt.Printf("已生成Excel报告: %s\n", outputFile)
	fmt.Printf("包含工作表:\n")
	fmt.Printf("  1. 汇总统计\n")
	fmt.Printf("  2. 从未使用过的成员 (%d人)\n", len(neverUsed))
	fmt.Printf("  3. 所有成员使用情况 (%d人)\n", len(userUsage))
}

// adjustColumnWidth 自动调整列宽
func adjustColumnWidth(f *excelize.File, sheet string, colCount int) {
	for i := 1; i <= colCount; i++ {
		colName, _ := excelize.ColumnNumberToName(i)
		maxWidth := 0

		// 获取该列所有行的数据
		rows, err := f.GetRows(sheet)
		if err != nil {
			continue
		}

		for rowIdx, row := range rows {
			if rowIdx == 0 { // 跳过标题行，因为标题行已经有样式
				continue
			}

			if len(row) >= i {
				cellValue := row[i-1]
				// 计算字符串长度（中文字符算2个宽度）
				width := calculateStringWidth(cellValue)
				if width > maxWidth {
					maxWidth = width
				}
			}
		}

		// 设置列宽，最小宽度为10，最大宽度为50
		if maxWidth < 10 {
			maxWidth = 10
		} else if maxWidth > 50 {
			maxWidth = 50
		}

		f.SetColWidth(sheet, colName, colName, float64(maxWidth))
	}
}

// calculateStringWidth 计算字符串宽度（中文字符算2个宽度）
func calculateStringWidth(s string) int {
	width := 0
	for _, r := range s {
		if r >= 0x4E00 && r <= 0x9FFF { // 中文字符范围
			width += 2
		} else {
			width += 1
		}
	}
	return width
}

// getSheetData 获取工作表数据（辅助函数）
func getSheetData(f *excelize.File, sheet string) ([][]string, error) {
	return f.GetRows(sheet)
}
