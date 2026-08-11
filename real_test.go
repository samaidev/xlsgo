package xlsgo

import (
	"strings"
	"testing"
)

// TestReadRealXlsx 回归：对真实 Excel 2007+ 文档（来自 LibreOffice 测试数据）
// 应成功列出工作表并读取单元格内容。
func TestReadRealXlsx(t *testing.T) {
	const path = "testdata/sample.xlsx"
	wb, err := Open(path)
	if err != nil {
		t.Skipf("跳过：未找到真实文件 %s（%v）", path, err)
	}
	sheets := wb.Sheets()
	if len(sheets) == 0 {
		t.Fatalf("真实 xlsx 未解析出任何工作表")
	}
	s := wb.GetSheet(sheets[0])
	if c, err := s.GetCell("A1"); err != nil || strings.TrimSpace(c) == "" {
		t.Fatalf("真实 xlsx 未读取到 A1 单元格内容 (cell=%q err=%v)", c, err)
	}
}

// TestReadRealXls 回归：对真实 BIFF(.xls) 文档应提取出表头与数据行。
func TestReadRealXls(t *testing.T) {
	const path = "testdata/sample.xls"
	wb, err := OpenLegacy(path)
	if err != nil {
		t.Skipf("跳过：未找到真实文件 %s（%v）", path, err)
	}
	txt, err := wb.Text()
	if err != nil {
		t.Fatalf("Text() 错误: %v", err)
	}
	if !strings.Contains(txt, "Col 1") {
		t.Fatalf("真实 xls 未提取出预期表头，preview=%q", txt[:min(60, len(txt))])
	}
}

func min(a, b int) int {
	if a < b {
		return a
	}
	return b
}
