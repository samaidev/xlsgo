package xlsgo

import (
	"encoding/binary"
	"errors"
	"math"
	"strings"
	"testing"

	"github.com/samaidev/officecore"
)

// biffBuilder 用于拼装 BIFF 记录流。
type biffBuilder struct{ buf []byte }

// rec 追加一条记录：2 字节类型 + 2 字节长度 + 负载。
func (b *biffBuilder) rec(typ uint16, payload []byte) {
	var hdr [4]byte
	binary.LittleEndian.PutUint16(hdr[0:], typ)
	binary.LittleEndian.PutUint16(hdr[2:], uint16(len(payload)))
	b.buf = append(b.buf, hdr[:]...)
	b.buf = append(b.buf, payload...)
}

func u16(v int) []byte {
	b := make([]byte, 2)
	binary.LittleEndian.PutUint16(b, uint16(v))
	return b
}

func u32(v uint32) []byte {
	b := make([]byte, 4)
	binary.LittleEndian.PutUint32(b, v)
	return b
}

// sstPayload 构造共享字符串表的负载，字符串采用压缩（单字节）形式。
func sstPayload(strs []string) []byte {
	out := append(u32(uint32(len(strs))), u32(uint32(len(strs)))...)
	for _, s := range strs {
		out = append(out, u16(len([]rune(s)))...)
		out = append(out, 0x00) // 标志位：压缩存储，无富文本
		for _, r := range s {
			out = append(out, byte(r))
		}
	}
	return out
}

// sstPayloadUTF16 构造采用 UTF-16 存储的共享字符串表，用于验证中文。
func sstPayloadUTF16(strs []string) []byte {
	out := append(u32(uint32(len(strs))), u32(uint32(len(strs)))...)
	for _, s := range strs {
		runes := []rune(s)
		out = append(out, u16(len(runes))...)
		out = append(out, 0x01) // 标志位：UTF-16LE
		for _, r := range runes {
			out = append(out, byte(r), byte(r>>8))
		}
	}
	return out
}

// labelSST 构造一个引用共享字符串表的单元格记录。
func labelSST(row, col, sstIdx int) []byte {
	p := append(u16(row), u16(col)...)
	p = append(p, u16(0)...) // 格式索引
	p = append(p, u32(uint32(sstIdx))...)
	return p
}

// numberRec 构造一个双精度数值单元格记录。
func numberRec(row, col int, v float64) []byte {
	p := append(u16(row), u16(col)...)
	p = append(p, u16(0)...)
	bits := make([]byte, 8)
	binary.LittleEndian.PutUint64(bits, math.Float64bits(v))
	return append(p, bits...)
}

// rkRec 构造一个 RK 压缩数值单元格记录。
func rkRec(row, col int, rk uint32) []byte {
	p := append(u16(row), u16(col)...)
	p = append(p, u16(0)...)
	return append(p, u32(rk)...)
}

// TestParseRecordsSplitsStream 验证记录切分逻辑。
func TestParseRecordsSplitsStream(t *testing.T) {
	var b biffBuilder
	b.rec(recBOF, []byte{1, 2, 3, 4})
	b.rec(recNumber, numberRec(0, 0, 42))
	b.rec(recEOF, nil)

	recs := parseRecords(b.buf)
	if len(recs) != 3 {
		t.Fatalf("期望 3 条记录，实际 %d 条", len(recs))
	}
	if recs[0].typ != recBOF || recs[2].typ != recEOF {
		t.Errorf("记录类型解析错误: %v", []uint16{recs[0].typ, recs[1].typ, recs[2].typ})
	}
}

// TestParseRecordsHandlesTruncation 确保截断的流不会导致越界。
func TestParseRecordsHandlesTruncation(t *testing.T) {
	var b biffBuilder
	b.rec(recNumber, numberRec(0, 0, 1))
	// 追加一个声称长度很大、实际数据不足的记录头。
	b.buf = append(b.buf, 0x03, 0x02, 0xFF, 0x00)

	recs := parseRecords(b.buf) // 不应 panic
	if len(recs) != 1 {
		t.Errorf("截断记录应被忽略，实际解析出 %d 条", len(recs))
	}
}

// TestSSTCompressedStrings 验证单字节压缩形式的字符串表。
func TestSSTCompressedStrings(t *testing.T) {
	var b biffBuilder
	b.rec(recSST, sstPayload([]string{"Hello", "World"}))

	got := parseSST(parseRecords(b.buf))
	want := []string{"Hello", "World"}

	if len(got) != len(want) {
		t.Fatalf("期望 %d 个字符串，实际 %d 个: %q", len(want), len(got), got)
	}
	for i := range want {
		if got[i] != want[i] {
			t.Errorf("字符串 %d: 期望 %q，实际 %q", i, want[i], got[i])
		}
	}
}

// TestSSTUnicodeStrings 验证 UTF-16 形式，中文必须正确解码。
func TestSSTUnicodeStrings(t *testing.T) {
	var b biffBuilder
	b.rec(recSST, sstPayloadUTF16([]string{"销售额", "利润"}))

	got := parseSST(parseRecords(b.buf))
	want := []string{"销售额", "利润"}

	if len(got) != len(want) {
		t.Fatalf("期望 %d 个字符串，实际 %d 个: %q", len(want), len(got), got)
	}
	for i := range want {
		if got[i] != want[i] {
			t.Errorf("字符串 %d: 期望 %q，实际 %q", i, want[i], got[i])
		}
	}
}

// TestSSTSpanningContinueRecords 覆盖最易出错的场景：
// 字符串表超过单条记录上限，溢出到 CONTINUE 记录中。
func TestSSTSpanningContinueRecords(t *testing.T) {
	payload := sstPayload([]string{"AAA", "BBB", "CCC"})

	// 在字符串边界处切开，模拟溢出。
	split := 12
	var b biffBuilder
	b.rec(recSST, payload[:split])
	b.rec(recContinue, payload[split:])

	got := parseSST(parseRecords(b.buf))
	if len(got) != 3 {
		t.Fatalf("跨 CONTINUE 的字符串表应解析出 3 个字符串，实际 %d 个: %q", len(got), got)
	}
	if got[2] != "CCC" {
		t.Errorf("最后一个字符串应为 CCC，实际 %q", got[2])
	}
}

// encodeRKInt 将整数编码为 RK 值，供测试构造输入使用。
// 使用变量而非常量表达式，以避免负数在编译期触发溢出检查。
func encodeRKInt(v int32, div100 bool) uint32 {
	rk := uint32(v<<2) | 0x02
	if div100 {
		rk |= 0x01
	}
	return rk
}

func TestDecodeRK(t *testing.T) {
	tests := []struct {
		name string
		rk   uint32
		want float64
	}{
		{"整数 100", encodeRKInt(100, false), 100},
		{"整数除以 100", encodeRKInt(12345, true), 123.45},
		{"负整数", encodeRKInt(-50, false), -50},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := decodeRK(tt.rk)
			if math.Abs(got-tt.want) > 1e-9 {
				t.Errorf("decodeRK(%#x) = %v，期望 %v", tt.rk, got, tt.want)
			}
		})
	}
}

// TestDecodeRKFloat 验证 RK 中的浮点形式。
func TestDecodeRKFloat(t *testing.T) {
	// 取 1.5 的双精度表示的高 30 位。
	bits := math.Float64bits(1.5)
	rk := uint32(bits >> 32 & 0xFFFFFFFC)

	if got := decodeRK(rk); math.Abs(got-1.5) > 1e-9 {
		t.Errorf("浮点形式的 RK 解码为 %v，期望 1.5", got)
	}
}

// TestParseSheetCellsBuildsGrid 验证单元格记录被正确填入网格。
func TestParseSheetCellsBuildsGrid(t *testing.T) {
	sst := []string{"名称", "数量"}

	var b biffBuilder
	b.rec(recBOF, []byte{0, 0, 0, 0})
	b.rec(recLabelSST, labelSST(0, 0, 0))
	b.rec(recLabelSST, labelSST(0, 1, 1))
	b.rec(recNumber, numberRec(1, 1, 42))
	b.rec(recRK, rkRec(2, 1, uint32(7)<<2|0x02))
	b.rec(recEOF, nil)

	grid := parseSheetCells(b.buf, sst)

	if len(grid) != 3 {
		t.Fatalf("期望 3 行，实际 %d 行: %v", len(grid), grid)
	}
	if grid[0][0] != "名称" || grid[0][1] != "数量" {
		t.Errorf("表头解析错误: %v", grid[0])
	}
	if grid[1][1] != "42" {
		t.Errorf("数值单元格应为 42，实际 %q", grid[1][1])
	}
	if grid[2][1] != "7" {
		t.Errorf("RK 单元格应为 7，实际 %q", grid[2][1])
	}
}

// TestParseSheetCellsStopsAtEOF 确保不会越界读到下一张工作表的数据。
func TestParseSheetCellsStopsAtEOF(t *testing.T) {
	sst := []string{"本表", "下一张表"}

	var b biffBuilder
	b.rec(recBOF, []byte{0, 0, 0, 0})
	b.rec(recLabelSST, labelSST(0, 0, 0))
	b.rec(recEOF, nil)
	// 下一张工作表的数据不应被计入。
	b.rec(recLabelSST, labelSST(5, 5, 1))

	grid := parseSheetCells(b.buf, sst)

	if len(grid) != 1 {
		t.Errorf("解析应止于 EOF，却得到 %d 行: %v", len(grid), grid)
	}
}

func TestBoolErrCells(t *testing.T) {
	var b biffBuilder
	b.rec(recBOF, []byte{0, 0, 0, 0})
	// 行 0 列 0：布尔真
	b.rec(recBoolErr, append(append(u16(0), u16(0)...), append(u16(0), 0x01, 0x00)...))
	// 行 1 列 0：布尔假
	b.rec(recBoolErr, append(append(u16(1), u16(0)...), append(u16(0), 0x00, 0x00)...))
	b.rec(recEOF, nil)

	grid := parseSheetCells(b.buf, nil)

	if len(grid) != 2 {
		t.Fatalf("期望 2 行，实际 %d 行: %v", len(grid), grid)
	}
	if grid[0][0] != "TRUE" {
		t.Errorf("期望 TRUE，实际 %q", grid[0][0])
	}
	if grid[1][0] != "FALSE" {
		t.Errorf("期望 FALSE，实际 %q", grid[1][0])
	}
}

func TestFormatFloat(t *testing.T) {
	cases := map[float64]string{
		42:      "42",
		-7:      "-7",
		0:       "0",
		1.5:     "1.5",
		123.45:  "123.45",
	}
	for in, want := range cases {
		if got := formatFloat(in); got != want {
			t.Errorf("formatFloat(%v) = %q，期望 %q", in, got, want)
		}
	}
}

func TestCellRef(t *testing.T) {
	cases := []struct {
		row, col int
		want     string
	}{
		{0, 0, "A1"},
		{0, 25, "Z1"},
		{0, 26, "AA1"},
		{9, 27, "AB10"},
	}
	for _, c := range cases {
		if got := cellRef(c.row, c.col); got != c.want {
			t.Errorf("cellRef(%d,%d) = %q，期望 %q", c.row, c.col, got, c.want)
		}
	}
}

// TestSaveReturnsReadOnlyError 验证尝试写入 .xls 会明确失败，
// 而不是让调用方误以为修改已保存。
func TestSaveReturnsReadOnlyError(t *testing.T) {
	w := &LegacyWorkbook{}

	err := w.Save("out.xls")
	if err == nil {
		t.Fatal("写入 .xls 应当返回错误")
	}
	if !errors.Is(err, officecore.ErrReadOnlyFormat) {
		t.Errorf("错误应可通过 errors.Is 匹配 ErrReadOnlyFormat，实际为: %v", err)
	}
}

func TestTextRendersSheets(t *testing.T) {
	w := &LegacyWorkbook{
		sheets: []LegacySheet{
			{Name: "第一张", Cells: [][]string{{"a", "b"}, {"c", "d"}}},
			{Name: "第二张", Cells: [][]string{{"x"}}},
		},
	}

	text, err := w.Text()
	if err != nil {
		t.Fatalf("Text: %v", err)
	}

	for _, want := range []string{"第一张", "第二张", "a\tb", "c\td", "x"} {
		if !strings.Contains(text, want) {
			t.Errorf("输出中缺少 %q:\n%s", want, text)
		}
	}
}

func TestSheetAccessors(t *testing.T) {
	w := &LegacyWorkbook{
		sheets: []LegacySheet{
			{Name: "Alpha"},
			{Name: "Beta"},
		},
	}

	if w.SheetCount() != 2 {
		t.Errorf("SheetCount = %d，期望 2", w.SheetCount())
	}
	if names := w.SheetNames(); names[0] != "Alpha" || names[1] != "Beta" {
		t.Errorf("SheetNames = %v", names)
	}
	if _, err := w.Sheet(5); err == nil {
		t.Error("越界下标应返回错误")
	}
	if _, err := w.SheetByName("不存在"); err == nil {
		t.Error("不存在的名称应返回错误")
	}
	if s, err := w.SheetByName("Beta"); err != nil || s.Name != "Beta" {
		t.Errorf("按名称查找失败: %v", err)
	}
}

// TestToXLSXConvertsContent 验证转换为可编辑格式的路径可用，
// 这是修改 .xls 内容的推荐做法。
func TestToXLSXConvertsContent(t *testing.T) {
	legacy := &LegacyWorkbook{
		sheets: []LegacySheet{
			{Name: "数据", Cells: [][]string{{"标题", "数值"}, {"甲", "100"}}},
		},
	}

	wb, err := legacy.ToXLSX()
	if err != nil {
		t.Fatalf("ToXLSX: %v", err)
	}

	data, err := wb.Bytes()
	if err != nil {
		t.Fatalf("生成 xlsx: %v", err)
	}
	if len(data) == 0 {
		t.Error("生成的 xlsx 为空")
	}
}

func TestOpenLegacyRejectsNonCFB(t *testing.T) {
	if _, err := OpenLegacyBytes([]byte("PK\x03\x04 这是一个 zip")); err == nil {
		t.Error("将 xlsx 作为 xls 打开应返回错误")
	}
}
