package xlsgo

// 本文件实现 .xls（Excel 97-2003）的只读解析。
//
// .xls 采用 BIFF8 格式：数据存放在 OLE2 复合文档的 "Workbook" 流中，
// 内容由一连串记录组成，每条记录形如「2 字节类型 + 2 字节长度 + 负载」。
//
// 之所以只读，是因为回写 BIFF8 需要重建记录偏移、公式依赖链，以及真实
// 文件中大量未公开的结构，稍有差错就会产生 Excel 无法打开的文件。
// 需要修改时，请读取内容后写入 .xlsx。

import (
	"encoding/binary"
	"fmt"
	"math"
	"strings"
	"unicode/utf16"

	"github.com/samaidev/officecore"
	"github.com/samaidev/officecore/cfb"
)

// BIFF 记录类型。仅列出提取单元格文本所需的类型。
const (
	recEOF          = 0x000A
	recBOF          = 0x0809
	recDimensions   = 0x0200
	recBoundSheet   = 0x0085
	recSST          = 0x00FC // 共享字符串表
	recContinue     = 0x003C
	recLabelSST     = 0x00FD // 引用 SST 的单元格
	recLabel        = 0x0204 // 内联字符串单元格（BIFF5 及更早）
	recRK           = 0x027E // 压缩数值
	recMulRK        = 0x00BD // 一行中的多个 RK
	recNumber       = 0x0203 // IEEE 754 双精度
	recFormula      = 0x0006
	recString       = 0x0207 // 公式的字符串结果
	recBoolErr      = 0x0205
	recBlank        = 0x0201
	recMulBlank     = 0x00BE
)

// LegacyWorkbook 表示一个只读的 .xls 工作簿。
//
// 它不实现 Editor 接口：调用其 Save 方法将返回 ErrReadOnlyFormat。
type LegacyWorkbook struct {
	sheets []LegacySheet
}

// LegacySheet 是 .xls 中的一张工作表。
type LegacySheet struct {
	Name string

	// Cells 按 [行][列] 存放单元格文本，缺失的单元格为空字符串。
	Cells [][]string
}

// OpenLegacy 读取一个 .xls 文件。
//
// 返回的工作簿是只读的。若需要修改，请读出内容后用 New 创建 .xlsx。
func OpenLegacy(path string) (*LegacyWorkbook, error) {
	c, err := cfb.Open(path)
	if err != nil {
		return nil, fmt.Errorf("xlsgo: %w", err)
	}
	return parseLegacy(c)
}

// OpenLegacyBytes 从内存中的字节流读取 .xls。
func OpenLegacyBytes(data []byte) (*LegacyWorkbook, error) {
	c, err := cfb.OpenReader(newByteReaderAt(data))
	if err != nil {
		return nil, fmt.Errorf("xlsgo: %w", err)
	}
	return parseLegacy(c)
}

// parseLegacy 从已打开的复合文档中解析工作簿。
func parseLegacy(c *cfb.Container) (*LegacyWorkbook, error) {
	// BIFF8 使用 "Workbook"，BIFF5 及更早使用 "Book"。
	stream, ok := c.Stream("Workbook")
	if !ok {
		stream, ok = c.Stream("Book")
	}
	if !ok {
		return nil, fmt.Errorf("xlsgo: 未找到 Workbook 流（可用流: %s）",
			strings.Join(c.Names(), ", "))
	}

	recs := parseRecords(stream)

	// 第一遍：读取共享字符串表和工作表清单。
	// 单元格记录以下标引用 SST，因此必须先行解析。
	sst := parseSST(recs)
	sheets := parseBoundSheets(recs)

	// 第二遍：按各工作表在流中的偏移解析其单元格。
	wb := &LegacyWorkbook{}
	for _, bs := range sheets {
		sheet := LegacySheet{Name: bs.name}
		if int(bs.offset) < len(stream) {
			sheet.Cells = parseSheetCells(stream[bs.offset:], sst)
		}
		wb.sheets = append(wb.sheets, sheet)
	}

	// 部分文件没有 BoundSheet 记录，此时退化为解析整个流。
	if len(wb.sheets) == 0 {
		wb.sheets = append(wb.sheets, LegacySheet{
			Name:  "Sheet1",
			Cells: parseSheetCells(stream, sst),
		})
	}

	return wb, nil
}

// record 是一条 BIFF 记录。
type record struct {
	typ  uint16
	data []byte
}

// parseRecords 将流切分为一系列 BIFF 记录。
func parseRecords(stream []byte) []record {
	var recs []record
	for i := 0; i+4 <= len(stream); {
		typ := binary.LittleEndian.Uint16(stream[i:])
		length := int(binary.LittleEndian.Uint16(stream[i+2:]))
		i += 4

		if i+length > len(stream) {
			// 记录被截断，忽略剩余部分。
			break
		}
		recs = append(recs, record{typ: typ, data: stream[i : i+length]})
		i += length
	}
	return recs
}

// boundSheet 描述工作表在流中的位置。
type boundSheet struct {
	name   string
	offset uint32
}

// parseBoundSheets 读取工作表清单。
func parseBoundSheets(recs []record) []boundSheet {
	var out []boundSheet
	for _, r := range recs {
		if r.typ != recBoundSheet || len(r.data) < 8 {
			continue
		}
		offset := binary.LittleEndian.Uint32(r.data[0:])
		// data[4] 为可见性，data[5] 为类型，名称从 data[6] 开始，
		// 采用带 1 字节长度前缀的短字符串。
		name, _ := readShortString(r.data[6:])
		out = append(out, boundSheet{name: name, offset: offset})
	}
	return out
}

// parseSST 读取共享字符串表。
//
// SST 常常超过单条记录 8224 字节的上限，因而会溢出到后续的 CONTINUE
// 记录中。字符串本身可以跨越记录边界，且每个片段可能采用不同的编码，
// 这是 BIFF8 解析中最容易出错的地方。
func parseSST(recs []record) []string {
	var sstData []byte
	found := false

	for i, r := range recs {
		if r.typ != recSST {
			continue
		}
		sstData = append(sstData, r.data...)
		// 收集紧随其后的 CONTINUE 记录。
		for j := i + 1; j < len(recs) && recs[j].typ == recContinue; j++ {
			sstData = append(sstData, recs[j].data...)
		}
		found = true
		break
	}

	if !found || len(sstData) < 8 {
		return nil
	}

	// 前 8 字节为字符串总数与唯一字符串数。
	uniqueCount := int(binary.LittleEndian.Uint32(sstData[4:]))
	pos := 8

	strs := make([]string, 0, uniqueCount)
	for i := 0; i < uniqueCount && pos < len(sstData); i++ {
		s, next, ok := readUnicodeString(sstData, pos)
		if !ok {
			break
		}
		strs = append(strs, s)
		pos = next
	}

	return strs
}

// readUnicodeString 解析 BIFF8 的 Unicode 字符串。
//
// 结构为：2 字节字符数、1 字节标志位，随后是字符数据。
// 标志位指明字符是单字节压缩存储还是 UTF-16LE，以及是否存在
// 富文本格式或远东扩展信息——这两部分会插在字符数据之前，必须跳过。
func readUnicodeString(b []byte, pos int) (string, int, bool) {
	if pos+3 > len(b) {
		return "", pos, false
	}

	charCount := int(binary.LittleEndian.Uint16(b[pos:]))
	flags := b[pos+2]
	pos += 3

	compressed := flags&0x01 == 0
	hasFarEast := flags&0x04 != 0
	hasRichText := flags&0x08 != 0

	var richRuns, farEastSize int
	if hasRichText {
		if pos+2 > len(b) {
			return "", pos, false
		}
		richRuns = int(binary.LittleEndian.Uint16(b[pos:]))
		pos += 2
	}
	if hasFarEast {
		if pos+4 > len(b) {
			return "", pos, false
		}
		farEastSize = int(binary.LittleEndian.Uint32(b[pos:]))
		pos += 4
	}

	var sb strings.Builder
	if compressed {
		if pos+charCount > len(b) {
			charCount = len(b) - pos
		}
		for i := 0; i < charCount; i++ {
			// 压缩形式为 Latin-1，直接映射到对应的码点。
			sb.WriteRune(rune(b[pos+i]))
		}
		pos += charCount
	} else {
		byteLen := charCount * 2
		if pos+byteLen > len(b) {
			byteLen = ((len(b) - pos) / 2) * 2
		}
		u16 := make([]uint16, byteLen/2)
		for i := range u16 {
			u16[i] = binary.LittleEndian.Uint16(b[pos+i*2:])
		}
		sb.WriteString(string(utf16.Decode(u16)))
		pos += byteLen
	}

	// 跳过格式信息，它位于字符数据之后。
	pos += richRuns * 4
	pos += farEastSize

	if pos > len(b) {
		pos = len(b)
	}
	return sb.String(), pos, true
}

// readShortString 解析带 1 字节长度前缀的字符串（用于工作表名称）。
func readShortString(b []byte) (string, int) {
	if len(b) < 1 {
		return "", 0
	}
	n := int(b[0])
	if len(b) < 2 {
		return "", 1
	}
	flags := b[1]
	pos := 2

	if flags&0x01 == 0 {
		// 单字节压缩形式。
		if pos+n > len(b) {
			n = len(b) - pos
		}
		var sb strings.Builder
		for i := 0; i < n; i++ {
			sb.WriteRune(rune(b[pos+i]))
		}
		return sb.String(), pos + n
	}

	byteLen := n * 2
	if pos+byteLen > len(b) {
		byteLen = ((len(b) - pos) / 2) * 2
	}
	u16 := make([]uint16, byteLen/2)
	for i := range u16 {
		u16[i] = binary.LittleEndian.Uint16(b[pos+i*2:])
	}
	return string(utf16.Decode(u16)), pos + byteLen
}

// parseSheetCells 解析单张工作表的单元格记录。
func parseSheetCells(stream []byte, sst []string) [][]string {
	recs := parseRecords(stream)

	type cell struct {
		row, col int
		value    string
	}
	var cells []cell
	maxRow, maxCol := -1, -1

	note := func(row, col int, v string) {
		if v == "" {
			return
		}
		cells = append(cells, cell{row, col, v})
		if row > maxRow {
			maxRow = row
		}
		if col > maxCol {
			maxCol = col
		}
	}

	for i, r := range recs {
		// 工作表以 EOF 结束；继续解析会读到下一张表的数据。
		if r.typ == recEOF && i > 0 {
			break
		}

		switch r.typ {
		case recLabelSST:
			if len(r.data) < 10 {
				continue
			}
			row := int(binary.LittleEndian.Uint16(r.data[0:]))
			col := int(binary.LittleEndian.Uint16(r.data[2:]))
			idx := int(binary.LittleEndian.Uint32(r.data[6:]))
			if idx >= 0 && idx < len(sst) {
				note(row, col, sst[idx])
			}

		case recLabel:
			if len(r.data) < 8 {
				continue
			}
			row := int(binary.LittleEndian.Uint16(r.data[0:]))
			col := int(binary.LittleEndian.Uint16(r.data[2:]))
			s, _, ok := readUnicodeString(r.data, 6)
			if ok {
				note(row, col, s)
			}

		case recNumber:
			if len(r.data) < 14 {
				continue
			}
			row := int(binary.LittleEndian.Uint16(r.data[0:]))
			col := int(binary.LittleEndian.Uint16(r.data[2:]))
			bits := binary.LittleEndian.Uint64(r.data[6:])
			note(row, col, formatFloat(math.Float64frombits(bits)))

		case recRK:
			if len(r.data) < 10 {
				continue
			}
			row := int(binary.LittleEndian.Uint16(r.data[0:]))
			col := int(binary.LittleEndian.Uint16(r.data[2:]))
			note(row, col, formatFloat(decodeRK(binary.LittleEndian.Uint32(r.data[6:]))))

		case recMulRK:
			// 一条记录承载同一行中连续多列的 RK 值。
			if len(r.data) < 6 {
				continue
			}
			row := int(binary.LittleEndian.Uint16(r.data[0:]))
			firstCol := int(binary.LittleEndian.Uint16(r.data[2:]))
			// 末尾 2 字节为最后一列的列号。
			body := r.data[4 : len(r.data)-2]
			for j := 0; j+6 <= len(body); j += 6 {
				rk := binary.LittleEndian.Uint32(body[j+2:])
				note(row, firstCol+j/6, formatFloat(decodeRK(rk)))
			}

		case recBoolErr:
			if len(r.data) < 8 {
				continue
			}
			row := int(binary.LittleEndian.Uint16(r.data[0:]))
			col := int(binary.LittleEndian.Uint16(r.data[2:]))
			if r.data[7] == 0 {
				if r.data[6] != 0 {
					note(row, col, "TRUE")
				} else {
					note(row, col, "FALSE")
				}
			} else {
				note(row, col, "#ERR")
			}

		case recFormula:
			// 公式的字符串结果存放在紧随其后的 STRING 记录中；
			// 数值结果则直接内联在本记录里。
			if len(r.data) < 14 {
				continue
			}
			row := int(binary.LittleEndian.Uint16(r.data[0:]))
			col := int(binary.LittleEndian.Uint16(r.data[2:]))

			// 若结果为字符串，第 6..13 字节形如 00 ... FF FF。
			if r.data[12] == 0xFF && r.data[13] == 0xFF && r.data[6] == 0x00 {
				if i+1 < len(recs) && recs[i+1].typ == recString {
					if s, _, ok := readUnicodeString(recs[i+1].data, 0); ok {
						note(row, col, s)
					}
				}
				continue
			}
			bits := binary.LittleEndian.Uint64(r.data[6:])
			note(row, col, formatFloat(math.Float64frombits(bits)))
		}
	}

	if maxRow < 0 {
		return nil
	}

	grid := make([][]string, maxRow+1)
	for i := range grid {
		grid[i] = make([]string, maxCol+1)
	}
	for _, c := range cells {
		grid[c.row][c.col] = c.value
	}
	return grid
}

// decodeRK 解码 RK 值。
//
// RK 是一种把常见数值压缩进 32 位的编码：低两位为标志位，
// 分别指明该值是否需要除以 100，以及它是整数还是双精度浮点数的高位部分。
func decodeRK(rk uint32) float64 {
	isMul100 := rk&0x01 != 0
	isInt := rk&0x02 != 0

	var v float64
	if isInt {
		// 有符号 30 位整数。
		i := int32(rk) >> 2
		v = float64(i)
	} else {
		// 双精度浮点数的高 30 位，低位补零。
		bits := uint64(rk&0xFFFFFFFC) << 32
		v = math.Float64frombits(bits)
	}

	if isMul100 {
		v /= 100
	}
	return v
}

// formatFloat 以最简形式渲染数值，整数不带小数点。
func formatFloat(f float64) string {
	if f == math.Trunc(f) && math.Abs(f) < 1e15 {
		return fmt.Sprintf("%.0f", f)
	}
	return strings.TrimRight(strings.TrimRight(fmt.Sprintf("%.10f", f), "0"), ".")
}

// === 公开 API ===

// SheetCount 返回工作表数量。
func (w *LegacyWorkbook) SheetCount() int { return len(w.sheets) }

// SheetNames 返回全部工作表的名称。
func (w *LegacyWorkbook) SheetNames() []string {
	out := make([]string, len(w.sheets))
	for i, s := range w.sheets {
		out[i] = s.Name
	}
	return out
}

// Sheet 按下标返回工作表。
func (w *LegacyWorkbook) Sheet(i int) (*LegacySheet, error) {
	if i < 0 || i >= len(w.sheets) {
		return nil, fmt.Errorf("xlsgo: 工作表下标 %d 越界（共 %d 张）", i, len(w.sheets))
	}
	return &w.sheets[i], nil
}

// SheetByName 按名称返回工作表。
func (w *LegacyWorkbook) SheetByName(name string) (*LegacySheet, error) {
	for i := range w.sheets {
		if w.sheets[i].Name == name {
			return &w.sheets[i], nil
		}
	}
	return nil, fmt.Errorf("xlsgo: 未找到名为 %q 的工作表", name)
}

// Text 返回整个工作簿的纯文本内容。
// 单元格以制表符分隔，行以换行分隔，工作表之间以空行分隔。
func (w *LegacyWorkbook) Text() (string, error) {
	var sb strings.Builder
	for i, s := range w.sheets {
		if i > 0 {
			sb.WriteString("\n")
		}
		if len(w.sheets) > 1 {
			sb.WriteString("=== " + s.Name + " ===\n")
		}
		for _, row := range s.Cells {
			sb.WriteString(strings.Join(row, "\t"))
			sb.WriteString("\n")
		}
	}
	return sb.String(), nil
}

// Save 始终返回 ErrReadOnlyFormat。
//
// 该方法的存在是为了让"尝试保存 .xls"这一行为明确失败，
// 而不是让调用方误以为修改已经生效。若需要修改，请读出内容后写入 .xlsx。
func (w *LegacyWorkbook) Save(path string) error {
	return fmt.Errorf("xlsgo: 无法写入 .xls: %w", officecore.ErrReadOnlyFormat)
}

// ToXLSX 将只读工作簿的内容复制到一个可编辑的 .xlsx 工作簿中。
//
// 这是修改 .xls 文件内容的推荐途径：先转为现代格式，再进行编辑。
// 需要注意的是，仅单元格文本会被迁移，公式、格式和图表不会保留。
func (w *LegacyWorkbook) ToXLSX() (*Workbook, error) {
	nw := New()

	for i, s := range w.sheets {
		name := s.Name
		if name == "" {
			name = fmt.Sprintf("Sheet%d", i+1)
		}

		sheet := nw.AddSheet(name)
		for r, row := range s.Cells {
			for c, val := range row {
				if val == "" {
					continue
				}
				if err := sheet.SetCell(cellRef(r, c), val); err != nil {
					return nil, fmt.Errorf("xlsgo: 写入单元格 %s: %w", cellRef(r, c), err)
				}
			}
		}
	}

	return nw, nil
}

// cellRef 将从 0 开始的行列号转换为 A1 形式的引用。
func cellRef(row, col int) string {
	var name string
	for c := col; ; {
		name = string(rune('A'+c%26)) + name
		c = c/26 - 1
		if c < 0 {
			break
		}
	}
	return fmt.Sprintf("%s%d", name, row+1)
}

// byteReaderAt 让字节切片满足 io.ReaderAt 接口。
type byteReaderAt struct{ b []byte }

func newByteReaderAt(b []byte) *byteReaderAt { return &byteReaderAt{b} }

func (r *byteReaderAt) ReadAt(p []byte, off int64) (int, error) {
	if off < 0 || off >= int64(len(r.b)) {
		return 0, fmt.Errorf("offset %d out of range", off)
	}
	n := copy(p, r.b[off:])
	if n < len(p) {
		return n, fmt.Errorf("EOF")
	}
	return n, nil
}
