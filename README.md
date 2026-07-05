# xlsgo

简洁的 Go 库，用于读写 .xlsx Excel 表格。设计参考 openpyxl，基于 excelize 封装。

## 安装

```bash
go get github.com/samaidev/xlsgo
```

## 快速开始

```go
package main

import "github.com/samaidev/xlsgo"

func main() {
    wb := xlsgo.New()
    ws := wb.ActiveSheet()
    ws.SetCell("A1", "姓名")
    ws.SetCell("B1", "分数")
    ws.SetCellStyleHeader("A1")
    ws.SetCell("A2", "张三")
    ws.SetCellInt("B2", 95)
    ws.SetCellFormula("C2", "=B2*0.9")
    wb.Save("scores.xlsx")
}
```

## 高级功能

- 图表：`ws.AddChart("bar", "A1:B5", "D1")`
- 条件格式：`ws.SetConditionalFormat("B2:B10", "greaterThan:90")`
- 冻结窗格：`ws.FreezePanes("A2")`
- 自动筛选：`ws.AutoFilter("A1:D10")`
- 数据验证：`ws.AddDataValidation("A2:A10", "Yes,No,Maybe")`

## 许可证

MIT
