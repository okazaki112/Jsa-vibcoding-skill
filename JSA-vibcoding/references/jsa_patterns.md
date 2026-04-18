# JSA 常用模式

JSA 开发中常用的操作模式和代码片段。

## 单元格操作模式

### 读取单元格

```javascript
// 基础读取
let value = Range("A1").Value2
let value2 = Cells(1, 1).Value2
let value3 = Sheets.Item(1).Range("A1").Value2

// 安全读取（处理错误值）
function SafeRead(rng) {
    const value = rng.Value2
    if (value === undefined || value === null) {
        return ""
    }
    return value
}

// 读取公式
let formula = Range("A1").Formula
let formulaR1C1 = Range("A1").FormulaR1C1

// 读取格式化文本
let text = Range("A1").Text  // 只读
```

### 写入单元格

```javascript
// 基础写入
Range("A1").Value2 = "数据"
Cells(1, 1).Value2 = 100

// 批量写入（数组）
Range("A1:D10").Value2 = dataArray

// 填充公式
Range("E2:E100").Formula = "=B2*C2"
Range("E2:E100").FormulaR1C1 = "=RC[-3]*RC[-2]"

// 公式转值
Range("E2:E100").Value2 = Range("E2:E100").Value2
```

### 复制粘贴

```javascript
// 直接复制到目标
Range("A1:D10").Copy(Range("F1"))

// 粘贴特殊
Range("A1:D10").Copy()
Range("F1").PasteSpecial(xlPasteValues)      // 只粘贴值
Range("F1").PasteSpecial(xlPasteFormats)     // 只粘贴格式
Range("F1").PasteSpecial(xlPasteFormulas)    // 只粘贴公式
```

### 选择和定位

```javascript
// 选择单元格
Range("A1").Select()
Range("A1:D10").Select()

// 激活单元格
Range("A1").Activate()

// 偏移选择
Range("A1").Offset(2, 3).Select()   // 向下2行，向右3列

// 调整大小
Range("A1").Resize(5, 3).Select()   // 5行3列

// 整行整列
Range("A1").EntireRow.Select()
Range("A1").EntireColumn.Select()
```

## 工作表操作模式

### 选择工作表

```javascript
// 按索引选择
Sheets.Item(1).Select()

// 按名称选择
Sheets.Item("Sheet1").Select()
Sheets.Item("数据表").Select()

// 激活工作表
Sheets.Item(1).Activate()
```

### 新建工作表

```javascript
// 新建工作表
function CreateSheet(name) {
    const ws = Worksheets.Add()
    ws.Name = name
    return ws
}

// 在最后位置新建
function CreateSheetAtEnd(name) {
    const ws = Worksheets.Add(null, Sheets.Item(Sheets.Count))
    ws.Name = name
    return ws
}

// 在指定位置新建
function CreateSheetAfter(name, afterName) {
    const ws = Worksheets.Add(null, Sheets.Item(afterName))
    ws.Name = name
    return ws
}
```

### 删除工作表

```javascript
function DeleteSheet(name) {
    Application.DisplayAlerts = false
    try {
        Sheets.Item(name).Delete()
    } catch (e) {
        console.log(`删除失败: ${e.message}`)
    }
    Application.DisplayAlerts = true
}
```

### 复制和移动工作表

```javascript
// 复制到工作簿末尾
Sheets.Item("Sheet1").Copy(null, Sheets.Item(Sheets.Count))

// 移动到工作簿末尾
Sheets.Item("Sheet1").Move(null, Sheets.Item(Sheets.Count))

// 复制到新工作簿
Sheets.Item("Sheet1").Copy()

// 重命名
Sheets.Item(1).Name = "新名称"
```

## 数据处理模式

### 筛选

```javascript
// 开启自动筛选
function EnableAutoFilter(ws) {
    ws.Range("A1").AutoFilter()
}

// 单条件筛选
function FilterByValue(ws, field, value) {
    ws.Range("A1").AutoFilter(field, value)
}

// 多条件筛选
function FilterByArray(ws, field, values) {
    ws.Range("A1").AutoFilter(field, values, xlFilterValues)
}

// 数值范围筛选
function FilterByRange(ws, field, minValue) {
    ws.Range("A1").AutoFilter(field, `>${minValue}`)
}

// 清除筛选
function ClearAutoFilter(ws) {
    if (ws.AutoFilterMode) {
        ws.AutoFilterMode = false
    }
}
```

### 排序

```javascript
function SortData(ws, keyCol, order) {
    const wsData = ws || ActiveSheet
    const lastRow = GetLastRow(wsData, 1)
    const rng = wsData.Range(`A1:D${lastRow}`)
    
    wsData.Sort.SortFields.Clear()
    wsData.Sort.SortFields.Add(
        wsData.Cells(2, keyCol),
        xlSortOnValues,
        order || xlAscending
    )
    
    wsData.Sort.SetRange(rng)
    wsData.Sort.Header = xlYes
    wsData.Sort.Apply()
}
```

### 删除重复值

```javascript
function RemoveDuplicates(ws, columns) {
    const wsData = ws || ActiveSheet
    const lastRow = GetLastRow(wsData, 1)
    wsData.Range(`A1:D${lastRow}`).RemoveDuplicates(columns)
}

// 示例：按第1、2列去重
RemoveDuplicates(ActiveSheet, [1, 2])
```

### 查找

```javascript
// 查找第一个匹配
function FindFirst(ws, searchValue) {
    const rng = ws.Cells.Find(searchValue)
    return rng
}

// 查找所有匹配
function FindAllMatches(ws, searchValue) {
    const results = ws.Cells.FindAll(searchValue)
    return results || []
}

// 在指定列查找
function FindInColumn(ws, col, value) {
    const rng = ws.Columns(col).Find(value)
    return rng
}
```

### 替换

```javascript
function ReplaceInSheet(ws, oldValue, newValue) {
    ws.Cells.Replace(oldValue, newValue)
}

// 替换指定范围
function ReplaceInRange(ws, rng, oldValue, newValue) {
    rng.Replace(oldValue, newValue)
}
```

## 格式化模式

### 字体和颜色

```javascript
function FormatFont(ws, rng) {
    rng.Font.Name = "微软雅黑"
    rng.Font.Size = 11
    rng.Font.Bold = true
    rng.Font.Italic = false
    rng.Font.Color = 255  // 红色
}

function FormatBackground(ws, rng, color) {
    rng.Interior.Color = color
    rng.Interior.Pattern = xlSolid
}
```

### 对齐方式

```javascript
function SetAlignment(ws, rng) {
    rng.HorizontalAlignment = xlHAlignCenter      // 水平居中
    rng.VerticalAlignment = xlVAlignCenter        // 垂直居中
    rng.WrapText = true                           // 自动换行
    rng.Orientation = 0                           // 文字方向
}
```

### 行列格式

```javascript
// 自动调整列宽
function AutoFitColumns(ws) {
    ws.Columns.AutoFit()
}

// 固定列宽
function SetColumnWidth(ws, col, width) {
    ws.Columns(col).ColumnWidth = width
}

// 隐藏/显示列
function HideColumn(ws, col, hide) {
    ws.Columns(col).Hidden = hide
}

// 自动调整行高
function AutoFitRows(ws) {
    ws.Rows.AutoFit()
}
```

### 边框

```javascript
function SetBorders(ws, rng) {
    // 外边框
    for (const idx of [xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight]) {
        const border = rng.Borders.Item(idx)
        border.LineStyle = xlContinuous
        border.Weight = xlMedium
        border.Color = 0  // 黑色
    }
    
    // 内边框
    for (const idx of [xlInsideHorizontal, xlInsideVertical]) {
        const border = rng.Borders.Item(idx)
        border.LineStyle = xlContinuous
        border.Weight = xlThin
        border.Color = 0
    }
}
```

## 合并单元格模式

```javascript
// 合并单元格
function MergeRange(ws, range) {
    ws.Range(range).Merge()
}

// 合并并居中
function MergeAndCenter(ws, range) {
    const rng = ws.Range(range)
    rng.Merge()
    rng.HorizontalAlignment = xlHAlignCenter
    rng.VerticalAlignment = xlVAlignCenter
}

// 取消合并
function UnMergeRange(ws, range) {
    ws.Range(range).UnMerge()
}

// 判断是否合并
function IsMerged(ws, range) {
    return ws.Range(range).MergeCells
}

// 获取合并区域
function GetMergeArea(ws, range) {
    return ws.Range(range).MergeArea
}
```

## 事件处理模式

### 工作表事件

```javascript
// 单元格值改变
function Workbook_SheetChange(Sh, rg) {
    console.log(`工作表 ${Sh.Name} 的 ${rg.Address()} 发生变化`)
}

// 工作表激活
function Workbook_SheetActivate(Sh) {
    console.log(`激活了工作表: ${Sh.Name}`)
}

// 工作表选择改变
function Workbook_SheetSelectionChange(Sh, rg) {
    console.log(`选中了: ${rg.Address()}`)
}

// 双击单元格
function Workbook_SheetBeforeDoubleClick(Sh, rg, Cancel) {
    console.log(`双击了 ${rg.Address()}`)
    Cancel = true  // 取消默认行为
}

// 右键单元格
function Workbook_SheetBeforeRightClick(Sh, rg, Cancel) {
    console.log(`右键了 ${rg.Address()}`)
}
```

### 工作簿事件

```javascript
// 打开工作簿
function Workbook_Open() {
    console.log("工作簿已打开")
}

// 保存前
function Workbook_BeforeSave(SaveAsUI, Cancel) {
    console.log("即将保存")
}

// 关闭前
function Workbook_BeforeClose(Cancel) {
    const result = MsgBox("确定要关闭吗?", jsYesNo, "确认")
    if (result != 6) {
        Cancel = true  // 取消关闭
    }
}

// 新建工作表
function Workbook_NewSheet(Sh) {
    console.log(`新建了工作表: ${Sh.Name}`)
}
```

## 用户交互模式

### 消息框

```javascript
// 简单消息
MsgBox("操作完成!")

// 带标题
MsgBox("操作完成!", "提示")

// 是/否选择
const result = MsgBox("确定删除吗?", jsYesNo, "确认")
if (result == 6) {
    // 点击了"是"
} else {
    // 点击了"否"
}

// 警告消息
MsgBox("数据有误!", "警告")
```

### 输入框

```javascript
// 简单输入
const name = InputBox("请输入姓名")
if (name) {
    console.log(`你输入了: ${name}`)
}

// 完整参数
const value = InputBox("请输入数值", "输入", "100", 100, 100)

// 选择单元格区域
const rng = Application.InputBox(
    "请选择区域",
    "选择",
    undefined,
    undefined,
    undefined,
    undefined,
    undefined,
    8  // 返回Range对象
)
if (rng) {
    console.log(`你选择了: ${rng.Address()}`)
}
```

## 调试模式

### 控制台输出

```javascript
// 基础输出
console.log("调试信息")

// 格式化输出
const lastRow = 100
console.log(`最后一行: ${lastRow}`)

// 输出对象
const data = { name: "张三", age: 25 }
console.log(JSON.stringify(data))

// 表格输出
console.log("序号\t值\t结果")
console.log("1\t100\t成功")
```

### 调试标记

```javascript
function DebugFunction(ws) {
    const startTime = Date.now()
    
    console.log("=".repeat(50))
    console.log(`开始处理: ${new Date().toLocaleString()}`)
    console.log("=".repeat(50))
    
    // 主逻辑
    console.log("步骤1: 读取数据")
    const data = readData(ws)
    console.log(`读取到 ${data.length} 行数据`)
    
    console.log("步骤2: 处理数据")
    process(data)
    
    console.log("步骤3: 写入结果")
    writeData(ws, data)
    
    const elapsed = ((Date.now() - startTime) / 1000).toFixed(2)
    console.log(`处理完成，用时: ${elapsed}秒`)
}
```

### 清除立即窗口

```javascript
Console.clear()
// 或
Debug.Print("\n".repeat(50))
```
