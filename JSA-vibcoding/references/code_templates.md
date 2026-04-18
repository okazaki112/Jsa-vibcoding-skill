# JSA 代码模板库

核心模板，按需复制使用。

## 标准函数结构

```javascript
/**
 * 函数功能描述
 * @param {类型} 参数名 - 参数说明
 * @returns {类型} 返回值说明
 */
function 标准函数(ws) {
    // 使用传入的工作表或当前活动表
    const wsData = ws || ActiveSheet
    const startTime = Date.now()
    
    try {
        console.log(`开始处理: ${wsData.Name}`)
        
        // ===== 主代码区域 =====
        
        // 你的处理逻辑
        
        // =====================
        
        const elapsed = ((Date.now() - startTime) / 1000).toFixed(2)
        console.log(`处理完成，用时: ${elapsed}秒`)
        
    } catch (e) {
        console.log(`错误: ${e.message}`)
        alert(`处理失败: ${e.message}`)
    }
}
```

## 数据边界检测函数

```javascript
/**
 * 获取指定列的最后一行
 * @param {Worksheet} ws - 工作表
 * @param {number} col - 列号
 * @returns {number} 最后一行号
 */
function GetLastRow(ws, col) {
    const lastCell = ws.Columns(col).Find("*", undefined, undefined, undefined, xlByRows, xlPrevious)
    return lastCell ? lastCell.Row : 0
}

/**
 * 获取指定行的最后一列
 * @param {Worksheet} ws - 工作表
 * @param {number} row - 行号
 * @returns {number} 最后一列号
 */
function GetLastCol(ws, row) {
    const lastCell = ws.Rows(row).Find("*", undefined, undefined, undefined, xlByColumns, xlPrevious)
    return lastCell ? lastCell.Column : 0
}

/**
 * 获取已使用区域
 * @param {Worksheet} ws - 工作表
 * @returns {Range} 已使用区域
 */
function GetUsedRange(ws) {
    let lastRow = 0, lastCol = 0
    const lastRowCell = ws.Cells.Find("*", undefined, undefined, undefined, xlByRows, xlPrevious)
    const lastColCell = ws.Cells.Find("*", undefined, undefined, undefined, xlByColumns, xlPrevious)
    
    if (lastRowCell) lastRow = lastRowCell.Row
    if (lastColCell) lastCol = lastColCell.Column
    
    if (lastRow > 0 && lastCol > 0) {
        return ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    }
    return null
}

/**
 * 获取列字母
 * @param {number} columnIndex - 列序数
 * @returns {string} 列字母
 */
function GetColumnLetter(columnIndex) {
    if (columnIndex <= 0 || columnIndex > 16384) {
        throw new Error('请确保 1 <= columnIndex <= 16384')
    }
    
    const address = ActiveSheet.Columns.Item(columnIndex).Address()
    return address.substr(1, address.indexOf(':') - 1)
}
```

## 数组操作模板

### 基础数组读写

```javascript
/**
 * 使用数组批量处理数据
 */
function ProcessWithArray(ws) {
    const wsData = ws || ActiveSheet
    const lastRow = GetLastRow(wsData, 1)
    
    if (lastRow < 2) return
    
    // 读取到数组（单次IO）
    const data = wsData.Range("A2:D" + lastRow).Value2
    
    // 内存中处理
    for (let i = 0; i < data.length; i++) {
        data[i][3] = data[i][1] * data[i][2]  // 金额 = 数量 * 单价
    }
    
    // 一次性写回（单次IO）
    wsData.Range("A2:D" + lastRow).Value2 = data
}
```

### 二维数组处理

```javascript
function Process2DArray(ws) {
    const wsData = ws || ActiveSheet
    const data = wsData.Range("A1:D10").Value2
    
    for (let rowIdx = 0; rowIdx < data.length; rowIdx++) {
        for (let colIdx = 0; colIdx < data[rowIdx].length; colIdx++) {
            // 处理每个单元格
            if (typeof data[rowIdx][colIdx] === 'string') {
                data[rowIdx][colIdx] = data[rowIdx][colIdx].toUpperCase()
            }
        }
    }
    
    wsData.Range("A1:D10").Value2 = data
}
```

### 分块处理大数组

```javascript
function ProcessLargeArrayInChunks(ws) {
    const wsData = ws || ActiveSheet
    const lastRow = GetLastRow(wsData, 1)
    const chunkSize = 5000  // 每块5000行
    
    for (let startRow = 2; startRow <= lastRow; startRow += chunkSize) {
        const endRow = Math.min(startRow + chunkSize - 1, lastRow)
        
        const data = wsData.Range(`A${startRow}:D${endRow}`).Value2
        
        // 处理当前块
        for (let i = 0; i < data.length; i++) {
            data[i][3] = data[i][1] * data[i][2]
        }
        
        wsData.Range(`A${startRow}:D${endRow}`).Value2 = data
        
        console.log(`已处理: ${startRow} - ${endRow}`)
    }
}
```

## 单元格格式设置模板

### 字体和颜色

```javascript
function SetFontAndColor(ws) {
    const wsData = ws || ActiveSheet
    const rng = wsData.Range("A1:D10")
    
    // 字体设置
    rng.Font.Name = "微软雅黑"
    rng.Font.Size = 11
    rng.Font.Bold = true
    rng.Font.Color = 255  // 红色
    
    // 背景色
    rng.Interior.Color = 65535  // 黄色
    
    // 边框
    rng.Borders.LineStyle = xlContinuous
    rng.Borders.Weight = xlThin
    
    // 对齐
    rng.HorizontalAlignment = xlHAlignCenter
    rng.VerticalAlignment = xlVAlignCenter
    rng.WrapText = true
}
```

### 设置边框

```javascript
function MakeBorders(ws) {
    const wsData = ws || ActiveSheet
    const rng = wsData.Range("A1:D10")
    
    const borderIndices = [
        xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight,
        xlInsideHorizontal, xlInsideVertical
    ]
    
    for (const idx of borderIndices) {
        const border = rng.Borders.Item(idx)
        border.Weight = xlThin
        border.LineStyle = xlContinuous
        border.Color = 0  // 黑色
    }
}
```

### 数字格式

```javascript
function SetNumberFormats(ws) {
    const wsData = ws || ActiveSheet
    
    wsData.Range("A1").NumberFormat = "@"              // 文本
    wsData.Range("A2").NumberFormat = "0.00"           // 数值（两位小数）
    wsData.Range("A3").NumberFormat = "yyyy/m/d"       // 日期
    wsData.Range("A4").NumberFormat = "0.00%"          // 百分比
    wsData.Range("A5").NumberFormat = "￥#,##0.00"     // 货币
}
```

## 工作表操作模板

### 创建工作表

```javascript
function CreateWorksheet(sheetName) {
    let ws = null
    
    // 检查是否已存在
    try {
        ws = ThisWorkbook.Worksheets.Item(sheetName)
    } catch (e) {
        // 不存在，创建新的
        ws = ThisWorkbook.Worksheets.Add()
        ws.Name = sheetName
    }
    
    // 清空内容
    ws.Cells.Clear()
    return ws
}
```

### 复制工作表

```javascript
function CopyWorksheet(sourceName, newName) {
    const ws = ThisWorkbook.Worksheets.Item(sourceName)
    ws.Copy(null, ThisWorkbook.Worksheets.Item(ThisWorkbook.Worksheets.Count))
    ActiveSheet.Name = newName
}
```

### 删除工作表

```javascript
function DeleteWorksheet(sheetName) {
    Application.DisplayAlerts = false
    try {
        ThisWorkbook.Worksheets.Item(sheetName).Delete()
    } catch (e) {
        console.log(`删除失败: ${e.message}`)
    }
    Application.DisplayAlerts = true
}
```

## 查找和替换模板

### 查找

```javascript
function FindInColumn(ws, col, searchValue) {
    const wsData = ws || ActiveSheet
    const rng = wsData.Columns(col)
    
    let found = rng.Find(searchValue)
    const results = []
    
    if (found) {
        const firstAddr = found.Address()
        results.push(found)
        
        found = rng.Find(searchValue, found)
        while (found && found.Address() !== firstAddr) {
            results.push(found)
            found = rng.Find(searchValue, found)
        }
    }
    
    return results
}

function FindAll(ws, searchValue) {
    const results = Cells.FindAll(searchValue)
    return results || []
}
```

### 替换

```javascript
function ReplaceValue(ws, oldValue, newValue) {
    const wsData = ws || ActiveSheet
    wsData.Cells.Replace(oldValue, newValue)
}
```

## 排序和筛选模板

### 排序

```javascript
function SortRange(ws) {
    const wsData = ws || ActiveSheet
    const rng = wsData.Range("A1:D100")
    
    // 清除现有排序
    wsData.Sort.SortFields.Clear()
    
    // 添加排序条件
    wsData.Sort.SortFields.Add(
        wsData.Range("A2"),  // 关键字段
        xlSortOnValues,
        xlAscending  // 升序
    )
    
    // 应用排序
    wsData.Sort.SetRange(rng)
    wsData.Sort.Header = xlYes
    wsData.Sort.Apply()
}
```

### 筛选

```javascript
function ApplyFilter(ws) {
    const wsData = ws || ActiveSheet
    
    // 清除现有筛选
    if (wsData.AutoFilterMode) {
        wsData.AutoFilterMode = false
    }
    
    // 应用筛选
    wsData.Range("A1:D100").AutoFilter(1, ">100")  // 第1列大于100
}

function ClearFilter(ws) {
    const wsData = ws || ActiveSheet
    if (wsData.AutoFilterMode) {
        wsData.ShowAllData()
    }
}
```

## 合并单元格模板

```javascript
function MergeCells(ws) {
    const wsData = ws || ActiveSheet
    
    // 合并
    wsData.Range("A1:D1").Merge()
    
    // 设置对齐
    wsData.Range("A1:D1").HorizontalAlignment = xlHAlignCenter
    wsData.Range("A1:D1").VerticalAlignment = xlVAlignCenter
}

function UnMergeCells(ws) {
    const wsData = ws || ActiveSheet
    wsData.Range("A1:D1").UnMerge()
}

function IsMerged(ws) {
    const wsData = ws || ActiveSheet
    return wsData.Range("A1").MergeCells
}
```

## 进度显示模板

```javascript
function ShowProgress(current, total) {
    const percent = Math.floor((current / total) * 100)
    console.log(`处理进度: ${current}/${total} (${percent}%)`)
}

// 使用示例
function ProcessWithProgress(ws) {
    const wsData = ws || ActiveSheet
    const lastRow = GetLastRow(wsData, 1)
    const data = wsData.Range("A2:D" + lastRow).Value2
    
    for (let i = 0; i < data.length; i++) {
        // 处理逻辑
        data[i][3] = data[i][1] * data[i][2]
        
        // 每100行显示一次进度
        if (i % 100 === 0) {
            ShowProgress(i + 1, data.length)
        }
    }
    
    wsData.Range("A2:D" + lastRow).Value2 = data
    console.log("处理完成！")
}
```

## 数据验证模板

```javascript
function AddValidation(ws) {
    const wsData = ws || ActiveSheet
    const rng = wsData.Range("B2:B100")
    
    // 删除现有验证
    rng.Validation.Delete()
    
    // 添加列表验证
    rng.Validation.Add(
        xlValidateList,
        xlValidAlertStop,
        xlBetween,
        "男,女,其他",
        undefined
    )
    
    // 设置提示信息
    rng.Validation.InputTitle = "性别"
    rng.Validation.InputMessage = "请选择：男,女,其他"
    rng.Validation.ErrorTitle = "输入错误"
    rng.Validation.ErrorMessage = "请从下拉列表中选择"
}
```

## 条件格式模板

```javascript
function AddConditionFormat(ws) {
    const wsData = ws || ActiveSheet
    const rng = wsData.Range("A2:A100")
    
    // 清除现有条件格式
    rng.FormatConditions.Delete()
    
    // 添加条件格式（大于60时变红）
    const fc = rng.FormatConditions.Add(
        xlCellValue,
        xlGreater,
        "60",
        undefined,
        undefined,
        undefined,
        undefined
    )
    
    fc.Font.Color = 255  // 红色字体
    fc.Interior.Color = 65535  // 黄色背景
}
```

## 图表创建模板

```javascript
function CreateChart(ws) {
    const wsData = ws || ActiveSheet
    const chart = wsData.Shapes.AddChart(xlColumnClustered)
    
    chart.Chart.SetSourceData(wsData.Range("A1:B10"))
    chart.Chart.HasTitle = true
    chart.Chart.ChartTitle.Text = "数据图表"
}
```

## 类封装模板

```javascript
/**
 * 数据表处理类
 */
class DataTableProcessor {
    constructor(ws) {
        this.ws = ws || ActiveSheet
        this.lastRow = 0
        this.lastCol = 0
        this.data = null
    }
    
    init() {
        this.lastRow = GetLastRow(this.ws, 1)
        this.lastCol = GetLastCol(this.ws, 1)
        console.log(`数据范围: ${this.lastRow}行 x ${this.lastCol}列`)
        return this
    }
    
    readData() {
        if (this.lastRow < 2) return null
        this.data = this.ws.Range("A1").Resize(this.lastRow, this.lastCol).Value2
        return this
    }
    
    writeData() {
        if (!this.data) return this
        this.ws.Range("A1").Resize(this.data.length, this.data[0].length).Value2 = this.data
        return this
    }
    
    process(processor) {
        if (!this.data) return this
        
        for (let i = 0; i < this.data.length; i++) {
            processor(this.data[i], i)
        }
        return this
    }
    
    formatHeader() {
        if (this.lastRow < 1) return this
        const header = this.ws.Range("A1").Resize(1, this.lastCol)
        header.Font.Bold = true
        header.Interior.Color = 12611584  // 蓝色
        header.Font.Color = 16777215  // 白色
        return this
    }
}

// 使用示例
new DataTableProcessor()
    .init()
    .readData()
    .process((row, idx) => {
        if (idx > 0 && row[1] && row[2]) {
            row[3] = row[1] * row[2]
        }
    })
    .writeData()
    .formatHeader()
```

## 工具函数集合

### RGB颜色计算

```javascript
/**
 * RGB转十进制颜色值
 */
function RGB(r, g, b) {
    return r + g * 256 + b * 256 * 256
}

// 使用示例
const red = RGB(255, 0, 0)       // 255
const blue = RGB(0, 0, 255)      // 16711680
const green = RGB(0, 128, 0)     // 32768
```

### 判断工作表是否存在

```javascript
function SheetExists(sheetName) {
    try {
        const ws = ThisWorkbook.Worksheets.Item(sheetName)
        return true
    } catch (e) {
        return false
    }
}
```

### 判断范围是否为空

```javascript
function IsRangeEmpty(rng) {
    // 检查是否有非空单元格
    const data = rng.Value2
    if (!data) return true
    
    if (Array.isArray(data)) {
        return data.every(row => row.every(cell => cell === undefined || cell === null || cell === ''))
    }
    return data === undefined || data === null || data === ''
}
```
