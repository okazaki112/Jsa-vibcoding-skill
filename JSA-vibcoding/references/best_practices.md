# JSA 最佳实践

代码规范、性能优化和调试技巧。

## 代码规范

### 变量声明

```javascript
// ✅ 使用 const/let（推荐）
const lastRow = GetLastRow(ActiveSheet, 1)
let data = Range("A1:D" + lastRow).Value2

// ❌ 避免使用 var
var oldStyle = "不推荐"

// ✅ 有意义的命名
const customerName = "张三"
const totalAmount = 1000
const wsData = ActiveSheet

// ❌ 避免无意义命名
const x = "张三"
const temp = 1000
```

### 命名规范

| 类型 | 规则 | 示例 |
|-----|------|------|
| 常量 | 全大写下划线 | `MAX_COUNT`, `DEFAULT_NAME`, `PI_VALUE` |
| 变量 | 驼峰命名 | `lastRow`, `fileName`, `totalAmount` |
| 函数 | 动词+名词 | `getData`, `processRecords`, `calculateSum` |
| 类 | 帕斯卡命名 | `DataProcessor`, `TableManager` |
| 私有方法 | 下划线前缀 | `_internalProcess`, `_validateInput` |

### 函数设计

```javascript
// ✅ 单一职责
function GetLastRow(ws, col) {
    const lastCell = ws.Columns(col).Find("*", undefined, undefined, undefined, xlByRows, xlPrevious)
    return lastCell ? lastCell.Row : 0
}

// ✅ 有文档注释
/**
 * 获取指定列的最后一行
 * @param {Worksheet} ws - 工作表对象
 * @param {number} col - 列号（1-based）
 * @returns {number} 最后一行号，无数据返回0
 */
function GetLastRow(ws, col) {
    // ...
}

// ✅ 参数校验
function ProcessData(ws, startRow, endRow) {
    if (!ws) {
        throw new Error("工作表不能为空")
    }
    if (startRow < 1 || endRow < startRow) {
        throw new Error("行范围无效")
    }
    // ...
}

// ✅ 返回明确的结果
function FindValue(ws, searchValue) {
    const rng = ws.Cells.Find(searchValue)
    return rng || null  // 明确返回 null 而不是 undefined
}
```

### 错误处理

```javascript
// ✅ 标准 try-catch 结构
function SafeProcess(ws) {
    try {
        // 主逻辑
        const data = ReadData(ws)
        ProcessData(data)
        WriteData(ws, data)
    } catch (e) {
        console.log(`错误: ${e.message}`)
        alert(`处理失败: ${e.message}`)
    } finally {
        // 清理资源
        console.log("处理结束")
    }
}

// ✅ 分层错误处理
function Main() {
    try {
        ProcessWorkbook()
    } catch (e) {
        console.log(`主流程错误: ${e.message}`)
        throw e  // 向上抛出
    }
}

function ProcessWorkbook() {
    try {
        ProcessSheet()
    } catch (e) {
        console.log(`工作表处理错误: ${e.message}`)
        throw new Error(`工作表处理失败: ${e.message}`)
    }
}
```

## 性能优化

### 数组批量操作

```javascript
// ❌ 慢：逐单元格操作
function SlowProcess(ws) {
    for (let i = 1; i <= 10000; i++) {
        ws.Cells(i, 1).Value2 = ws.Cells(i, 1).Value2 * 2
    }
}

// ✅ 快：数组批量操作
function FastProcess(ws) {
    const data = ws.Range("A1:A10000").Value2
    for (let i = 0; i < data.length; i++) {
        data[i][0] = data[i][0] * 2
    }
    ws.Range("A1:A10000").Value2 = data
}
```

### 缓存对象引用

```javascript
// ❌ 重复获取对象
function BadExample() {
    for (let i = 1; i <= 100; i++) {
        ActiveSheet.Cells(i, 1).Value2 = i
        ActiveSheet.Cells(i, 2).Value2 = i * 2
        ActiveSheet.Cells(i, 3).Value2 = i * 3
    }
}

// ✅ 缓存引用
function GoodExample() {
    const ws = ActiveSheet
    for (let i = 1; i <= 100; i++) {
        ws.Cells(i, 1).Value2 = i
        ws.Cells(i, 2).Value2 = i * 2
        ws.Cells(i, 3).Value2 = i * 3
    }
}
```

### 减少选择操作

```javascript
// ❌ 频繁选择
function BadSelect() {
    Range("A1").Select()
    Selection.Font.Bold = true
    Range("A2").Select()
    Selection.Font.Bold = true
}

// ✅ 直接操作
function GoodDirect() {
    Range("A1").Font.Bold = true
    Range("A2").Font.Bold = true
}

// ✅ 更好：批量操作
function BestBatch() {
    Range("A1:A2").Font.Bold = true
}
```

### 分块处理大数据

```javascript
function ProcessLargeData(ws, chunkSize = 5000) {
    const lastRow = GetLastRow(ws, 1)
    
    for (let startRow = 2; startRow <= lastRow; startRow += chunkSize) {
        const endRow = Math.min(startRow + chunkSize - 1, lastRow)
        const data = ws.Range(`A${startRow}:D${endRow}`).Value2
        
        // 处理当前块
        for (let i = 0; i < data.length; i++) {
            data[i][3] = data[i][1] * data[i][2]
        }
        
        // 写回
        ws.Range(`A${startRow}:D${endRow}`).Value2 = data
        
        // 进度显示
        console.log(`已处理: ${startRow} - ${endRow}`)
    }
}
```

### 避免不必要的类型转换

```javascript
// ❌ 重复转换
function BadConvert() {
    for (let i = 1; i <= 100; i++) {
        const value = String(ws.Cells(i, 1).Value2)
        const num = Number(value)
        // ...
    }
}

// ✅ 一次转换
function GoodConvert() {
    const data = ws.Range("A1:A100").Value2
    for (let i = 0; i < data.length; i++) {
        const value = Number(data[i][0])
        // ...
    }
}
```

## 调试技巧

### console.log 最佳实践

```javascript
// ✅ 带时间戳
console.log(`[${new Date().toLocaleTimeString()}] 开始处理`)

// ✅ 带上下文
console.log(`[数据处理] 读取到 ${data.length} 行`)

// ✅ 表格格式
console.log("序号\t姓名\t金额")
data.forEach((row, i) => {
    console.log(`${i + 1}\t${row[0]}\t${row[1]}`)
})

// ✅ 对象输出
console.log(JSON.stringify(obj, null, 2))

// ✅ 性能计时
const start = Date.now()
// ... 处理逻辑
console.log(`耗时: ${Date.now() - start}ms`)
```

### 调试函数模板

```javascript
function DebugProcess(ws) {
    console.log("=".repeat(50))
    console.log(`开始时间: ${new Date().toLocaleString()}`)
    console.log(`工作表: ${ws.Name}`)
    console.log("=".repeat(50))
    
    const startTime = Date.now()
    
    try {
        // 步骤1
        console.log("\n[步骤1] 读取数据")
        const data = ReadData(ws)
        console.log(`  数据行数: ${data.length}`)
        
        // 步骤2
        console.log("\n[步骤2] 处理数据")
        const result = ProcessData(data)
        console.log(`  处理结果: ${result.length} 行`)
        
        // 步骤3
        console.log("\n[步骤3] 写入结果")
        WriteData(ws, result)
        
        const elapsed = ((Date.now() - startTime) / 1000).toFixed(2)
        console.log("\n" + "=".repeat(50))
        console.log(`处理完成，总耗时: ${elapsed}秒`)
        
    } catch (e) {
        console.log(`\n❌ 错误: ${e.message}`)
        console.log(`堆栈: ${e.stack}`)
    }
}
```

### 条件断点模拟

```javascript
function ProcessWithBreakpoint(ws) {
    const breakCondition = { row: 100, value: "特定值" }
    
    const lastRow = GetLastRow(ws, 1)
    const data = ws.Range("A2:D" + lastRow).Value2
    
    for (let i = 0; i < data.length; i++) {
        // 模拟条件断点
        if (i + 2 === breakCondition.row) {
            console.log(`[断点] 行 ${i + 2}:`)
            console.log(`  数据: ${JSON.stringify(data[i])}`)
            // 在这里可以手动检查变量
        }
        
        if (data[i][0] === breakCondition.value) {
            console.log(`[断点] 找到目标值在第 ${i + 2} 行`)
        }
        
        // 处理逻辑
    }
}
```

## 安全实践

### 参数验证

```javascript
function SafeFunction(ws, range, options) {
    // 验证必填参数
    if (!ws) {
        throw new Error("工作表参数不能为空")
    }
    
    // 验证类型
    if (typeof range !== 'string') {
        throw new Error("range 参数必须是字符串")
    }
    
    // 验证选项
    const defaultOptions = {
        skipEmpty: true,
        maxRows: 10000
    }
    const opts = { ...defaultOptions, ...options }
    
    // 验证范围
    if (opts.maxRows < 1 || opts.maxRows > 100000) {
        throw new Error("maxRows 范围无效")
    }
    
    // 主逻辑
}
```

### 危险操作确认

```javascript
function DangerousOperation(ws) {
    // 确认操作
    const result = MsgBox(
        "此操作将删除所有数据，是否继续？",
        jsYesNo,
        "警告"
    )
    
    if (result !== 6) {  // 不是"是"
        console.log("操作已取消")
        return
    }
    
    // 执行危险操作
    try {
        ws.Cells.Clear()
        console.log("操作完成")
    } catch (e) {
        console.log(`操作失败: ${e.message}`)
    }
}
```

### 数据备份

```javascript
function ProcessWithBackup(ws) {
    // 创建备份
    const backupName = `${ws.Name}_备份_${Date.now()}`
    ws.Copy(null, ThisWorkbook.Sheets.Item(ThisWorkbook.Sheets.Count))
    ActiveSheet.Name = backupName
    
    // 激活原表
    ws.Activate()
    
    try {
        // 主处理逻辑
        ProcessData(ws)
        console.log("处理完成")
    } catch (e) {
        // 恢复备份
        console.log(`处理失败，正在恢复备份: ${backupName}`)
        ThisWorkbook.Sheets.Item(backupName).Copy(null, ws)
        ThisWorkbook.Sheets.Item(backupName).Delete()
    }
}
```

## 代码组织

### 模块化设计

```javascript
// ===== 工具函数模块 =====
const Utils = {
    getLastRow: function(ws, col) {
        const lastCell = ws.Columns(col).Find("*", undefined, undefined, undefined, xlByRows, xlPrevious)
        return lastCell ? lastCell.Row : 0
    },
    
    getLastCol: function(ws, row) {
        const lastCell = ws.Rows(row).Find("*", undefined, undefined, undefined, xlByColumns, xlPrevious)
        return lastCell ? lastCell.Column : 0
    },
    
    RGB: function(r, g, b) {
        return r + g * 256 + b * 256 * 256
    }
}

// ===== 数据处理模块 =====
const DataProcessor = {
    read: function(ws, range) {
        return ws.Range(range).Value2
    },
    
    write: function(ws, range, data) {
        ws.Range(range).Value2 = data
    },
    
    process: function(data, processor) {
        return data.map((row, i) => processor(row, i))
    }
}

// 使用示例
function Main() {
    const lastRow = Utils.getLastRow(ActiveSheet, 1)
    const data = DataProcessor.read(ActiveSheet, `A1:D${lastRow}`)
    DataProcessor.process(data, (row, i) => {
        if (i > 0) row[3] = row[1] * row[2]
        return row
    })
    DataProcessor.write(ActiveSheet, `A1:D${lastRow}`, data)
}
```

### 类设计

```javascript
/**
 * 数据表处理器
 */
class TableProcessor {
    constructor(ws) {
        this.ws = ws || ActiveSheet
        this.data = null
        this.lastRow = 0
    }
    
    // 初始化
    init() {
        this.lastRow = this._getLastRow()
        return this
    }
    
    // 读取数据
    read() {
        if (this.lastRow < 2) return this
        this.data = this.ws.Range(`A1:D${this.lastRow}`).Value2
        return this
    }
    
    // 处理数据
    process(fn) {
        if (!this.data) return this
        this.data = this.data.map((row, i) => fn(row, i))
        return this
    }
    
    // 写入数据
    write() {
        if (!this.data) return this
        this.ws.Range(`A1:D${this.lastRow}`).Value2 = this.data
        return this
    }
    
    // 私有方法
    _getLastRow() {
        const lastCell = this.ws.Columns(1).Find("*", undefined, undefined, undefined, xlByRows, xlPrevious)
        return lastCell ? lastCell.Row : 0
    }
}

// 使用
new TableProcessor()
    .init()
    .read()
    .process((row, i) => {
        if (i > 0 && row[1] && row[2]) {
            row[3] = row[1] * row[2]
        }
        return row
    })
    .write()
```

## 检查清单

### 代码审查清单

- [ ] 使用 `const`/`let` 而非 `var`
- [ ] 变量命名有意义
- [ ] 函数有文档注释
- [ ] 参数有验证
- [ ] 有错误处理
- [ ] 无硬编码值

### 性能检查清单

- [ ] 使用数组批量读写
- [ ] 缓存常用对象引用
- [ ] 减少选择操作
- [ ] 大数据分块处理
- [ ] 避免循环内重复计算

### 安全检查清单

- [ ] 危险操作有确认
- [ ] 关键数据有备份
- [ ] 参数有验证
- [ ] 错误有处理
- [ ] 敏感信息不打印
