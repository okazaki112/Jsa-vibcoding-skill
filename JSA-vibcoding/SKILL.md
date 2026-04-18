---
name: JSA-vibcoding
description: WPS JSA 智能编程助手 - 基于 JavaScript 的 WPS 宏开发。通过自然语言描述需求，自动生成高质量 JSA 代码。支持单元格操作、工作表管理、数据处理、事件处理等完整开发流程。支持 ES6-ES2019 语法，提供现代 JavaScript 编程体验。
---

# JSA VibCoding - WPS智能编程助手

基于 JavaScript 的 WPS 宏开发解决方案：高质量代码生成、现代 JS 语法、完整 API 支持。

## 核心能力

```
┌─────────────────────────────────────────────────────────────────────┐
│  自然语言需求 → 高质量JSA代码生成 → WPS宏编辑器运行 → 执行验证        │
│         ↑                                              ↓            │
│         └──────────── 调试/优化/扩展 ←─────────────────┘            │
└─────────────────────────────────────────────────────────────────────┘
```

## JSA 简介

从 WPS 2021 版本开始，WPS 正式支持使用 JavaScript 作为宏语言，称为 JSA (JS for Application)。

**核心特性：**
- 内嵌 V8 引擎，支持 ES6-ES2019 语法
- 支持 JavaScript 标准内置对象
- 不支持浏览器内置对象（如 window、document）
- API 与 VBA 高度兼容，索引从 1 开始

**JSA 优势：**
- 基于 JavaScript，语言特性现代化
- 背靠 JS 活跃社区，资源丰富
- 跨平台能力
- 安全性好，无外部库依赖

## 快速开始

### 启用 JS 宏

1. 点击【文件】→【选项】
2. 勾选【默认 JS 开发环境】
3. 按 `Alt+F11` 打开宏编辑器

### 基础示例

```javascript
// 基础单元格操作
function 单元格操作() {
    // 读取单元格
    let value = Range("A1").Value2
    console.log("A1的值: " + value)
    
    // 写入单元格
    Range("A1").Value2 = "Hello JSA"
    
    // 批量写入
    Range("A1:C1").Value2 = ["姓名", "年龄", "城市"]
    
    // 写入二维数组
    let data = [
        ["张三", 25, "北京"],
        ["李四", 30, "上海"]
    ]
    Range("A2:C3").Value2 = data
}

// 选择工作表
function 选择工作表() {
    Sheets.Item(1).Select()           // 按索引
    Sheets.Item("Sheet2").Select()    // 按名称
}

// 消息框和输入框
function 用户交互() {
    // 消息框
    let result = MsgBox("确定继续吗?", jsYesNo, "确认")
    if (result == 6) {
        alert("点击了是")
    }
    
    // 输入框
    let name = InputBox("请输入姓名", "提示", "默认值")
    alert("你输入了: " + name)
}
```

## JSA 与 VBA 差异对照

| 特性 | VBA | JSA |
|------|-----|-----|
| 有参数属性 | `Range("A1").Font` | `Range("A1").Font` 无变化 |
| 省略括号 | 可以省略 | **必须加括号** |
| 默认值参数 | 可省略 | **必须写 undefined** |
| 命名参数 | 支持 | **不支持**，按顺序传递 |
| 变量声明 | `Dim x As Long` | `let x = 0` |
| 循环 | `For i = 1 To 10` | `for (let i = 1; i <= 10; i++)` |
| 数组 | `arr(0)` 0-based | `arr[0]` 0-based |

### 参数传递示例

```javascript
// VBA: 可省略参数
// Range("A1").AutoFilter Field:=1, Criteria1:="苹果"

// JSA: 必须完整传递
Range("A1").AutoFilter(1, "苹果", undefined, undefined, undefined)
```

## JSA 代码规范

### 1. 现代JavaScript风格

```javascript
// ✅ 使用 const/let 替代 var
const lastRow = GetLastRow(ActiveSheet, 1)
let data = Range("A1:D" + lastRow).Value2

// ✅ 使用箭头函数
const double = x => x * 2
arr.forEach(item => console.log(item))

// ✅ 使用模板字符串
console.log(`处理第 ${i} 行，值: ${value}`)

// ✅ 使用解构赋值
const { Count: rowCount } = Sheets
const [first, ...rest] = arr

// ✅ 使用扩展运算符
const newArr = [...oldArr, newItem]

// ✅ 使用类和模块化
class DataProcessor {
    constructor(ws) {
        this.ws = ws
    }
    
    process() {
        // 处理逻辑
    }
}
```

### 2. 错误处理

```javascript
function 安全处理() {
    try {
        // 主代码
        ProcessData()
    } catch (e) {
        console.log("错误: " + e.message)
        alert("处理失败: " + e.message)
    } finally {
        // 清理代码
    }
}
```

### 3. 性能优化

```javascript
function 性能优化示例() {
    // 使用数组批量处理
    let lastRow = GetLastRow(ActiveSheet, 1)
    let data = Range("A2:D" + lastRow).Value2
    
    // 内存中处理（快）
    for (let i = 0; i < data.length; i++) {
        data[i][3] = data[i][1] * data[i][2]  // 金额 = 数量 * 单价
    }
    
    // 一次性写回
    Range("A2:D" + lastRow).Value2 = data
}
```

## 常用API速查

### Range对象

```javascript
// 值操作
Range("A1").Value2              // 读取值
Range("A1").Value2 = "数据"      // 写入值
Range("A1").Formula             // 读取公式
Range("A1").Formula = "=B1+C1"  // 写入公式
Range("A1").Text                // 格式化文本（只读）

// 选择和定位
Range("A1").Select()            // 选择
Range("A1").Activate()          // 激活
Range("A1").Offset(2, 3)        // 偏移
Range("A1").Resize(5, 3)        // 调整大小
Range("A1").EntireRow           // 整行
Range("A1").EntireColumn        // 整列

// 格式设置
Range("A1").Font.Name = "微软雅黑"
Range("A1").Font.Size = 12
Range("A1").Font.Bold = true
Range("A1").Font.Color = 255    // 红色
Range("A1").Interior.Color = 65535  // 黄色背景
Range("A1").NumberFormat = "@"  // 文本格式

// 合并
Range("A1:D1").Merge()
Range("A1:D1").UnMerge()
Range("A1").MergeCells          // 是否合并

// 清除
Range("A1").ClearContents()     // 清除内容
Range("A1").ClearFormats()      // 清除格式
Range("A1").Clear()             // 清除全部

// 行列信息
Range("A1").Row                 // 行号
Range("A1").Column              // 列号
Range("A1").Count               // 单元格数
```

### Sheets/Worksheets

```javascript
// 选择
Sheets.Item(1).Select()         // 按索引
Sheets.Item("Sheet1").Select()  // 按名称
Sheets.Count                    // 工作表数量

// 新建
let ws = Worksheets.Add()
ws.Name = "新工作表"

// 在最后新建
let ws = Worksheets.Add(null, Sheets(Sheets.Count))

// 删除
Application.DisplayAlerts = false
Sheets.Item("Sheet2").Delete()
Application.DisplayAlerts = true

// 复制/移动
Sheets.Item("Sheet1").Copy(null, Sheets.Item(Sheets.Count))
Sheets.Item("Sheet1").Move(null, Sheets.Item(Sheets.Count))
```

### Workbook

```javascript
// 当前工作簿
ThisWorkbook.Save()
ThisWorkbook.Close()

// 活动工作表
ActiveSheet.Name

// 打开/新建
Workbooks.Open("D:\\test.xlsx")
Workbooks.Add()
```

### 查找替换

```javascript
// 查找
let rng = Cells.Find("搜索内容")
if (rng) {
    rng.Select()
}

// 查找所有
let results = Cells.FindAll("搜索内容")
for (let addr of results) {
    Range(addr).Select()
}

// 替换
Cells.Replace("旧内容", "新内容")
```

## 事件处理

JSA 通过特定函数名实现事件处理：

```javascript
// 单元格值改变
function Workbook_SheetChange(Sh, rg) {
    console.log(`工作表: ${Sh.Name}`)
    console.log(`改变的单元格: ${rg.Address()}`)
}

// 工作表激活
function Workbook_SheetActivate(Sh) {
    console.log(`激活了工作表: ${Sh.Name}`)
}

// 工作表选择改变
function Workbook_SheetSelectionChange(Sh, rg) {
    console.log(`选中了: ${rg.Address()}`)
}

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
    console.log("即将关闭")
}
```

## 自定义公式函数

```javascript
/**
 * 删除给定文本的某字符及后面的字符
 * @param {Range|string} target - 要处理的目标
 * @param {Range|string} from - 要删除的起始字符
 */
function DeleteFrom(target, from) {
    let value
    if (target.constructor.name == 'Range') {
        if (target.Areas.Count > 1 || target.Areas.Item(1).Cells.Count > 1) {
            throw new Error('本函数只能处理一个单元格')
        }
        value = target.Value().toString()
    } else {
        value = target.toString()
    }
    
    if (from.constructor.name == 'Range') {
        from = from.Value().toString()
    } else {
        from = from.toString()
    }
    
    let index = value.indexOf(from)
    return index == -1 ? value : value.substr(0, index)
}

// 在单元格中使用: =DeleteFrom(A1, "@")
```

## 调试技巧

```javascript
// 控制台输出（立即窗口）
console.log("调试信息")
console.log(`变量值: ${value}`)

// 清除立即窗口
Console.clear()

// Debug.Print（同 console.log）
Debug.Print("调试信息")

// 垃圾回收
Debug.GC()
```

## WebOffice SDK（在线版）

JSA 也支持 WebOffice 在线环境：

```javascript
// 初始化
const instance = WebOfficeSDK.init({
    officeType: WebOfficeSDK.OfficeType.Spreadsheet,
    appId: 'your-app-id',
    fileId: 'your-file-id'
})

// 等待就绪
await instance.ready()

// 获取 Application
const app = instance.Application

// 操作单元格
app.Range('A1').Value = 'Hello WebOffice'
```

## 标准代码模板

### 完整数据处理模板

```javascript
/**
 * 数据处理主函数
 * @param {Worksheet} ws - 目标工作表（可选）
 */
function 数据处理(ws) {
    // 使用传入的工作表或当前活动表
    const wsData = ws || ActiveSheet
    const startTime = Date.now()
    
    try {
        console.log(`开始处理: ${wsData.Name}`)
        
        // 获取数据边界
        const lastRow = GetLastRow(wsData, 1)
        if (lastRow < 2) {
            alert("数据不足")
            return
        }
        
        // 读取数据到数组
        const data = wsData.Range("A2:D" + lastRow).Value2
        
        // 内存处理
        for (let i = 0; i < data.length; i++) {
            if (data[i][1] && data[i][2]) {
                data[i][3] = data[i][1] * data[i][2]  // 金额
            }
        }
        
        // 写回结果
        wsData.Range("A2:D" + lastRow).Value2 = data
        
        const elapsed = ((Date.now() - startTime) / 1000).toFixed(2)
        alert(`处理完成！用时: ${elapsed}秒`)
        
    } catch (e) {
        console.log(`错误: ${e.message}`)
        alert(`处理失败: ${e.message}`)
    }
}

/**
 * 获取指定列的最后一行
 */
function GetLastRow(ws, col) {
    const lastCell = ws.Columns(col).Find("*", undefined, undefined, undefined, xlByRows, xlPrevious)
    return lastCell ? lastCell.Row : 0
}

/**
 * 获取指定行的最后一列
 */
function GetLastCol(ws, row) {
    const lastCell = ws.Rows(row).Find("*", undefined, undefined, undefined, xlByColumns, xlPrevious)
    return lastCell ? lastCell.Column : 0
}
```

### 类封装示例

```javascript
/**
 * 数据表处理类
 */
class DataTableProcessor {
    constructor(ws) {
        this.ws = ws || ActiveSheet
        this.lastRow = 0
        this.lastCol = 0
    }
    
    init() {
        this.lastRow = GetLastRow(this.ws, 1)
        this.lastCol = GetLastCol(this.ws, 1)
        console.log(`数据范围: ${this.lastRow}行 x ${this.lastCol}列`)
        return this
    }
    
    readData() {
        if (this.lastRow < 2) return null
        return this.ws.Range("A1").Resize(this.lastRow, this.lastCol).Value2
    }
    
    writeData(data) {
        this.ws.Range("A1").Resize(data.length, data[0].length).Value2 = data
        return this
    }
    
    process(processor) {
        const data = this.readData()
        if (!data) return this
        
        for (let i = 0; i < data.length; i++) {
            processor(data[i], i)
        }
        
        this.writeData(data)
        return this
    }
}

// 使用示例
new DataTableProcessor()
    .init()
    .process((row, idx) => {
        if (idx > 0) {  // 跳过标题行
            row[3] = row[1] * row[2]
        }
    })
```

## 知识库导航

详细知识库在 `references/` 目录：

| 文件 | 内容 |
|------|------|
| **code_templates.md** | 完整代码模板库（数据处理、表格操作、图表生成） |
| **jsa_patterns.md** | JSA 常用模式（单元格、工作表、事件处理） |
| **best_practices.md** | 最佳实践清单（命名规范、性能优化、调试技巧） |
| **enum_constants.md** | 枚举常量速查表（对齐、边框、颜色、格式等） |

## 触发关键词

**代码生成**: "JSA代码"、"WPS宏"、"JS宏"、"WPS自动化"

**功能需求**: "数据处理"、"单元格操作"、"工作表管理"、"图表生成"

**调优**: "优化JSA性能"、"调试代码"、"数组处理"

## 最佳实践总结

### 命名规范
| 类型 | 规则 | 示例 |
|-----|------|------|
| 常量 | 全大写下划线 | `MAX_COUNT`, `DEFAULT_NAME` |
| 变量 | 驼峰命名 | `lastRow`, `fileName`, `totalAmount` |
| 函数 | 动词+名词 | `getData`, `processRecords`, `calculateSum` |
| 类 | 帕斯卡命名 | `DataProcessor`, `TableManager` |

### 性能检查清单
- [ ] 使用数组批量读写
- [ ] 避免循环中重复获取属性
- [ ] 使用 `const` 缓存常用对象
- [ ] 减少单元格选择操作
- [ ] 使用模板字符串拼接

### 安全实践
```javascript
// 操作前检查
function 安全操作(ws) {
    const wsData = ws || ActiveSheet
    
    if (!wsData) {
        alert("未选择工作表")
        return
    }
    
    try {
        // 主逻辑
    } catch (e) {
        console.log(`错误: ${e.message}`)
    }
}

// 危险操作确认
function 删除确认() {
    const result = MsgBox("确定删除？", jsYesNo, "确认")
    if (result == 6) {
        // 执行删除
    }
}
```

## 参考资源

- **WPS 开放平台**: https://open.wps.cn/docs/office
- **WebOffice API**: https://solution.wps.cn/docs/client/api/summary.html
- **MDN JavaScript 参考**: https://developer.mozilla.org/zh-CN/docs/Web/JavaScript/Reference
