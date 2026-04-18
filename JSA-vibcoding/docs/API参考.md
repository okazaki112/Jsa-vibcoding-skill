# JSA API 参考

JSA 核心 API 快速参考文档。

## Application 对象

顶层应用程序对象。

### 属性

| 属性 | 说明 |
|------|------|
| ActiveSheet | 活动工作表 |
| ActiveWorkbook | 活动工作簿 |
| Cells | 所有单元格 |
| Columns | 所有列 |
| Rows | 所有行 |
| Selection | 当前选择 |
| ThisWorkbook | 当前工作簿 |
| DisplayAlerts | 是否显示警告 |
| ScreenUpdating | 是否更新屏幕 |

### 方法

| 方法 | 说明 |
|------|------|
| InputBox() | 显示输入框 |
| Run() | 运行宏 |

## Range 对象

表示单元格区域。

### 属性

| 属性 | 说明 |
|------|------|
| Value2 | 值（读写） |
| Value | 值（写入） |
| Formula | 公式 |
| FormulaArray | 数组公式 |
| Text | 格式化文本（只读） |
| NumberFormat | 数字格式 |
| Row | 行号 |
| Column | 列号 |
| Count | 单元格数量 |
| Cells | 单元格集合 |
| Rows | 行集合 |
| Columns | 列集合 |
| Font | 字体对象 |
| Interior | 内部（背景）对象 |
| Borders | 边框对象 |
| MergeArea | 合并区域 |
| MergeCells | 是否合并 |

### 方法

| 方法 | 说明 |
|------|------|
| Select() | 选择 |
| Activate() | 激活 |
| Offset(row, col) | 偏移 |
| Resize(rows, cols) | 调整大小 |
| Merge() | 合并 |
| UnMerge() | 取消合并 |
| Clear() | 清除全部 |
| ClearContents() | 清除内容 |
| ClearFormats() | 清除格式 |
| Copy() | 复制 |
| PasteSpecial() | 粘贴特殊 |
| Delete() | 删除 |
| Insert() | 插入 |
| Find() | 查找 |
| FindAll() | 查找所有 |
| Replace() | 替换 |
| AutoFill() | 自动填充 |
| AutoFilter() | 自动筛选 |
| RemoveDuplicates() | 删除重复值 |
| Sort() | 排序 |
| Address() | 获取地址 |

## Sheets/Worksheets 集合

工作表集合。

### 属性

| 属性 | 说明 |
|------|------|
| Count | 工作表数量 |
| Item(index) | 获取工作表 |

### 方法

| 方法 | 说明 |
|------|------|
| Add() | 新建工作表 |
| Delete() | 删除工作表 |
| Copy() | 复制工作表 |
| Move() | 移动工作表 |

## Worksheet 对象

单个工作表。

### 属性

| 属性 | 说明 |
|------|------|
| Name | 工作表名称 |
| Index | 工作表索引 |
| Cells | 所有单元格 |
| Columns | 所有列 |
| Rows | 所有行 |
| Range | 区域对象 |
| AutoFilterMode | 是否开启筛选 |
| UsedRange | 已使用区域 |
| Visible | 是否可见 |

### 方法

| 方法 | 说明 |
|------|------|
| Select() | 选择 |
| Activate() | 激活 |
| Copy() | 复制 |
| Move() | 移动 |
| Delete() | 删除 |
| Calculate() | 计算 |

## Workbook 对象

工作簿对象。

### 属性

| 属性 | 说明 |
|------|------|
| Name | 工作簿名称 |
| Path | 文件路径 |
| FullName | 完整路径 |
| Sheets | 工作表集合 |
| Worksheets | 工作表集合 |
| ReadOnly | 是否只读 |
| Saved | 是否已保存 |

### 方法

| 方法 | 说明 |
|------|------|
| Save() | 保存 |
| SaveAs() | 另存为 |
| Close() | 关闭 |

## Workbooks 集合

工作簿集合。

### 方法

| 方法 | 说明 |
|------|------|
| Open() | 打开工作簿 |
| Add() | 新建工作簿 |
| Close() | 关闭所有工作簿 |

## Font 对象

字体对象。

### 属性

| 属性 | 说明 |
|------|------|
| Name | 字体名称 |
| Size | 字体大小 |
| Bold | 是否加粗 |
| Italic | 是否斜体 |
| Underline | 下划线样式 |
| Color | 字体颜色 |
| Strikethrough | 是否删除线 |

## Interior 对象

背景填充对象。

### 属性

| 属性 | 说明 |
|------|------|
| Color | 背景色 |
| Pattern | 填充图案 |
| PatternColor | 图案颜色 |

## Borders 对象

边框对象。

### 属性

| 属性 | 说明 |
|------|------|
| LineStyle | 线条样式 |
| Weight | 线条粗细 |
| Color | 边框颜色 |
| Item(index) | 特定边框 |

## Validation 对象

数据有效性对象。

### 方法

| 方法 | 说明 |
|------|------|
| Add() | 添加验证 |
| Delete() | 删除验证 |

### 属性

| 属性 | 说明 |
|------|------|
| InputTitle | 输入标题 |
| InputMessage | 输入提示 |
| ErrorTitle | 错误标题 |
| ErrorMessage | 错误提示 |

## FormatConditions 集合

条件格式集合。

### 方法

| 方法 | 说明 |
|------|------|
| Add() | 添加条件格式 |
| Delete() | 删除条件格式 |
| Item(index) | 获取条件格式 |

## Sort 对象

排序对象。

### 属性

| 属性 | 说明 |
|------|------|
| SortFields | 排序字段集合 |
| Header | 是否有标题 |
| Orientation | 排序方向 |

### 方法

| 方法 | 说明 |
|------|------|
| SetRange() | 设置范围 |
| Apply() | 应用排序 |

## 内置函数

### MsgBox

显示消息框。

```javascript
// 简单消息
MsgBox("消息内容")

// 带标题
MsgBox("消息内容", "标题")

// 是/否选择
let result = MsgBox("确定吗?", jsYesNo, "确认")
if (result == 6) {
    // 点击了"是"
}
```

### InputBox

显示输入框。

```javascript
// 简单输入
let value = InputBox("请输入")

// 完整参数
let value = InputBox("提示", "标题", "默认值", x, y)

// 选择区域
let rng = Application.InputBox("选择区域", "标题", undefined, undefined, undefined, undefined, undefined, 8)
```

### alert

显示警告框。

```javascript
alert("提示信息")
```

### console.log

输出到立即窗口。

```javascript
console.log("调试信息")
console.log(`变量值: ${value}`)
```

## WebOffice SDK API

### 初始化

```javascript
const instance = WebOfficeSDK.init({
    officeType: WebOfficeSDK.OfficeType.Spreadsheet,
    appId: 'your-app-id',
    fileId: 'your-file-id',
    mount: document.querySelector('#app')
})
```

### OfficeType 枚举

| 值 | 说明 |
|------|------|
| Spreadsheet | 表格 |
| Writer | 文字 |
| Presentation | 演示 |
| Pdf | PDF |

### 异步调用

```javascript
// 等待就绪
await instance.ready()

// 获取 Application
const app = instance.Application

// 操作（异步）
const range = await app.Range('A1')
range.Value = 'Hello'

// 枚举
range.HorizontalAlignment = await app.Enum.XlHAlign.xlHAlignCenter
```

### 事件监听

```javascript
instance.on('fileOpen', (data) => {
    console.log('文件打开成功', data)
})

instance.on('error', (error) => {
    console.error('错误', error)
})
```

## 常用代码片段

### 读取/写入值

```javascript
// 读取
let value = Range("A1").Value2

// 写入
Range("A1").Value2 = "数据"

// 批量写入
Range("A1:D10").Value2 = dataArray
```

### 获取边界

```javascript
// 最后一行
function GetLastRow(ws, col) {
    let lastCell = ws.Columns(col).Find("*", undefined, undefined, undefined, xlByRows, xlPrevious)
    return lastCell ? lastCell.Row : 0
}

// 最后一列
function GetLastCol(ws, row) {
    let lastCell = ws.Rows(row).Find("*", undefined, undefined, undefined, xlByColumns, xlPrevious)
    return lastCell ? lastCell.Column : 0
}
```

### 颜色计算

```javascript
function RGB(r, g, b) {
    return r + g * 256 + b * 256 * 256
}

// 使用
Range("A1").Font.Color = RGB(255, 0, 0)  // 红色
```
