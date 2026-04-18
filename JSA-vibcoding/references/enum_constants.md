# JSA 枚举常量速查表

JSA 中常用的枚举常量参考。

## 对齐方式

### 水平对齐 (XlHAlign)

| 常量 | 值 | 说明 |
|------|-----|------|
| xlHAlignGeneral | 1 | 常规 |
| xlHAlignLeft | -4131 | 靠左 |
| xlHAlignCenter | -4108 | 居中 |
| xlHAlignRight | -4152 | 靠右 |
| xlHAlignFill | 5 | 填充 |
| xlHAlignJustify | -4130 | 两端对齐 |
| xlHAlignCenterAcrossSelection | 7 | 跨列居中 |
| xlHAlignDistributed | -4117 | 分散对齐 |

### 垂直对齐 (XlVAlign)

| 常量 | 值 | 说明 |
|------|-----|------|
| xlVAlignTop | -4160 | 顶端对齐 |
| xlVAlignCenter | -4108 | 垂直居中 |
| xlVAlignBottom | -4107 | 底端对齐 |
| xlVAlignJustify | -4130 | 两端对齐 |
| xlVAlignDistributed | -4117 | 分散对齐 |

**使用示例：**
```javascript
// 设置对齐方式
Range("A1:D10").HorizontalAlignment = xlHAlignCenter
Range("A1:D10").VerticalAlignment = xlVAlignCenter
```

## 边框

### 边框位置 (XlBordersIndex)

| 常量 | 值 | 说明 |
|------|-----|------|
| xlDiagonalDown | 5 | 斜向下 |
| xlDiagonalUp | 6 | 斜向上 |
| xlEdgeLeft | 7 | 左边框 |
| xlEdgeTop | 8 | 上边框 |
| xlEdgeBottom | 9 | 下边框 |
| xlEdgeRight | 10 | 右边框 |
| xlInsideHorizontal | 12 | 内部水平 |
| xlInsideVertical | 11 | 内部垂直 |

### 线条样式 (XlLineStyle)

| 常量 | 值 | 说明 |
|------|-----|------|
| xlContinuous | 1 | 实线 |
| xlDash | -4115 | 虚线 |
| xlDashDot | 4 | 点划线 |
| xlDashDotDot | 5 | 双点划线 |
| xlDot | -4118 | 点线 |
| xlDouble | -4119 | 双线 |
| xlSlantDashDot | 13 | 斜点划线 |
| xlLineStyleNone | -4142 | 无线条 |

### 边框粗细 (XlBorderWeight)

| 常量 | 值 | 说明 |
|------|-----|------|
| xlHairline | 1 | 极细 |
| xlThin | 2 | 细 |
| xlMedium | -4138 | 中等 |
| xlThick | 4 | 粗 |

**使用示例：**
```javascript
// 设置边框
const rng = Range("A1:D10")

// 外边框
for (const idx of [xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight]) {
    const border = rng.Borders.Item(idx)
    border.LineStyle = xlContinuous
    border.Weight = xlMedium
    border.Color = 0  // 黑色
}
```

## 数据有效性

### 数据类型 (XlDVType)

| 常量 | 值 | 说明 |
|------|-----|------|
| xlValidateInputOnly | 0 | 仅在用户更改值时验证 |
| xlValidateWholeNumber | 1 | 整数 |
| xlValidateDecimal | 2 | 小数 |
| xlValidateList | 3 | 列表 |
| xlValidateDate | 4 | 日期 |
| xlValidateTime | 5 | 时间 |
| xlValidateTextLength | 6 | 文本长度 |
| xlValidateCustom | 7 | 自定义公式 |

### 警告样式 (XlDVAlertStyle)

| 常量 | 值 | 说明 |
|------|-----|------|
| xlValidAlertStop | 1 | 停止 |
| xlValidAlertWarning | 2 | 警告 |
| xlValidAlertInformation | 3 | 信息 |

**使用示例：**
```javascript
// 添加下拉列表验证
const rng = Range("B2:B100")
rng.Validation.Delete()
rng.Validation.Add(xlValidateList, xlValidAlertStop, xlBetween, "男,女,其他", undefined)
```

## 文件格式

### 文件格式 (XlFileFormat)

| 常量 | 值 | 说明 |
|------|-----|------|
| xlOpenXMLWorkbook | 51 | xlsx |
| xlOpenXMLWorkbookMacroEnabled | 52 | xlsm |
| xlExcel8 | 56 | xls |
| xlCSV | 6 | CSV |
| xlPDF | 57 | PDF |
| xlOpenDocumentSpreadsheet | 60 | ODS |

## 查找

### 查找范围 (LookIn)

| 常量 | 值 | 说明 |
|------|-----|------|
| xlFormulas | -4123 | 公式 |
| xlValues | -4163 | 值 |
| xlComments | -4144 | 批注 |

### 查找方式 (LookAt)

| 常量 | 值 | 说明 |
|------|-----|------|
| xlWhole | 1 | 整个单元格匹配 |
| xlPart | 2 | 部分匹配 |

### 搜索顺序 (XlSearchOrder)

| 常量 | 值 | 说明 |
|------|-----|------|
| xlByRows | 1 | 按行 |
| xlByColumns | 2 | 按列 |

### 搜索方向 (XlSearchDirection)

| 常量 | 值 | 说明 |
|------|-----|------|
| xlNext | 1 | 向下搜索 |
| xlPrevious | 2 | 向上搜索 |

**使用示例：**
```javascript
// 查找示例
let rng = Cells.Find("搜索内容", undefined, xlValues, xlPart, xlByRows, xlNext)
```

## 引用样式

### 引用样式 (XlReferenceStyle)

| 常量 | 值 | 说明 |
|------|-----|------|
| xlA1 | 1 | A1 样式 |
| xlR1C1 | -4150 | R1C1 样式 |

**使用示例：**
```javascript
// 获取单元格地址
const addr1 = Range("A1").Address()                          // $A$1
const addr2 = Range("A1").Address(false, false)              // A1
const addr3 = Range("A1").Address(true, true, xlR1C1)        // R1C1
```

## 排序

### 排序顺序 (XlSortOrder)

| 常量 | 值 | 说明 |
|------|-----|------|
| xlAscending | 1 | 升序 |
| xlDescending | 2 | 降序 |

### 排序类型 (XlSortOn)

| 常量 | 值 | 说明 |
|------|-----|------|
| xlSortOnValues | 0 | 值 |
| xlSortOnCellColor | 1 | 单元格颜色 |
| xlSortOnFontColor | 2 | 字体颜色 |
| xlSortOnIcon | 3 | 图标 |

### 排序方向 (XlSortOrientation)

| 常量 | 值 | 说明 |
|------|-----|------|
| xlSortColumns | 1 | 按列排序 |
| xlSortRows | 2 | 按行排序 |

**使用示例：**
```javascript
// 排序
ActiveSheet.Sort.SortFields.Clear()
ActiveSheet.Sort.SortFields.Add(Range("A2"), xlSortOnValues, xlAscending)
ActiveSheet.Sort.SetRange(Range("A1:D100"))
ActiveSheet.Sort.Header = xlYes
ActiveSheet.Sort.Apply()
```

## 图表

### 图表类型 (XlChartType)

| 常量 | 值 | 说明 |
|------|-----|------|
| xlColumnClustered | 51 | 簇状柱形图 |
| xlColumnStacked | 52 | 堆积柱形图 |
| xlBarClustered | 57 | 簇状条形图 |
| xlBarStacked | 58 | 堆积条形图 |
| xlLine | 4 | 折线图 |
| xlLineMarkers | 65 | 带标记折线图 |
| xlPie | 5 | 饼图 |
| xlXYScatter | -4169 | 散点图 |
| xlArea | 1 | 面积图 |
| xlDoughnut | -4120 | 圆环图 |
| xlRadar | -4151 | 雷达图 |

**使用示例：**
```javascript
// 创建图表
const chart = ActiveSheet.Shapes.AddChart(xlColumnClustered)
chart.Chart.SetSourceData(Range("A1:B10"))
```

## 粘贴类型

### 粘贴类型 (XlPasteType)

| 常量 | 值 | 说明 |
|------|-----|------|
| xlPasteAll | -4104 | 全部 |
| xlPasteFormulas | -4123 | 公式 |
| xlPasteValues | -4163 | 值 |
| xlPasteFormats | -4122 | 格式 |
| xlPasteComments | -4144 | 批注 |
| xlPasteAllExceptBorders | 7 | 除边框外的全部内容 |

### 粘贴运算 (XlPasteSpecialOperation)

| 常量 | 值 | 说明 |
|------|-----|------|
| xlPasteSpecialOperationNone | -4142 | 无运算 |
| xlPasteSpecialOperationAdd | 2 | 加 |
| xlPasteSpecialOperationSubtract | 3 | 减 |
| xlPasteSpecialOperationMultiply | 4 | 乘 |
| xlPasteSpecialOperationDivide | 5 | 除 |

**使用示例：**
```javascript
// 粘贴特殊
Range("A1:D10").Copy()
Range("F1").PasteSpecial(xlPasteValues)      // 只粘贴值
Range("F1").PasteSpecial(xlPasteFormats)     // 只粘贴格式
```

## 自动填充

### 填充类型 (XlAutoFillType)

| 常量 | 值 | 说明 |
|------|-----|------|
| xlFillDefault | 0 | 默认填充 |
| xlFillCopy | 1 | 复制单元格 |
| xlFillSeries | 2 | 填充序列 |
| xlFillFormats | 3 | 仅填充格式 |
| xlFillValues | 4 | 仅填充值 |
| xlFillDays | 5 | 以天数填充 |
| xlFillWeekdays | 6 | 以工作日填充 |
| xlFillMonths | 7 | 以月填充 |
| xlFillYears | 8 | 以年填充 |

**使用示例：**
```javascript
// 自动填充
Range("A1:A2").AutoFill(Range("A1:A20"), xlFillSeries)
```

## 颜色值参考

### 常用颜色 RGB 值

| 颜色名 | RGB | 十进制值 |
|--------|-----|---------|
| 黑色 | (0, 0, 0) | 0 |
| 红色 | (255, 0, 0) | 255 |
| 绿色 | (0, 128, 0) | 32768 |
| 蓝色 | (0, 0, 255) | 16711680 |
| 黄色 | (255, 255, 0) | 65535 |
| 青色 | (0, 255, 255) | 16776960 |
| 紫色 | (255, 0, 255) | 16711935 |
| 白色 | (255, 255, 255) | 16777215 |
| 橙色 | (255, 165, 0) | 49407 |
| 粉色 | (255, 192, 203) | 13353215 |
| 灰色 | (128, 128, 128) | 8421504 |

### 颜色计算函数

```javascript
/**
 * RGB转十进制颜色值
 * @param {number} r - 红色 (0-255)
 * @param {number} g - 绿色 (0-255)
 * @param {number} b - 蓝色 (0-255)
 * @returns {number} 十进制颜色值
 */
function RGB(r, g, b) {
    return r + g * 256 + b * 256 * 256
}

// 使用示例
const red = RGB(255, 0, 0)       // 255
const blue = RGB(0, 0, 255)      // 16711680
const green = RGB(0, 128, 0)     // 32768
const yellow = RGB(255, 255, 0)  // 65535

// 设置颜色
Range("A1").Font.Color = RGB(255, 0, 0)       // 红色字体
Range("A1").Interior.Color = RGB(255, 255, 0) // 黄色背景
```

## 消息框常量

### MsgBox 返回值

| 常量 | 值 | 说明 |
|------|-----|------|
| jsOK | 1 | 确定 |
| jsCancel | 2 | 取消 |
| jsAbort | 3 | 终止 |
| jsRetry | 4 | 重试 |
| jsIgnore | 5 | 忽略 |
| jsYes | 6 | 是 |
| jsNo | 7 | 否 |

**使用示例：**
```javascript
// 是/否消息框
const result = MsgBox("确定删除吗?", jsYesNo, "确认")
if (result == 6) {  // 点击了"是"
    // 执行删除
}
```

## 数字格式字符串

| 格式 | 说明 | 示例 |
|------|------|------|
| `@` | 文本 | "ABC" |
| `0` | 数字占位符 | 123 |
| `0.00` | 两位小数 | 123.45 |
| `#,##0` | 千分位 | 1,234 |
| `#,##0.00` | 千分位+两位小数 | 1,234.56 |
| `0%` | 百分比（整数） | 50% |
| `0.00%` | 百分比（两位小数） | 50.00% |
| `￥#,##0.00` | 货币 | ￥1,234.56 |
| `yyyy/m/d` | 日期 | 2024/1/15 |
| `yyyy-mm-dd` | 日期 | 2024-01-15 |
| `hh:mm:ss` | 时间 | 14:30:00 |
| `yyyy/m/d hh:mm` | 日期时间 | 2024/1/15 14:30 |

**使用示例：**
```javascript
Range("A1").NumberFormat = "@"              // 文本
Range("A2").NumberFormat = "0.00"           // 两位小数
Range("A3").NumberFormat = "yyyy/m/d"       // 日期
Range("A4").NumberFormat = "￥#,##0.00"     // 货币
Range("A5").NumberFormat = "0.00%"          // 百分比
```
