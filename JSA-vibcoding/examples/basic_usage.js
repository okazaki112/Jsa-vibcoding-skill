/**
 * JSA 基础使用示例
 * 演示 JSA 开发的常用场景和代码模式
 */

// ============================================
// 一、单元格操作
// ============================================

/**
 * 基础单元格读写
 */
function 基础单元格操作() {
    // 读取单元格
    let value = Range("A1").Value2
    console.log("A1的值: " + value)
    
    // 写入单元格
    Range("A1").Value2 = "Hello JSA"
    
    // 批量写入（一维数组）
    Range("A1:C1").Value2 = ["姓名", "年龄", "城市"]
    
    // 批量写入（二维数组）
    let data = [
        ["张三", 25, "北京"],
        ["李四", 30, "上海"],
        ["王五", 28, "广州"]
    ]
    Range("A2:C4").Value2 = data
    
    // 使用公式
    Range("D2").Formula = "=B2*C2"
    Range("D3:D4").Formula = "=B3*C3"
}

/**
 * 单元格格式设置
 */
function 单元格格式设置() {
    let rng = Range("A1:D10")
    
    // 字体设置
    rng.Font.Name = "微软雅黑"
    rng.Font.Size = 12
    rng.Font.Bold = true
    rng.Font.Color = 255  // 红色
    
    // 背景色
    rng.Interior.Color = 65535  // 黄色
    
    // 对齐
    rng.HorizontalAlignment = xlHAlignCenter
    rng.VerticalAlignment = xlVAlignCenter
    
    // 边框
    rng.Borders.LineStyle = xlContinuous
    rng.Borders.Weight = xlThin
    
    // 数字格式
    rng.NumberFormat = "0.00"
}

/**
 * 合并单元格
 */
function 合并单元格示例() {
    // 合并
    Range("A1:D1").Merge()
    Range("A1:D1").HorizontalAlignment = xlHAlignCenter
    
    // 取消合并
    // Range("A1:D1").UnMerge()
}

// ============================================
// 二、工作表操作
// ============================================

/**
 * 工作表管理
 */
function 工作表操作() {
    // 选择工作表
    Sheets.Item(1).Select()           // 按索引
    Sheets.Item("Sheet2").Select()    // 按名称
    
    // 新建工作表
    let ws = Worksheets.Add()
    ws.Name = "新工作表"
    
    // 在最后位置新建
    let wsLast = Worksheets.Add(null, Sheets.Item(Sheets.Count))
    wsLast.Name = "最后一个工作表"
    
    // 删除工作表
    Application.DisplayAlerts = false
    // Sheets.Item("Sheet2").Delete()
    Application.DisplayAlerts = true
    
    // 重命名
    Sheets.Item(1).Name = "数据表"
    
    // 工作表数量
    console.log("工作表数量: " + Sheets.Count)
}

/**
 * 遍历所有工作表
 */
function 遍历工作表() {
    for (let i = 1; i <= Sheets.Count; i++) {
        console.log(`工作表 ${i}: ${Sheets.Item(i).Name}`)
    }
}

// ============================================
// 三、数据处理
// ============================================

/**
 * 获取最后一行
 */
function GetLastRow(ws, col) {
    let lastCell = ws.Columns(col).Find("*", undefined, undefined, undefined, xlByRows, xlPrevious)
    return lastCell ? lastCell.Row : 0
}

/**
 * 数组批量处理（推荐方式）
 */
function 数组批量处理() {
    let ws = ActiveSheet
    let lastRow = GetLastRow(ws, 1)
    
    if (lastRow < 2) {
        alert("数据不足")
        return
    }
    
    // 读取到数组
    let data = ws.Range("A2:D" + lastRow).Value2
    
    // 内存处理
    for (let i = 0; i < data.length; i++) {
        // 金额 = 数量 * 单价
        if (data[i][1] && data[i][2]) {
            data[i][3] = data[i][1] * data[i][2]
        }
    }
    
    // 一次性写回
    ws.Range("A2:D" + lastRow).Value2 = data
    
    alert("处理完成！")
}

/**
 * 查找数据
 */
function 查找数据() {
    let rng = Cells.Find("搜索内容")
    
    if (rng) {
        rng.Select()
        alert("找到了: " + rng.Address())
    } else {
        alert("未找到")
    }
}

/**
 * 查找所有匹配项
 */
function 查找所有() {
    let results = Cells.FindAll("搜索内容")
    
    if (results && results.length > 0) {
        console.log(`找到 ${results.length} 个匹配项`)
        for (let addr of results) {
            console.log(addr)
        }
    }
}

/**
 * 替换数据
 */
function 替换数据() {
    let count = Cells.Replace("旧内容", "新内容")
    console.log(`替换了 ${count} 处`)
}

/**
 * 删除重复值
 */
function 删除重复值() {
    Range("A1:D100").RemoveDuplicates([1, 2])  // 按第1、2列去重
}

// ============================================
// 四、用户交互
// ============================================

/**
 * 输入框
 */
function 输入框示例() {
    // 简单输入
    let name = InputBox("请输入姓名")
    if (name) {
        alert("你输入了: " + name)
    }
    
    // 完整参数
    let value = InputBox("请输入数值", "输入框", "100", 200, 200)
    console.log("输入的值: " + value)
}

/**
 * 消息框
 */
function 消息框示例() {
    // 简单消息
    MsgBox("操作完成!")
    
    // 带标题
    MsgBox("操作完成!", "提示")
    
    // 是/否选择
    let result = MsgBox("确定删除吗?", jsYesNo, "确认")
    if (result == 6) {  // 6 = 是
        alert("点击了是")
    } else {
        alert("点击了否")
    }
}

/**
 * 选择单元格区域
 */
function 选择区域() {
    let rng = Application.InputBox(
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
        alert("你选择了: " + rng.Address())
    }
}

// ============================================
// 五、事件处理
// ============================================

/**
 * 单元格值改变事件
 */
function Workbook_SheetChange(Sh, rg) {
    console.log(`工作表 ${Sh.Name} 的 ${rg.Address()} 发生变化`)
}

/**
 * 工作表激活事件
 */
function Workbook_SheetActivate(Sh) {
    console.log(`激活了工作表: ${Sh.Name}`)
}

/**
 * 工作表选择改变事件
 */
function Workbook_SheetSelectionChange(Sh, rg) {
    // 高亮当前行（示例）
    // Cells.Interior.ColorIndex = xlNone
    // rg.EntireRow.Interior.Color = RGB(240, 240, 240)
}

// ============================================
// 六、自定义公式函数
// ============================================

/**
 * 删除指定字符及后面的内容
 * 使用方式: =DeleteFrom(A1, "@")
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

/**
 * 将区域转为文本表格
 * 使用方式: =ToTextTable(A1:D10, ",")
 */
function ToTextTable(target, sep) {
    if (target.constructor.name != 'Range') {
        throw new TypeError('target 参数必须是单元格区域')
    }
    
    sep = sep ? sep.toString() : ','
    let values = []
    
    for (let iRow = 1; iRow <= target.Rows.Count; iRow++) {
        let rowValues = []
        for (let iCol = 1; iCol <= target.Columns.Count; iCol++) {
            rowValues.push(target.Cells.Item(iRow, iCol).Value())
        }
        values.push(rowValues.join(sep))
    }
    
    return values.join('\n')
}

// ============================================
// 七、类封装示例
// ============================================

/**
 * 数据表处理类
 */
class DataTableProcessor {
    constructor(ws) {
        this.ws = ws || ActiveSheet
        this.data = null
        this.lastRow = 0
        this.lastCol = 0
    }
    
    init() {
        this.lastRow = this._getLastRow()
        this.lastCol = this._getLastCol()
        console.log(`数据范围: ${this.lastRow}行 x ${this.lastCol}列`)
        return this
    }
    
    read() {
        if (this.lastRow < 1) return this
        this.data = this.ws.Range("A1").Resize(this.lastRow, this.lastCol).Value2
        return this
    }
    
    write() {
        if (!this.data) return this
        this.ws.Range("A1").Resize(this.data.length, this.data[0].length).Value2 = this.data
        return this
    }
    
    process(fn) {
        if (!this.data) return this
        for (let i = 0; i < this.data.length; i++) {
            this.data[i] = fn(this.data[i], i)
        }
        return this
    }
    
    formatHeader() {
        if (this.lastRow < 1) return this
        let header = this.ws.Range("A1").Resize(1, this.lastCol)
        header.Font.Bold = true
        header.Interior.Color = 12611584  // 蓝色
        header.Font.Color = 16777215      // 白色
        return this
    }
    
    _getLastRow() {
        let lastCell = this.ws.Columns(1).Find("*", undefined, undefined, undefined, xlByRows, xlPrevious)
        return lastCell ? lastCell.Row : 0
    }
    
    _getLastCol() {
        let lastCell = this.ws.Rows(1).Find("*", undefined, undefined, undefined, xlByColumns, xlPrevious)
        return lastCell ? lastCell.Column : 0
    }
}

/**
 * 使用类处理数据
 */
function 使用类处理数据() {
    new DataTableProcessor()
        .init()
        .read()
        .process((row, idx) => {
            // 跳过标题行，计算金额
            if (idx > 0 && row[1] && row[2]) {
                row[3] = row[1] * row[2]
            }
            return row
        })
        .write()
        .formatHeader()
    
    alert("处理完成！")
}

// ============================================
// 八、完整示例：数据处理流程
// ============================================

/**
 * 完整数据处理示例
 */
function 完整数据处理示例() {
    const startTime = Date.now()
    const ws = ActiveSheet
    
    try {
        console.log("=".repeat(50))
        console.log("开始数据处理")
        console.log("=".repeat(50))
        
        // 1. 读取数据
        console.log("\n[步骤1] 读取数据")
        const lastRow = GetLastRow(ws, 1)
        if (lastRow < 2) {
            alert("数据不足")
            return
        }
        const data = ws.Range("A2:D" + lastRow).Value2
        console.log(`  读取到 ${data.length} 行数据`)
        
        // 2. 处理数据
        console.log("\n[步骤2] 处理数据")
        for (let i = 0; i < data.length; i++) {
            // 金额 = 数量 * 单价
            if (data[i][1] && data[i][2]) {
                data[i][3] = data[i][1] * data[i][2]
            }
        }
        console.log("  数据处理完成")
        
        // 3. 写入结果
        console.log("\n[步骤3] 写入结果")
        ws.Range("A2:D" + lastRow).Value2 = data
        
        // 4. 格式化
        console.log("\n[步骤4] 格式化")
        ws.Range("D2:D" + lastRow).NumberFormat = "#,##0.00"
        
        // 5. 完成
        const elapsed = ((Date.now() - startTime) / 1000).toFixed(2)
        console.log("\n" + "=".repeat(50))
        console.log(`处理完成，总耗时: ${elapsed}秒`)
        alert(`处理完成！用时: ${elapsed}秒`)
        
    } catch (e) {
        console.log(`\n错误: ${e.message}`)
        alert(`处理失败: ${e.message}`)
    }
}

// 按F5运行此函数测试
function RunExample() {
    完整数据处理示例()
}
