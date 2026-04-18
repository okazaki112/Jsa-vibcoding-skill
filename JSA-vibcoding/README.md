# JSA-VibCoding

**WPS JSA 智能编程助手 - AI Agent Skill**

基于 JavaScript 的 WPS 宏开发解决方案。通过自然语言描述需求，自动生成高质量 JSA 代码，支持单元格操作、工作表管理、数据处理、事件处理等完整开发流程。

[![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![WPS](https://img.shields.io/badge/WPS-2021+-green.svg)](https://www.wps.cn)

---

## ✨ 特性

- 🚀 **代码生成** - 自然语言描述需求，自动生成高质量 JSA 代码
- 📚 **完整知识库** - 涵盖 Range、Sheet、Workbook 等 API 详细说明
- ⚡ **现代语法** - 支持 ES6-ES2019 语法（箭头函数、类、解构、模板字符串等）
- 🔧 **丰富模板** - 提供数据处理、表格操作、事件处理等标准代码模板
- 🎯 **最佳实践** - 性能优化、命名规范、调试技巧一应俱全

## 📖 什么是 JSA？

JSA (JavaScript for Application) 是 WPS 从 2021 版本开始支持的宏语言，基于 JavaScript：

| 特性 | 说明 |
|------|------|
| 运行时 | 内嵌 V8 引擎 |
| 语法支持 | ES6 - ES2019 |
| API 兼容 | 与 VBA 高度兼容 |
| 索引规则 | 从 1 开始（与 VBA 一致） |

### JSA vs VBA

| 特性 | VBA | JSA |
|------|-----|-----|
| 语言 | Visual Basic | JavaScript |
| 变量 | `Dim x As Long` | `let x = 0` |
| 循环 | `For i = 1 To 10` | `for (let i = 1; i <= 10; i++)` |
| 数组 | `arr(0)` | `arr[0]` |
| 现代特性 | 较少 | 丰富（箭头函数、类、解构等） |

## 🚀 快速开始

### 1. 安装

将本项目复制到你的 AI Agent skill 目录



### 2. 启用 JS 宏环境

1. 打开 WPS 表格
2. 点击【文件】→【选项】
3. 勾选【默认 JS 开发环境】
4. 按 `Alt + F11` 打开宏编辑器

### 3. 基础示例

```javascript
function 数据处理() {
    const ws = ActiveSheet
    const lastRow = GetLastRow(ws, 1)
    
    // 读取到数组（性能优化关键）
    const data = ws.Range("A2:D" + lastRow).Value2
    
    // 内存中处理
    for (let i = 0; i < data.length; i++) {
        if (data[i][1] && data[i][2]) {
            data[i][3] = data[i][1] * data[i][2]  // 金额 = 数量 × 单价
        }
    }
    
    // 一次性写回
    ws.Range("A2:D" + lastRow).Value2 = data
}

function GetLastRow(ws, col) {
    const lastCell = ws.Columns(col).Find("*", undefined, undefined, undefined, xlByRows, xlPrevious)
    return lastCell ? lastCell.Row : 0
}
```

## 📁 项目结构

```
JSA-vibcoding/
├── SKILL.md                      # Skill 主定义文件
├── docs/
│   ├── 使用指南.md                # 详细使用教程
│   └── API参考.md                 # API 文档
├── references/
│   ├── code_templates.md         # 代码模板库
│   ├── jsa_patterns.md           # JSA 常用模式
│   ├── best_practices.md         # 最佳实践
│   └── enum_constants.md         # 枚举常量速查
└── examples/
    └── basic_usage.js            # 示例代码
```

## 🎯 使用方式

### 触发 Skill

在与 AI 对话时使用以下关键词：

- "帮我写一个 JSA 代码..."
- "用 WPS 宏实现..."
- "JSA 数据处理..."
- "优化 JSA 性能..."

### 示例对话

```
用户：帮我用 JSA 实现一个功能，从 A 列读取产品名称，B 列读取数量，
     C 列读取单价，计算金额写入 D 列。

AI：好的，我来为你生成 JSA 代码...
    [自动调用 JSA-vibcoding skill 生成代码]
```

## 📚 知识库内容

| 文档 | 内容 |
|------|------|
| **API 参考** | Range、Sheet、Workbook 等 API 详细说明 |
| **代码模板** | 标准函数结构、数组操作、格式设置等模板 |
| **常用模式** | 单元格操作、数据处理、事件处理等模式 |
| **最佳实践** | 命名规范、性能优化、调试技巧 |
| **枚举常量** | 对齐、边框、颜色、格式等常量速查 |

## ⚡ 性能优化

### ✅ 推荐：数组批量操作

```javascript
// 快：数组批量处理（10000 行 < 1 秒）
const data = ws.Range("A1:D10000").Value2
for (let i = 0; i < data.length; i++) {
    data[i][3] = data[i][1] * data[i][2]
}
ws.Range("A1:D10000").Value2 = data
```

### ❌ 避免：逐单元格操作

```javascript
// 慢：逐单元格操作（10000 行 > 30 秒）
for (let i = 1; i <= 10000; i++) {
    ws.Cells(i, 4).Value2 = ws.Cells(i, 2).Value2 * ws.Cells(i, 3).Value2
}
```

## 🔧 事件处理

JSA 通过特定函数名实现事件处理：

```javascript
// 单元格值改变
function Workbook_SheetChange(Sh, rg) {
    console.log(`工作表 ${Sh.Name} 的 ${rg.Address()} 发生变化`)
}

// 工作表激活
function Workbook_SheetActivate(Sh) {
    console.log(`激活了工作表: ${Sh.Name}`)
}
```

## 🔗 参考资源

- [WPS 开放平台](https://open.wps.cn/docs/office)
- [WebOffice API](https://solution.wps.cn/docs/client/api/summary.html)
- [MDN JavaScript 参考](https://developer.mozilla.org/zh-CN/docs/Web/JavaScript/Reference)
- [ES6 入门教程](https://es6.ruanyifeng.com/)

## 🤝 贡献

欢迎提交 Issue 和 Pull Request！

## 📄 许可证

[MIT License](LICENSE)
