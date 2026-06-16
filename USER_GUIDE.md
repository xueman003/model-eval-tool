# 使用说明：下载并使用工具

这是一份面向非技术用户的说明文档。正常情况下，只需要下载、解压、双击启动文件即可使用。

## 1. 下载项目

在 GitHub 仓库页面点击：

```text
Code -> Download ZIP
```

下载后解压到本地任意目录，例如桌面或文档目录。

解压后，项目目录中应包含这些文件：

```text
app.py
index.html
README.md
requirements.txt
start.sh
start.bat
examples/
```

## 2. 启动工具

### macOS

双击：

```text
start.command
```

如果系统提示无法打开，可以右键点击 `start.command`，选择“打开”。

### Windows

双击：

```text
start.bat
```

### Linux

双击或运行：

```bash
start.sh
```

首次运行时，工具会自动检查依赖。如果缺少依赖，会自动安装。安装完成后会打开浏览器页面。

## 3. 打开页面

启动成功后，会自动打开浏览器。如果没有自动打开，请手动访问：

```text
http://127.0.0.1:8765
```

如果终端窗口里显示的是其他端口，例如：

```text
http://127.0.0.1:8766
```

请打开终端窗口里显示的实际地址。

## 4. 使用流程

1. 点击“评估结果 Excel”，选择要分析的 `.xlsx` 文件。
2. 点击“读取表格”。
3. 如果 Excel 有多个工作表，选择要处理的工作表。
4. 在“原始 JSON 列”中选择模型输出所在列。
5. 点击“拆分 JSON 并刷新列”。
6. 在“最终结果列（模型分）”中选择用于计算阈值的分数字段，例如 `prob_1`。
7. 在“正确结果列”中选择人工结果或真实标签列。
8. 在“正例取值”中填写目标类别，例如：

```text
target_a
```

如果有多个正例，可以用逗号分隔：

```text
target_a,target_b
```

9. 选择分数方向：

```text
分越高越像正例
```

或：

```text
分越低越像正例
```

10. 点击“生成统计与 Excel”。
11. 查看“分层分布”和“阈值准召”，或点击“下载处理结果”导出 Excel。

## 5. 输出文件说明

下载的 Excel 包含：

- `拆分明细`：原始数据 + JSON 拆分字段 + 解析状态
- `统计概览`：样本数、正例数、选中的模型分列和正确结果列
- `分层分布`：按 `0.01` 模型分区间统计样本数量、正例数量、正例率、召回占比，并按分数区间倒序展示
- `阈值准召`：按 `0.01` 阈值统计 precision、recall、F1、TP、FP、FN，并按阈值倒序展示

## 6. 示例文件

项目中的 `examples/` 目录提供了脱敏示例：

```text
examples/sample_basic.xlsx
examples/sample_multi_sheet.xlsx
```

可以先用示例文件体验完整流程。

## 7. 常见问题

### 双击后提示未检测到 Python

本工具需要本机安装 Python。

推荐版本：

```text
Python 3.9 或更高版本
```

安装 Python 后，再重新双击启动文件。

### 双击后提示依赖安装失败

通常是网络不可用，或 Python/pip 没有正确安装。可以在网络恢复后重新双击启动文件。

也可以在项目目录中手动运行：

```bash
python3 -m pip install -r requirements.txt
```

Windows 用户可尝试：

```bat
python -m pip install -r requirements.txt
```

如果双击后没有反应，可以在终端中进入项目目录，执行一次：

```bash
chmod +x start.command
```

然后再双击 `start.command`。

### 打开 index.html 后不能上传处理

不要直接双击打开 `index.html` 使用完整功能。

完整功能需要先启动本地服务：

```bash
python3 app.py
```

然后访问：

```text
http://127.0.0.1:8765
```

### 双击 start.command 后仍然无法访问 127.0.0.1

请确认双击后出现的终端窗口没有报错，并且窗口保持打开。

如果窗口提示本地服务启动失败，可以依次检查：

1. 是否已经安装依赖：

```bash
pip install -r requirements.txt
```

2. 是否有其他程序占用了 `8765` 端口。可以关闭其他正在运行的本工具窗口后重试。

3. macOS 是否拦截了终端或脚本权限。可以右键点击 `start.command`，选择“打开”，再确认运行。

正常情况下，终端窗口保持打开时，浏览器才能访问：

```text
http://127.0.0.1:8765
```

如果看到：

```text
OSError: [Errno 48] Address already in use
```

说明 `8765` 端口已经被占用。通常是本工具已经在另一个终端窗口中运行。新版 `start.command` 会自动尝试使用其他可用端口，例如：

```text
http://127.0.0.1:8766
```

如果浏览器没有自动打开，请看终端窗口里显示的实际地址。

如果确认已有工具在运行，也可以直接打开：

```text
http://127.0.0.1:8765
```

或者关闭之前启动工具的终端窗口后，再重新双击 `start.command`。

### precision、recall、F1 都是 0

通常是“正例取值”没有匹配到正确结果列里的真实取值。

例如正确结果列里是：

```text
target_a
target_b
target_c
```

如果目标是召回 `target_a` 类样本，正例取值应填写：

```text
target_a
```

### 找不到 JSON 拆分字段

请确认“原始 JSON 列”选的是模型输出列，并且单元格内容是类似：

```json
{"prob_1": 0.82, "prob_2": 0.18, "reason": "xxx"}
```

空输出、运行失败、非法 JSON 会被保留，但不会拆出字段。

### 文件包含多个 sheet

上传后在“工作表”下拉框中选择要处理的 sheet。

### 数据会不会上传到外部

不会。工具运行在本机：

```text
http://127.0.0.1:8765
```

Excel 文件只在本地处理，不会发送到外部 API。

## 8. 停止工具

回到启动工具的终端窗口，按：

```text
Ctrl + C
```

即可停止本地服务。
