# 多平台订单导入分析 V1

这是一个本地运行的 Python 工具，用来把不同平台导出的订单表按不同模板解析后统一汇总。

当前已支持的模板：

- 拼多多-发货单导出
- 拼多多-打印单
- 拼多多-订单导出

## 当前能力

- 自动识别模板
- 导入 `.xlsx` / `.csv`
- 自动拆分一单多商品
- 区分下单时间、付款时间、打印时间、发货时间
- 区分发货状态、退款/退货/异常状态
- 自动提取商品种类、款式键、变体、尺码
- 按平台、店铺、商品种类、模板、时间范围、退款状态、发货状态筛选
- 导出明细 CSV 和汇总 CSV

## 技术栈

- Python 3.12
- 标准库
  - `sqlite3`
  - `http.server`
  - `zipfile`
  - `xml.etree.ElementTree`

当前版本**不依赖第三方包**。

## 目录

- `order_analysis_v1.py`：主程序
- `run_order_analysis_v1.ps1`：启动脚本
- `stop_order_analysis_v1.ps1`：停止脚本
- `start_order_analysis_v1.cmd`：双击启动
- `stop_order_analysis_v1.cmd`：双击停止

## 启动

### 方式 1：双击

双击：

```bat
start_order_analysis_v1.cmd
```

### 方式 2：PowerShell

```powershell
powershell -ExecutionPolicy Bypass -File .\run_order_analysis_v1.ps1
```

### 方式 3：后台启动

```powershell
powershell -ExecutionPolicy Bypass -File .\run_order_analysis_v1.ps1 -Background
```

启动后打开：

```text
http://127.0.0.1:8791
```

## 停止

```powershell
powershell -ExecutionPolicy Bypass -File .\stop_order_analysis_v1.ps1
```

## 默认数据目录

程序默认会在当前仓库同级目录下创建：

- `work/order_analysis_v1/`

里面包含：

- SQLite 数据库
- 导入的原始文件副本
- PID 文件

## 后续扩展建议

如果你后面要支持淘宝、抖店、京东，推荐继续沿用当前结构：

1. 为每个平台增加一个模板识别器
2. 为每个平台增加一个行解析器
3. 最终统一写入同一张标准化明细表

## GitHub 推送建议

如果你已经在本机安装并登录了 GitHub CLI：

```powershell
gh repo create order-analysis-v1 --private --source . --remote origin --push
```

如果没有 `gh`，也可以：

```powershell
git remote add origin <你的仓库地址>
git push -u origin main
```
