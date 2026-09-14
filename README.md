# MX 销售数据自动化 / MX Sales Data

用紫鸟（Ziniao）指纹浏览器自动登录墨西哥 MercadoLibre 与 MercadoPago 店铺，
下载全部报表，清洗成结构化数据，并为每家店生成一份对账工作簿。

全程无人值守：**一条命令，从下载到出报表**。

---

## 1. 它做什么

```
紫鸟浏览器 ──► 下载原始报表 ──► 清洗成 Parquet ──► 每店一份对账 Excel
   12 家店         15 种表单          DuckDB 视图        data/reports/
```

| 阶段 | 内容 |
|---|---|
| **下载** | 4 条路线：Facturación（账单）、Ventas（销售）、Reportes de stock（库存）、MercadoPago（6 种报表） |
| **清洗** | `downloads/mx_sales`（独立子项目）把 xlsx/csv 洗成 Parquet，并建 DuckDB 视图 |
| **对账** | 每家店一份 `.xlsx`，含月度汇总、差异、包裹拆分、明细 |

一家店一次完整下载约 **15–18 分钟**，12 家店整轮约 **3 小时 45 分**。

---

## 2. 首次安装

### 2.1 必须自己装的两样东西

| 依赖 | 版本 | 说明 |
|---|---|---|
| **紫鸟客户端** | 6.27.1 (V6) | 从紫鸟控制台下载安装 |
| **Python** | 3.11 | 3.8+ 均可，3.11 是验证过的版本 |

> ⚠️ 紫鸟账号必须是**企业登录**（公司名 + 用户名 + 密码），并且在紫鸟控制台里
> **开启 WebDriver 权限**。个人账号无法使用本项目。

### 2.2 装环境

```bat
cd D:\Projects\mx_sales_data

python -m venv .venv
.venv\Scripts\pip install -r requirements.txt

copy config.example.json config.json
```

### 2.3 填 config.json

```json
{
  "client_path": null,
  "client_version": "v6",
  "socket_port": 16851,
  "user_info": {
    "company": "公司名",
    "username": "用户名",
    "password": "密码"
  },
  "stores": {
    "default": "BOCINA_TA02",
    "allowed": ["BOCINA_TA02", "UNIT_TA04"],
    "batch": []
  }
}
```

- `client_path`：填 `null` 让程序自动找紫鸟安装位置。
- `allowed`：**白名单**——脚本只允许打开这里列出的店铺。
- `batch`：`run_batch.py` 依次跑哪些店；**留空 = 跑 allowed 里的全部**，一般留空即可。

> 🔒 白名单是**失败即拒绝**的：`config.json` 格式错误时脚本直接报错退出（退出码 4），
> 而不会退化成"不限制"。账号下能看到 93 家店，误开一家不是小事。

### 2.4 自检

```bat
.venv\Scripts\python bootstrap.py
```

看到 `environment OK` 再往下走。它会逐项告诉你缺什么、怎么补。

---

## 3. 日常使用

### 最常用的一条命令

```bat
cd D:\Projects\mx_sales_data
.venv\Scripts\python run_batch.py --workbooks
```

下载全部店铺 → 清洗 → 每店一份对账工作簿。中途不需要任何人工操作。

### 全部命令

| 命令 | 下载 | 清洗 | 对账簿 |
|---|---|---|---|
| `run_batch.py --workbooks` | 全部店铺 | ✅ | ✅ 每店一份 |
| `run_batch.py` | 全部店铺 | ✅ | ❌ |
| `run_batch.py --collect-only` | 只补没取回的报表 | ✅ | ❌ |
| `run_batch.py --stores 店名 --workbooks` | 指定一家 | ✅ | ✅ |
| `run_batch.py --dry-run` | 不开浏览器，只列出会跑哪些店 | ❌ | ❌ |
| `run_downloads.py --store 店名` | 指定一家 | ❌ | ❌ |
| `cd downloads` + `python -m mx_sales build` | 不下载 | ✅ | ❌ |

### 新增一家店

1. 把店名加进 `config.json` 的 `stores.allowed`
2. 跑：

```bat
.venv\Scripts\python run_batch.py --stores 新店名 --workbooks
```

### 补取上次没下完的报表

MercadoPago 有些报表要生成几十分钟甚至几小时。跑完一轮后，没取到的会记在
`logs/pending.json` 里。补取时**不用重新下载**：

```bat
.venv\Scripts\python run_batch.py --collect-only
```

它只会打开确实欠着报表的那几家店。

---

## 4. 输出在哪

```
downloads/<店名>/<时间戳>/          原始报表 + REPORT.md（本次下载说明）
downloads/data/processed/           清洗后的 Parquet
downloads/data/mx_sales.duckdb      可直接 SQL 查询的视图
downloads/data/reports/<店名>_<时间戳>.xlsx    对账工作簿
logs/batch_<时间戳>.json            每店状态（供定时任务读取）
logs/pending.json                   还欠着的报表 → 用 --collect-only 补
```

查数据：

```bat
cd downloads
..\.venv\Scripts\python -m mx_sales query "SELECT store, count(*) FROM facturacion GROUP BY store"
```

---

## 5. 新人最容易踩的坑

### ① 目录不对

- `run_batch.py` / `run_downloads.py` → 在**项目根目录**跑
- `python -m mx_sales ...` → 在 **`downloads/` 目录**跑

### ② Python 用错了

机器上同时有 conda 的 `(base)` 环境。直接敲 `python` 可能用到 conda 那个，
**它没装 `pyarrow` 和 `duckdb`，清洗会失败**。永远写全路径：

```bat
.venv\Scripts\python ...          在根目录
..\.venv\Scripts\python ...       在 downloads/ 目录
```

### ③ `allowed` 是"权限清单"，不是"任务清单"

往 `allowed` 里加两家店，**不会**让 `run_downloads.py` 下载两家——那个脚本
设计上只跑一家。要下多家请用 `run_batch.py`。脚本本身会提醒你：

```
note: running BOCINA_TA02 only. 1 other store(s) allowed (UNIT_TA04) - use run_batch.py
```

### ④ `run_downloads.py` 不做清洗、不出对账簿

它是单店调试工具。要端到端出结果，一律用 `run_batch.py --stores 店名`。

### ⑤ 对账簿被 Excel 占用会失败

生成对账簿时如果同名文件正在 Excel 里开着，会写入失败并记为该店的错误。
跑批之前先把 `downloads/data/reports/` 下的工作簿都关掉。

### ⑥ 定时任务不能用"不登录也运行"

紫鸟是带界面的 Electron 程序，需要真实桌面。Windows 计划任务必须选
**"只在用户登录时运行"**，并且那台机器要保持登录、不休眠。

---

## 6. 看懂运行结果

跑完会打印一张表：

```
  STORE            STATUS        FILES  NOTE
  清货2_EE02         ok               19  workbook
  EWTTO_SM         partial          14  1 error(s): ... | workbook
  NARWAL_EE01      ok               17  1 still generating | workbook
  UNIT_TA04        ok               17  1 no movements | workbook
```

| 状态 | 含义 |
|---|---|
| `ok` | 正常 |
| `partial` | 下到了文件，但有路线出错——**要看一眼** |
| `failed` | 这家店整个没跑起来 |
| `not-allowed` | 不在白名单里 |

备注列的含义：

- **`still generating`**：MercadoPago 还在生成，请求已提交、不会丢，下次
  `--collect-only` 就能取回。
- **`no movements`**：所选时间段内没有任何流水，MercadoPago 不会生成这张表。
  **这不是错误**，换个时间段才可能有。
- **`workbook`**：这家店的对账簿已生成。

退出码：`0` 全部正常；`1` 有店需要关注；`2` 紫鸟客户端出错；`3` 店名不在白名单；`4` 配置文件读不了。

---

## 7. 两个不是代码问题的坑

**MercadoPago 报表权限。** 有些店的登录账号没有报表权限，页面会显示
*"No tienes permisos para ver los reportes"*。脚本会明确报出来，但这需要
**账号管理员**在 MercadoPago 的「Colaboradores」里授予两项权限：
*Acceder a reportes de tus cobros y facturación* 和 *Acceder a reportes de operaciones*。
代码绕不过去。

**报表生成慢。** `Liberaciones`（release）这张表在多家店都出现过要几小时的情况。
脚本不会一直干等——记进 `logs/pending.json`，之后补取。

---

## 8. 项目结构

| 文件 | 作用 |
|---|---|
| `run_batch.py` | **多店入口**，日常就用它 |
| `run_downloads.py` | 单店入口；`run_store()` 是两个入口共用的主流程 |
| `store_config.py` | 店铺白名单（读 `config.json`） |
| `ziniao_client.py` | 紫鸟客户端生命周期、HTTP IPC、Selenium 附着 |
| `meli_forms.py` | MercadoLibre 三条下载路线 |
| `mercadopago.py` | MercadoPago 6 种报表 |
| `pending_store.py` | 已请求但还没取回的报表清单 |
| `transform.py` | 调用 `downloads/mx_sales` 清洗管线的接口 |
| `run_report.py` | 往每个下载目录写 `REPORT.md` |
| `bootstrap.py` | 新机器环境自检 |
| `downloads/mx_sales/` | **独立子项目**：清洗 + 对账规则（有自己的 README 和测试） |

---

## 9. 想了解更多

| 文档 | 内容 |
|---|---|
| `GUIDE.md` | 完整运维指南：连接原理、每条路线的实现细节、踩过的坑 |
| `FORMS.md` | 15 张表单逐一说明：来源页面、怎么点出来的、文件长什么样 |
| `downloads/README.md` | 清洗管线与 9 条对账规则 |

---

## 10. 出问题了先做这三步

1. `.venv\Scripts\python bootstrap.py` —— 环境是不是好的
2. `.venv\Scripts\python run_batch.py --dry-run` —— 配置读出来对不对
3. 看 `logs/batch_<时间戳>.json` —— 到底哪家店、哪条路线出的问题

绝大多数"跑不起来"都是前两步能查出来的：Python 用错了、目录不对、
或者紫鸟客户端没装/没开 WebDriver 权限。
