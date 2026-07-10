---
name: inspection-agent
description: "摄像头每日巡检助手。执行完整流程：从 Excel 读取摄像头清单 → 并发对海康摄像头截图 → 质量检测 → 汇总在线/离线/异常状态 → 更新 Excel 状态列 → 更新 Word 巡查台账（截图贴入对应摄像头行）→ 发送飞书通知报告。TRIGGER when: \"执行巡检\"、\"摄像头巡检\"、\"inspection\"。品牌文本（巡检报告标题、飞书卡片抬头等）在 config/branding.json 中配置，切换场景只需改配置。"
allowed-tools: [Bash, Read, Write, Glob, Grep]
---

# 摄像头巡检助手

## 品牌配置（config/branding.json）

本项目使用 `config/branding.json` 管理所有业务品牌相关的文本。
**切换业务场景时只需修改该文件，无需改动任何代码或 SKILL.md。**

| 配置项 | 说明 | 示例值（环保危废场景） |
|--------|------|----------------------|
| `project_label` | 项目业务标签 | 环保危废 |
| `feishu_card_title` | 飞书巡检报告卡片标题 | 🔔 环保危废摄像头巡检报告 |
| `feishu_card_footer` | 飞书卡片页脚 | 环保小脑巡检助手 · 自动生成 |
| `inspection_start_banner` | 启动横幅文本 | 环保危废摄像头巡检 |
| `trigger_keywords_extra` | SKILL.md 描述中触发词补充说明 | 环保危废检查、危废仓库检查 |
| `default_excel_name` | 文档示例中引用的默认清单文件名 | 环保小脑摄像头清单.xlsx |
| `default_docx_name` | 文档示例中引用的默认台账文件名 | 危废仓库巡查台账_新模版.docx |

> **换场景示例**：将 `branding.json` 中的 `project_label` 改为 "安全生产"，
> `feishu_card_title` 改为 "🔔 安全生产摄像头巡检报告"，即可适配新场景。
> 不再需要修改 Python 脚本或 SKILL.md 中的任何硬编码文本。

## ⚠️ 核心规则：禁止自动读取任何配置文件

**每次执行巡检前，必须向用户逐一询问并收集以下所有信息。禁止从以下位置自动读取：**
- ❌ 禁止从 `config/platform.json` 自动加载
- ❌ 禁止从环境变量自动读取
- ❌ 禁止使用任何默认路径（如 `./摄像头清单.xlsx`）
- ❌ 禁止猜测文件位置

**所有路径、凭据、配置项，必须由用户在当前对话中明确提供。**

---

## 执行前信息收集（必须先完成）

收到巡检指令后，必须分两轮向用户收集信息，不得跳过：

### 第一轮：必填项（全部必答）

使用 `ask_followup_question` 工具一次性收集：

| 问题 | 说明 |
|------|------|
| 摄像头清单文件路径 | `.xlsx` 或 `.csv` 文件的绝对路径 |
| 截图保存目录 | 截图和 JSON 结果保存到哪个目录 |
| 海康平台地址 | IP 地址或域名 |
| appKey | 海康 Artemis 平台 appKey |

### 第二轮：补充信息（部分可跳过）

收到第一轮答案后，继续收集：

| 问题 | 说明 | 默认值 |
|------|------|--------|
| appSecret | 海康 Artemis 平台 appSecret | 必填 |
| 端口 | HTTPS 端口 | 443 |
| Word 台账模版路径 | 需要更新 Word 台账时提供，不需要则留空 | 空（跳过台账步骤） |
| 飞书 Webhook URL | 需要飞书通知时提供，不需要则留空 | 空（跳过通知步骤） |

### 第三轮：确认

将所有收集到的信息汇总，展示给用户确认。确认无误后才能开始执行。

---

## 功能概述

对工业园区的海康威视摄像头执行巡检，完成以下 6 个步骤：

1. **读取摄像头清单** — 从用户提供的 Excel/CSV 文件中解析摄像头名称和编码
2. **并发截图** — 同时对所有摄像头截图 + 质量检测
3. **汇总结果** — 统计在线/离线/异常数量，判断大面积离线等异常
4. **更新 Excel** — 按日期追加状态列，保留历史记录
5. **更新 Word 巡查台账** — 将截图贴入对应摄像头名称行的"监控截图"列（仅用户提供了模版时执行）
6. **飞书通知** — 发送卡片消息推送巡检报告（仅用户提供了 webhook 时执行）

## 编码约定

**所有文件读写操作必须显式指定 UTF-8 编码**。Hermes 调度引擎在部分环境下可能默认使用 GBK 编码，
因此在 `open()` 调用中必须指定 `encoding='utf-8'`，`openpyxl` 读取 `.xlsx` 时无需额外指定（二进制格式）。

---

## 🚨 强制规则：必须使用技能包自带脚本

**禁止 AI 智能体自行新建 Python 脚本来执行巡检流程。** 所有执行必须通过技能包自带的 CLI 脚本 `scripts/orchestrate.py`：

- ✅ 允许：通过命令行运行 `python scripts/orchestrate.py ...`
- ❌ 禁止：在临时目录或工作目录中新建 `.py` 文件，把参数 hard-code 进去再执行
- ❌ 禁止：复制 SKILL.md 中的示例代码到新文件执行

本 SKILL.md 中的 Python 代码块仅为**逻辑说明和文档参考**，不是让你复制出来执行的。实际执行请使用 `scripts/orchestrate.py`，所有步骤（读取清单 → 并发截图 → 更新 Excel → 更新 Word台账 → 飞书通知）已在该脚本中完整实现。

同理，截图能力通过 `camera-capture` 技能包自带的 `scripts/capture.py` 调用，不要绕开它自行写 RTSP 抓帧代码。

---

执行时使用技能包自带脚本 `scripts/orchestrate.py`，将所有用户提供的参数通过命令行传入。

**⚠️ 重要**：在执行任何步骤之前，必须先完成上述"执行前信息收集"流程。以下代码中的路径和凭据均来自用户输入，不可使用硬编码默认值。

---

### Step 0：安装 Python 依赖（必须最先执行）

执行任何业务逻辑之前，必须先安装所需依赖：

```bash
pip install requests>=2.28.0 opencv-python>=4.8.0 numpy>=1.24.0 openpyxl>=3.1.0 python-docx>=1.1.0
```

或使用技能包自带的 `requirements.txt`：

```bash
pip install -r requirements.txt
```

依赖说明：
| 包 | 用途 |
|----|------|
| `requests` | HTTP 请求（海康 API 调用、飞书通知） |
| `opencv-python` | RTSP 视频流读取、图像质量检测 |
| `numpy` | 图像数值计算 |
| `openpyxl` | 读取/写入摄像头清单 Excel |
| `python-docx` | 更新 Word 巡查台账 |

---

### Step 1：读取摄像头清单

从用户提供的路径读取摄像头清单文件。

```python
import os
from openpyxl import load_workbook

def is_enabled(val):
    """判断"是否更新到台账"列：是/true/1/y → True"""
    if val is None:
        return False
    return str(val).strip().lower() in ("是", "true", "1", "y", "yes", "√", "✓")

excel_path = "<用户提供的摄像头清单路径>"
ext = os.path.splitext(excel_path)[1].lower()

cameras = []

if ext == ".csv":
    # CSV 格式：显式指定 UTF-8 编码（兼容 BOM 头）
    import csv
    with open(excel_path, "r", encoding="utf-8-sig") as f:
        reader = csv.reader(f)
        next(reader)  # 跳过表头
        for row in reader:
            if len(row) >= 2 and row[0] and row[1]:
                cameras.append({
                    "name": str(row[1]).strip(),
                    "code": str(row[0]).strip(),
                    "ledger_enabled": is_enabled(row[2]) if len(row) > 2 else False,
                })
else:
    # XLSX 格式：openpyxl 处理二进制格式，无需指定编码
    wb = load_workbook(excel_path)
    ws = wb.active
    for row in ws.iter_rows(min_row=2, values_only=True):
        code, name = row[0], row[1]
        if name and code:
            cameras.append({
                "name": str(name).strip(),
                "code": str(code).strip(),
                "ledger_enabled": is_enabled(row[2]) if len(row) > 2 else False,
            })

print(f"共 {len(cameras)} 个摄像头")
print(f"其中 {sum(1 for c in cameras if c['ledger_enabled'])} 个标记为更新台账")
```

清单格式：
| 列 | 字段 | 说明 |
|----|------|------|
| A | 摄像头编码 | 传给 Hikvision API |
| B | 摄像头名称 | 用于文件命名和台账匹配 |
| C | 是否更新到台账 | **是**=截图写入Word台账，空/否=跳过。运营人员维护此列即可控制台账内容 |
| D 起 | 每日状态 | 系统自动填入（如 D=2026/6/30，值为"在线/离线/异常"），按日追加新列 |

### Step 2：并发截图（核心步骤）

使用用户提供的平台凭据，对所有摄像头并发执行截图。

**关键规则**：
- 使用 `ThreadPoolExecutor` **同时**对所有摄像头发起截图，不要串行
- 每个任务独立超时，互不影响
- **同步屏障**：必须等全部完成后再汇总

```python
import json, os, sys
from concurrent.futures import ThreadPoolExecutor, as_completed

# 使用用户提供的凭据（来自询问环节）
config = {
    "host": "<用户提供的 host>",
    "port": <用户提供的 port>,
    "app_key": "<用户提供的 appKey>",
    "app_secret": "<用户提供的 appSecret>",
    "save_dir": "<用户提供的保存目录>",
    "timeout": 10,
    "retry_count": 2,
    "max_workers": 30,
    "api_path": "/artemis",
}

# 导入截图能力（技能包已自带 camera_capture 模块）
sys.path.insert(0, "scripts")
from camera_capture import capture_single

def run_one(cam):
    result = capture_single(
        host=config["host"],
        port=config.get("port", 443),
        app_key=config["app_key"],
        app_secret=config["app_secret"],
        camera_index_code=cam["code"],
        camera_name=cam["name"],
        save_dir=config.get("save_dir", "./screenshots"),
        timeout=config.get("timeout", 10),
        retry_count=config.get("retry_count", 2),
        api_path=config.get("api_path", "/artemis"),
    )
    result["cameraName"] = cam["name"]
    result["cameraCode"] = cam["code"]
    return result

max_workers = min(config.get("max_workers", 30), len(cameras))
all_results = []
with ThreadPoolExecutor(max_workers=max_workers) as pool:
    futures = {pool.submit(run_one, c): c for c in cameras}
    for future in as_completed(futures):
        cam = futures[future]
        try:
            all_results.append(future.result(timeout=45))
        except Exception as e:
            all_results.append({
                "status": "offline", "cameraName": cam["name"], "cameraCode": cam["code"],
                "screenshotPath": None, "qualityScore": 0.0, "errorMsg": f"线程异常: {e}",
            })
```

### Step 3：汇总结果

```python
online = [r for r in all_results if r["status"] == "online"]
offline = [r for r in all_results if r["status"] == "offline"]
abnormal = [r for r in all_results if r["status"] == "abnormal"]
total = len(all_results)

large_offline = len(offline) / total > 0.5 if total > 0 else False
all_failed = total > 0 and len(online) == 0

print(f"总计: {total}  |  在线: {len(online)}  |  离线: {len(offline)}  |  异常: {len(abnormal)}")
```

**异常升级策略**：
| 条件 | 动作 |
|------|------|
| >50% 摄像头离线 | 通知中附加"⚠️ 大面积离线警告" |
| 全部摄像头离线 | 追加"🚨 全部离线——可能是平台/网络中断" |
| 单路失败 | 标记为离线/异常，不影响其他路 |

### Step 4：更新 Excel

在用户提供的摄像头清单上按日期追加状态列。

Excel 结构：A=编码，B=名称，C=台账标记，D 起为每日状态列。匹配时优先按名称（B 列），回退按编码（A 列）。

```python
from datetime import date
from openpyxl import load_workbook

excel_path = "<用户提供的摄像头清单路径>"
wb = load_workbook(excel_path)
ws = wb.active
today = date.today().strftime("%Y%m%d")

# 扫描表头，找到今天的日期列；没有则在最右侧新增
headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
col = headers.index(today) + 1 if today in headers else ws.max_column + 1
if today not in headers:
    ws.cell(1, col, today)

status_cn = {"online": "在线", "offline": "离线", "abnormal": "异常"}
name_map = {r.get("cameraName"): r for r in all_results}
code_map = {r.get("cameraCode"): r for r in all_results}

for row in range(2, ws.max_row + 1):
    cam_code = ws.cell(row, 1).value   # A 列：编码
    cam_name = ws.cell(row, 2).value   # B 列：名称
    r = name_map.get(cam_name) or code_map.get(cam_code)
    if r:
        ws.cell(row, col, status_cn.get(r["status"], "未知"))

wb.save(excel_path)
print(f"Excel 已更新，列: {today}")
```

### Step 5：更新 Word 巡查台账（仅用户提供了模版时执行）

**仅当用户在询问环节提供了 Word 模版路径时才执行此步骤。**

**台账写入规则**：根据 Excel 清单中 E 列"是否更新到台账"（`ledger_enabled`）过滤摄像头。
只有标记为"是"的摄像头才会将其截图写入 Word 台账，其余跳过。
运营人员只需在 Excel 中维护该列即可控制台账内容，无需改代码。

```python
# 按 Excel E 列过滤：仅 ledgerEnabled=True 的摄像头写入台账
ledger_results = [r for r in all_results if r.get("ledgerEnabled", False)]
print(f"台账写入: {len(ledger_results)} 路，跳过: {len(all_results) - len(ledger_results)} 路")
```

文档结构：
- 第 0 行：巡查日期（自动更新为今天）
- 第 1 行：表头（序号 | 企业名称 | **摄像头名称** | 是否超阈值 | 其他异常情况 | **监控截图** | 处置情况）
- 第 2 行起：数据行，按摄像头名称精确匹配
- 列索引 5 = 监控截图（插入图片位置）

> 台账模版路径、输出路径等均在运行时通过命令行参数传入，不再硬编码文件名。
> 默认示例文件名见 `config/branding.json` → `default_docx_name`。

### Step 6：飞书通知（仅用户提供了 webhook 时执行）

**仅当用户在询问环节提供了飞书 webhook URL 时才执行此步骤。**
网络请求用 python 写，不要用 curl。

```python
import requests
from datetime import datetime

feishu_webhook = "<用户提供的 webhook URL>"

offline_text = "\n".join(
    f"❌ {r['cameraName']}：{r.get('errorMsg', '超时')}" for r in offline
) or "无"
abnormal_text = "\n".join(
    f"⚠️ {r['cameraName']}：{r.get('errorMsg', '画面质量差')}" for r in abnormal
) or "无"

alert = ""
if all_failed:
    alert = "\n🚨 **全部摄像头离线**：可能是海康平台停机或网络中断！"
elif large_offline:
    alert = "\n⚠️ **大面积离线警告**：超过50%摄像头离线，可能是平台或网络问题！"

body = {
    "msg_type": "interactive",
    "card": {
        "header": {"title": {"tag": "plain_text", "content": "<从 config/branding.json 读取 feishu_card_title>"}, "template": "blue"},
        "elements": [
            {"tag": "div", "text": {"tag": "lark_md", "content": (
                f"**执行时间**: {datetime.now().strftime('%Y-%m-%d %H:%M')}\n"
                f"**总计**: {total} 路  "
                f"**✅ 在线**: {len(online)}  "
                f"**❌ 离线**: {len(offline)}  "
                f"**⚠️ 异常**: {len(abnormal)}"
                f"{alert}"
            )}},
            {"tag": "hr"},
            {"tag": "div", "text": {"tag": "lark_md", "content": f"**❌ 离线摄像头 ({len(offline)})：**\n{offline_text}"}},
            {"tag": "hr"},
            {"tag": "div", "text": {"tag": "lark_md", "content": f"**⚠️ 异常摄像头 ({len(abnormal)})：**\n{abnormal_text}"}},
            {"tag": "hr"},
            {"tag": "note", "elements": [{"tag": "plain_text", "content": "<从 config/branding.json 读取 feishu_card_footer>"}]},
        ],
    },
}

resp = requests.post(feishu_webhook, json=body, timeout=10)
print(f"飞书通知: {resp.json()}")
```

### 保存结果

```python
import json, os
today = date.today().strftime("%Y%m%d")
save_dir = "<用户提供的保存目录>"
os.makedirs(save_dir, exist_ok=True)
with open(f"{save_dir}/inspection_{today}.json", "w", encoding="utf-8") as f:
    json.dump(all_results, f, ensure_ascii=False, indent=2)
```

---

## 依赖安装

执行前确保依赖已安装：

```bash
pip install requests opencv-python numpy openpyxl python-docx
```

## 退出码

| 退出码 | 含义 |
|--------|------|
| 0 | 所有摄像头在线 |
| 1 | 存在离线或异常摄像头 |

## 技能关系

此技能为**巡检编排 Expert**，依赖 `camera-capture` 技能提供的单路截图能力。两个技能配合使用：

- **camera-capture**：原子能力，每次处理一路摄像头
- **inspection-agent**（本技能）：编排调度，批量并发执行全部流程
