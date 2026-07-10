---
name: camera-capture
description: "海康威视摄像头截图 + 质量检测。传入摄像头编码和平台凭证，自动调用海康 Artemis API 获取 RTSP 流地址，使用 OpenCV 抓取一帧画面并保存到指定目录，返回在线/离线/异常状态和质量评分。此技能为原子能力——每次只处理一路摄像头。当需要\"截图摄像头\"、\"抓拍\"、\"摄像头巡检\"、\"海康截图\"时使用。"
allowed-tools: [Bash]
---

# 海康威视摄像头截图技能

## 品牌配置（config/branding.json）

本项目使用 `config/branding.json` 管理业务品牌相关的可配置文本。
切换业务场景时只需修改该文件，无需改动任何代码或 SKILL.md。

| 配置项 | 说明 | 当前值 |
|--------|------|--------|
| `project_label` | 项目业务标签 | 环保危废 |
| `example_camera_name` | 文档示例用的摄像头名称 | 危废仓库1 |

> **注意**：本技能为原子能力，`branding.json` 仅用于文档示例展示，不影响运行时行为。

## ⚠️ 核心规则：禁止自动读取配置文件

**此技能为原子能力，所有参数必须由调用方显式传入。禁止从以下位置自动读取：**
- ❌ 禁止从 `config/platform.json` 自动加载凭据
- ❌ 禁止使用任何硬编码的默认凭据
- ❌ 禁止从环境变量自动读取

**调用方（如 `inspection-agent`）负责收集用户提供的凭据并传入。如果本技能被直接调用，必须先向用户询问所有必填参数。**

---

## 功能

对单个海康威视摄像头执行截图 + 质量检测流程：

1. 调用海康 Artemis API（HMAC-SHA256 签名）获取 RTSP 预览地址
2. OpenCV 连接 RTSP 流，跳过前 15 帧避开 HEVC 花屏
3. 抓取有效帧 → 保存 JPEG
4. 质量检测（Laplacian 清晰度 + 亮度 + 颜色多样性加权评分）
5. 质量不合格时抓第 2 帧复检
6. 返回结构化结果

## 编码约定

**所有文件读写操作必须显式指定 UTF-8 编码**。Hermes 调度引擎在部分环境下可能默认使用 GBK 编码，因此在 `open()` 调用中必须指定 `encoding='utf-8'`。

## 输入参数

所有参数由调用方传入，本技能不自动加载任何配置文件。

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| host | string | 是 | 海康平台 IP/域名 |
| port | int | 否 | 端口（默认 443） |
| app_key | string | 是 | 海康平台 appKey |
| app_secret | string | 是 | 海康平台 appSecret |
| camera_index_code | string | 是 | 摄像头编码 |
| camera_name | string | 否 | 摄像头名称（用于文件名） |
| save_dir | string | 否 | 截图保存目录（默认 ./screenshots） |
| timeout | int | 否 | RTSP 超时秒数/次（默认 10） |
| retry_count | int | 否 | 重试次数（默认 2） |

## 🚨 强制规则：必须使用技能包自带脚本

**禁止 AI 智能体自行新建 Python 脚本来调用本技能。** 所有调用必须通过技能包自带的 CLI 脚本 `scripts/capture.py` 执行：

- ✅ 允许：通过命令行运行 `python scripts/capture.py ...` 或 `echo '{...}' | python scripts/capture.py --stdin`
- ✅ 允许：Python 中 `import` 技能包模块调用 `capture_single()`（需 `sys.path.insert(0, "scripts")`）
- ❌ 禁止：在临时目录或工作目录中新建 `.py` 文件，把参数 hard-code 进去再执行
- ❌ 禁止：复制 SKILL.md 中的示例代码到新文件执行

**理由**：技能包自带的 `scripts/capture.py` 和 `scripts/camera_capture.py` 已经过完整的参数校验、错误处理和输出格式化。自行新建脚本会丢失这些保障，且可能导致参数传递错误。

---

## 调用方式

**⚠️ 执行任何调用之前，必须先安装依赖：**

```bash
pip install requests>=2.28.0 opencv-python>=4.8.0 numpy>=1.24.0
```

| 包 | 用途 |
|----|------|
| `requests` | HTTP 请求（海康 API 调用） |
| `opencv-python` | RTSP 视频流读取、图像质量检测 |
| `numpy` | 图像数值计算 |

---

### 作为子技能被调用（推荐方式）

由 `inspection-agent` 传入参数调用 `capture_single()`：

```python
from camera_capture import capture_single

result = capture_single(
    host="<用户提供的 host>",
    port=443,
    app_key="<用户提供的 appKey>",
    app_secret="<用户提供的 appSecret>",
    camera_index_code="CAM-001",
    camera_name="<摄像头名称，如 config/branding.json 中的示例>",
    save_dir="<用户提供的保存目录>",
)
```

### 独立命令行调用

如果直接调用此技能，必须先向用户询问所有必填参数：

```bash
python scripts/capture.py \
  --host <用户提供的地址> --port 443 \
  --app-key <用户提供的appKey> --app-secret <用户提供的appSecret> \
  --camera-code CAM-001 --camera-name "<摄像头名称>" \
  --save-dir <用户提供的目录> --json
```

### JSON 管道调用

```bash
echo '{"host":"...","app_key":"...","app_secret":"...","camera_index_code":"CAM-001"}' \
  | python scripts/capture.py --stdin
```

## 输出结果

```json
{
  "status": "online | offline | abnormal",
  "screenshotPath": "/path/to/screenshot.jpg",
  "qualityScore": 0.82,
  "qualityDetail": { "laplacianScore": 0.85, "brightnessScore": 0.90, "colorDiversityScore": 0.72 },
  "errorMsg": null,
  "captureTime": "2026-06-25 15:01:23",
  "retryUsed": 0
}
```

### 状态说明

| 状态 | 含义 | 条件 |
|------|------|------|
| online | 在线 | RTSP 连接成功 + 画面质量 >= 0.4 |
| offline | 离线 | RTSP 超时重试全部失败 |
| abnormal | 异常 | 画面质量 < 0.4（模糊/黑屏/全灰/花屏） |

### 质量评分公式

```
总分 = Laplacian清晰度得分 × 0.5 + 亮度得分 × 0.3 + 颜色多样性得分 × 0.2
满分 1.0，低于 0.4 判定为异常
全灰画面（OpenCV解码异常）→ 一票否决，直接判定 abnormal
```


