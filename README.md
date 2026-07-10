# 海康摄像头巡检技能包

两个 WorkBuddy 技能，用于海康威视摄像头的截图、质量检测和每日巡检：

- `camera-capture-skill/` — 海康摄像头截图 + 质量检测（原子能力，一次处理一路摄像头）
- `inspection-agent-skill/` — 摄像头每日巡检编排（读取清单 → 并发截图 → 质量检测 → 更新 Excel / 飞书通知）

## 凭证配置（重要）

两个技能包在 `config/platform.json` 中存放海康平台凭证。**该文件不纳入版本库**（已在 `.gitignore` 中忽略，避免泄露密钥）。

首次使用时，复制模板并填入你自己的凭证：

```bash
cp camera-capture-skill/config/platform.template.json camera-capture-skill/config/platform.json
cp inspection-agent-skill/config/platform.template.json inspection-agent-skill/config/platform.json
# 编辑上面两个 platform.json，填入 host / app_key / app_secret（以及飞书 webhook）
```

## 依赖

```bash
pip install opencv-python requests openpyxl python-docx pillow
```

## 目录结构

```
camera-capture-skill/      海康截图原子能力
  SKILL.md
  scripts/camera_capture.py
  config/{platform.template.json, branding.json}
inspection-agent-skill/    每日巡检编排
  SKILL.md
  scripts/orchestrate.py
  config/{platform.template.json, branding.json}
```
