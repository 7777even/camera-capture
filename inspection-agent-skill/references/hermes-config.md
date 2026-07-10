# Hermes 定时调度配置

## 概述

Hermes 是 WorkBuddy 的内置调度引擎，支持自然语言解析定时规则，自动触发已安装的技能包（Expert）执行。

## 调度示例

| 用户指令 | 解析结果 |
|---------|---------|
| "每天下午三点执行巡检" | `FREQ=DAILY;BYHOUR=15;BYMINUTE=0` |
| "十分钟后执行巡检" | 一次性调度，当前时间 + 10min |
| "明早八点巡检" | 一次性调度，明天 08:00 |
| "每周一上午九点巡检" | `FREQ=WEEKLY;BYDAY=MO;BYHOUR=9;BYMINUTE=0` |

## 触发流程

用户在 WorkBuddy 上安装 `inspection-agent` 技能包后，通过自然语言即可创建定时任务：

```
"每天下午三点执行摄像头每日巡检"
```
```
"每 60 秒 执行摄像头每日巡检"
```

Hermes 自动解析时间和动作，在到达触发时间时启动 Agent，Agent 加载技能包的 SKILL.md 指令，执行 5 步巡检流程：

```
到达触发时间 → WorkBuddy 启动 Agent
  ├── 加载 inspection-agent 技能指令
  ├── Step 1: 读取摄像头清单 Excel
  ├── Step 2: 并发截图（ThreadPoolExecutor + camera_capture）
  ├── Step 3: 汇总结果（在线/离线/异常）
  ├── Step 4: 更新 Excel 状态列
  └── Step 5: 飞书通知 → 退出
```

## 快速执行

技能包附带了独立编排脚本，可一键执行：

```bash
cd <项目目录>
python scripts/orchestrate.py --config config.json
```

## 首次部署检查清单

- [ ] 确认 `config.json` 中海康平台参数正确（host/app_key/app_secret）
- [ ] 确认摄像头清单 Excel 文件存在且格式正确（文件名见 `config/branding.json` 的 `default_excel_name`）
- [ ] 安装依赖：`pip install -r requirements.txt`
- [ ] 测试单路截图：`python scripts/orchestrate.py --config config.json --test-single <camera_code>`
- [ ] 测试全量巡检：`python scripts/orchestrate.py --config config.json`
- [ ] 确认飞书 webhook URL 有效
- [ ] 在 WorkBuddy 中输入指令创建定时任务

## 排障

| 问题 | 可能原因 | 解决 |
|------|----------|------|
| 全部摄像头离线 | 海康平台不可达 / 网络不通 | 检查 host 和端口；ping 测试 |
| 部分摄像头离线 | 单个摄像头断电/断网 | 人工核实现场 |
| Excel 未更新 | openpyxl 未安装 / 文件被占用 | `pip install openpyxl`；关闭 Excel |
| 飞书未收到通知 | webhook URL 失效 | 检查飞书机器人配置 |
| HEVC 花屏 | OpenCV 解码兼容性问题 | camera-capture 已内置跳帧逻辑 |
