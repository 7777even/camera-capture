#!/usr/bin/env python
"""
camera-capture 技能 CLI 包装器 — 海康摄像头截图 + 质量检测

支持两种调用方式：
  1. 命令行参数：--host ... --camera-code ...
  2. JSON 标准输入：echo '{"host":"..."}' | python capture.py --stdin

输出 JSON 格式结果到 stdout。
"""
import json
import sys
import os

# 从同目录导入摄像头截图模块
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from camera_capture import capture_single


def parse_args():
    import argparse
    parser = argparse.ArgumentParser(description="海康摄像头截图 + 质量检测")
    parser.add_argument("--host", help="海康平台地址")
    parser.add_argument("--port", type=int, default=443, help="端口(默认443)")
    parser.add_argument("--app-key", help="appKey")
    parser.add_argument("--app-secret", help="appSecret")
    parser.add_argument("--camera-code", help="摄像头编码")
    parser.add_argument("--camera-name", default="", help="摄像头名称")
    parser.add_argument("--save-dir", default="./screenshots", help="截图保存目录")
    parser.add_argument("--timeout", type=int, default=10, help="RTSP超时秒数")
    parser.add_argument("--retry", type=int, default=2, help="重试次数")
    parser.add_argument("--warmup", type=float, default=1.5, help="预热秒数(慢加载摄像头可加大)")
    parser.add_argument("--api-path", default="/artemis", help="API前缀")
    parser.add_argument("--json", action="store_true", help="输出JSON格式结果")
    parser.add_argument("--stdin", action="store_true", help="从stdin读取JSON参数")
    return parser.parse_args()


def main():
    args = parse_args()

    if args.stdin:
        raw = sys.stdin.read()
        if not raw.strip():
            print(json.dumps({"error": "stdin 输入为空", "status": "error"}, ensure_ascii=False))
            sys.exit(1)
        params = json.loads(raw)
    else:
        params = {
            "host": args.host,
            "app_key": args.app_key,
            "app_secret": args.app_secret,
            "camera_index_code": args.camera_code,
            "camera_name": args.camera_name,
            "save_dir": args.save_dir,
            "timeout": args.timeout,
            "retry_count": args.retry,
            "warmup_seconds": args.warmup,
            "api_path": args.api_path,
            "port": args.port,
        }

    required = ["host", "app_key", "app_secret", "camera_index_code"]
    missing = [k for k in required if not params.get(k)]
    if missing:
        result = {"error": f"缺少必填参数: {missing}", "status": "error"}
        print(json.dumps(result, ensure_ascii=False, indent=2))
        sys.exit(1)

    try:
        result = capture_single(
            host=params["host"],
            port=int(params.get("port", 443)),
            app_key=params["app_key"],
            app_secret=params["app_secret"],
            camera_index_code=params["camera_index_code"],
            camera_name=params.get("camera_name", ""),
            save_dir=params.get("save_dir", "./screenshots"),
            timeout=int(params.get("timeout", 10)),
            retry_count=int(params.get("retry_count", 2)),
            warmup_seconds=float(params.get("warmup_seconds", 1.5)),
            api_path=params.get("api_path", "/artemis"),
        )
    except Exception as e:
        result = {
            "status": "error",
            "error": str(e),
            "screenshotPath": None,
            "qualityScore": 0.0,
            "qualityDetail": {},
            "captureTime": "",
        }

    print(json.dumps(result, ensure_ascii=False, indent=2))

    if result.get("status") in ("offline", "abnormal", "error"):
        sys.exit(1)


if __name__ == "__main__":
    main()
