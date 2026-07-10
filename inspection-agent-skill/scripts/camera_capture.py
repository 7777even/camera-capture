"""
camera-capture skill — 海康摄像头截图 + 质量检测

纯原子能力：给定摄像头参数，返回截图和质量判定结果。
不关心：摄像头数量、目录组织、通知方式、Excel 更新。
只关心：能不能连上 RTSP、能不能抓到帧、画面质量如何。
"""

import base64
import hashlib
import hmac
import json
import os
import time
from datetime import datetime, timezone
from email.utils import formatdate

import cv2
import numpy as np
import requests
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# 全局 session：禁用环境代理（Windows 上企业代理/SSL 拦截会导致请求挂起）
_http = requests.Session()
_http.verify = False
_http.trust_env = False


# ──────────────────────────────────────────────
# 1. 海康 Artemis API 认证（5行签名格式）
# ──────────────────────────────────────────────

def _gmt_date() -> str:
    """生成 HTTP GMT 日期字符串，如 'Thu, 23 Apr 2026 03:00:00 GMT'"""
    return formatdate(timeval=time.time(), usegmt=True)


def sign_request(app_secret: str, method: str, path: str) -> dict:
    """
    构建海康 Artemis API 签名头（5行格式）。

    签名串 (5行，\\n 分隔):
        METHOD
        */*
        application/json
        {GMT-Date}
        {url-path}

    x-ca-key/x-ca-nonce 作为 HTTP 头发送，但不参与签名计算。
    """
    date = _gmt_date()

    # 5行签名串
    string_to_sign = (
        f"{method}\n"
        f"*/*\n"
        f"application/json\n"
        f"{date}\n"
        f"{path}"
    )

    sig_bytes = hmac.new(
        app_secret.encode("utf-8"),
        string_to_sign.encode("utf-8"),
        hashlib.sha256,
    ).digest()
    signature_b64 = base64.b64encode(sig_bytes).decode("utf-8")

    return {
        "x-ca-signature": signature_b64,
        "Date": date,
        "Content-Type": "application/json",
        "Accept": "*/*",
    }


def get_rtsp_url(host: str, port: int, app_key: str, app_secret: str,
                 camera_index_code: str, api_path: str = "/artemis") -> tuple:
    """
    调用海康 Artemis API 获取 RTSP 预览地址。

    接口: POST {api_path}/api/video/v1/cameras/previewURLs
    文档参考: 海康 Artemis 平台 API 参考 §2
    返回 (rtsp_url, error_msg)。
    """
    path = f"{api_path}/api/video/v1/cameras/previewURLs"
    url = f"https://{host}:{port}{path}"

    payload = {
        "cameraIndexCode": camera_index_code,
        "streamType": 0,         # 0=主码流
        "protocol": "rtsp",
        "expand": "streamform=rtp",
        "transmode": 1,          # 1=TCP（UDP在高并发下极易超时）
    }

    headers = sign_request(app_secret, "POST", path)
    headers["x-ca-key"] = app_key

    try:
        resp = _http.post(url, headers=headers, json=payload,
                          timeout=15)
        resp.raise_for_status()
        data = resp.json()

        if data.get("code") != "0":
            return "", f"API错误: {data.get('msg', data)}"

        rtsp_url = data.get("data", {}).get("url", "")
        if not rtsp_url:
            return "", "RTSP地址为空"

        return rtsp_url, None

    except requests.exceptions.Timeout:
        return "", "API请求超时"
    except requests.exceptions.ConnectionError:
        return "", "无法连接到海康平台"
    except Exception as e:
        return "", f"API调用异常: {str(e)}"


# ──────────────────────────────────────────────
# 2. RTSP 截图 + 重试
# ──────────────────────────────────────────────

def _is_corrupted(frame: np.ndarray) -> bool:
    """
    综合检测各种花屏：垂直/水平条纹、灰屏/黑屏（全部或部分）、撕裂/拉伸。
    宁可误判（丢弃正常帧）也不要保存花屏帧。
    """
    if frame is None or frame.size == 0:
        return True

    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY) if len(frame.shape) == 3 else frame
    h, w = gray.shape
    if h < 10 or w < 10:
        return True

    # ── 1. 条纹检测（降低阈值，更敏感） ──
    col_diff = np.abs(gray[:, 1:] - gray[:, :-1]).astype(np.float32)
    col_diff_mean = col_diff.mean(axis=0)
    if len(col_diff_mean) > 10:
        if float(col_diff_mean.mean()) > 3.0 and float(col_diff_mean.std()) < 3.0:
            return True

    row_diff = np.abs(gray[1:, :] - gray[:-1, :]).astype(np.float32)
    row_diff_mean = row_diff.mean(axis=1)
    if len(row_diff_mean) > 10:
        if float(row_diff_mean.mean()) > 3.0 and float(row_diff_mean.std()) < 3.0:
            return True

    # ── 2. 局部灰屏/黑屏检测（覆盖"上半正常+下半灰色"型花屏） ──
    h_third = max(h // 3, 1)
    band_vars = []
    for i in range(3):
        y_start = i * h_third
        y_end = (i + 1) * h_third if i < 2 else h
        band = gray[y_start:y_end, :]
        if band.size > 0:
            band_vars.append(float(band.var()))

    if len(band_vars) == 3:
        min_var = min(band_vars)
        max_var = max(band_vars)
        # 某条带灰屏（方差<30）而其他条带正常（>150）→ 花屏
        if min_var < 30.0 and max_var > 150.0:
            return True

    # ── 3. 整帧灰屏/黑屏 ──
    if float(gray.var()) < 50:
        return True

    # ── 4. 撕裂/拉伸检测（更细粒度的块） ──
    block_h, block_w = max(h // 6, 1), max(w // 6, 1)
    block_vars = []
    for y in range(0, h, block_h):
        for x in range(0, w, block_w):
            block = gray[y:y+block_h, x:x+block_w]
            if block.size > 0:
                block_vars.append(float(block.var()))

    if len(block_vars) >= 4:
        block_vars_arr = np.array(block_vars)
        zero_ratio = float(np.sum(block_vars_arr < 15.0)) / len(block_vars_arr)
        high_ratio = float(np.sum(block_vars_arr > 500.0)) / len(block_vars_arr)
        if zero_ratio > 0.25 and high_ratio > 0.08:
            return True

    return False


def capture_frame(rtsp_url: str, timeout: int = 15, warmup_seconds: float = 5.0) -> tuple:
    """
    从 RTSP 流抓取一帧（终极版）。

    反花屏策略（针对壳牌/可隆等慢加载摄像头）：
    1. FFmpeg 参数调优：增大缓冲区和分析时长，加速 I-frame 到达
    2. 时间预热：丢弃前 N 秒所有帧，等待解码器拿到关键帧
    3. 强制等待期：预热后再等 1.5 秒，清理解码器残留
    4. 综合花屏检测：条纹 + 灰屏/黑屏 + 撕裂，每帧都检测
    5. 帧稳定检测：连续 2 帧差异低 → 流已稳定
    6. 质量校验：方差+饱和度双重验证
    7. 最佳帧降级：超时后只返回**非花屏**的方差最高帧

    返回 (got_frame: bool, frame: np.ndarray | None, error: str | None, frames_seen: int)
    """
    # 增大 FFmpeg 缓冲和分析参数，加速慢摄像头解码
    os.environ["OPENCV_FFMPEG_CAPTURE_OPTIONS"] = (
        "rtsp_transport;tcp|stimeout;5000000|"
        "analyzeduration;10000000|probesize;10000000|"
        "max_delay;10000000|buffer_size;1310720"
    )

    cap = cv2.VideoCapture(rtsp_url)
    cap.set(cv2.CAP_PROP_BUFFERSIZE, 5)

    start = time.time()
    prev_frame = None
    stable_count = 0
    frames_seen = 0
    best_frame = None
    best_var = -1
    post_warmup_extra = 1.5  # 预热后额外等待

    while time.time() - start < timeout:
        ret, f = cap.read()
        if not ret or f is None:
            time.sleep(0.1)
            continue
        frames_seen += 1

        elapsed = time.time() - start

        # 阶段1: 预热期（丢弃所有帧，等待 I-frame）
        if elapsed < warmup_seconds:
            # 预热期内仍追踪非花屏最佳帧，确保慢摄像头有备选
            if not _is_corrupted(f):
                gray = cv2.cvtColor(f, cv2.COLOR_BGR2GRAY)
                var = float(gray.var())
                if var > best_var:
                    best_var = var
                    best_frame = f.copy()
            continue

        # 阶段1.5: 预热后额外等待，清理解码器残留
        if elapsed < warmup_seconds + post_warmup_extra:
            if not _is_corrupted(f):
                gray = cv2.cvtColor(f, cv2.COLOR_BGR2GRAY)
                var = float(gray.var())
                if var > best_var:
                    best_var = var
                    best_frame = f.copy()
            continue

        # 阶段2: 综合花屏检测
        if _is_corrupted(f):
            time.sleep(0.1)
            continue

        gray = cv2.cvtColor(f, cv2.COLOR_BGR2GRAY)
        var = float(gray.var())

        # 更新最佳帧（只接受非花屏帧）
        if var > best_var:
            best_var = var
            best_frame = f.copy()

        # 方差极低 → 丢弃
        if var < 50:
            time.sleep(0.1)
            continue

        # 阶段3: 帧稳定检测
        if prev_frame is not None:
            diff = cv2.absdiff(f, prev_frame)
            mean_diff = float(diff.mean())
            if mean_diff < 15.0:
                stable_count += 1
            else:
                stable_count = max(0, stable_count - 1)
        prev_frame = f.copy()

        if stable_count < 2:
            time.sleep(0.05)
            continue

        # 阶段4: 质量校验
        hsv = cv2.cvtColor(f, cv2.COLOR_BGR2HSV)
        mean_sat = float(hsv[:, :, 1].mean())

        if var >= 80 and mean_sat >= 15:
            cap.release()
            return True, f, None, frames_seen

        if var >= 150 and mean_sat >= 5:
            cap.release()
            return True, f, None, frames_seen

        time.sleep(0.05)

    cap.release()

    # 超时：返回最佳非花屏帧（若有）
    if best_frame is not None:
        return True, best_frame, (
            f"预热{warmup_seconds:.0f}s+稳定{post_warmup_extra:.0f}s后未捕获到合格帧(共{frames_seen}帧,bestVar={best_var:.0f})"
        ), frames_seen

    return False, None, f"RTSP无数据(超时{timeout}s)", 0


def capture_with_retry(rtsp_url: str, timeout: int = 15,
                       retry_count: int = 3,
                       warmup_seconds: float = 5.0) -> tuple:
    """
    带重试的 RTSP 截图。
    每次重试都是全新的 RTSP 连接，避免复用已污染的解码器状态。
    策略：质量门通过 → 立即返回；最佳帧降级 → 继续重试期待更优帧。

    返回 (success, frame, retry_used, error, total_frames)
      - total_frames: 所有重试中读取到的总帧数（0 = 真正离线）
    """
    best_frame = None
    best_var = -1
    total_frames = 0
    last_error = None

    for attempt in range(retry_count):
        got_frame, frame, error, frames_seen = capture_frame(rtsp_url, timeout, warmup_seconds)
        total_frames += frames_seen

        if got_frame and error is None:
            return True, frame, attempt, None, total_frames

        if got_frame and error is not None:
            # 最佳帧降级 → 保留最好的非花屏帧，继续重试
            if not _is_corrupted(frame):
                gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
                var = float(gray.var())
                if var > best_var:
                    best_var = var
                    best_frame = frame.copy()
                    last_error = error

        if not got_frame:
            # 完全无数据（真正离线）
            last_error = error

        if attempt < retry_count - 1:
            time.sleep(2)

    if best_frame is not None:
        return True, best_frame, retry_count, last_error, total_frames

    return False, None, retry_count, error or "未知错误", total_frames


# ──────────────────────────────────────────────
# 3. 质量检测
# ──────────────────────────────────────────────

def assess_quality(frame: np.ndarray) -> dict:
    """
    对单帧图像进行质量检测。
    返回质量评分字典。
    """
    if frame is None or frame.size == 0:
        return {"score": 0.0, "reason": "空帧", "isGood": False}

    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY) if len(frame.shape) == 3 else frame

    # 1. Laplacian 方差（清晰度检测）
    lap_var = cv2.Laplacian(gray, cv2.CV_64F).var()
    laplacian_score = min(float(lap_var) / 500.0, 1.0)

    # 2. 亮度检测
    mean_brightness = float(gray.mean())
    if mean_brightness < 15:
        brightness_score = 0.0
    elif mean_brightness > 240:
        brightness_score = 0.0
    else:
        brightness_score = 1.0 - abs(mean_brightness - 128) / 128.0

    # 3. 颜色多样性
    hsv = cv2.cvtColor(frame, cv2.COLOR_BGR2HSV) if len(frame.shape) == 3 else gray
    hist = cv2.calcHist([hsv], [0], None, [180], [0, 180])
    hist_sum = hist.sum()
    if hist_sum > 0:
        dominant_ratio = float(hist.max()) / hist_sum
        color_score = 0.0 if dominant_ratio > 0.9 else 1.0
    else:
        color_score = 1.0

    # 加权综合
    total = laplacian_score * 0.5 + brightness_score * 0.3 + color_score * 0.2

    return {
        "score": round(total, 2),
        "reason": None,
        "laplacianScore": round(laplacian_score, 2),
        "brightnessScore": round(brightness_score, 2),
        "colorDiversityScore": round(color_score, 2),
        "isGood": total >= 0.1,
    }


# ──────────────────────────────────────────────
# 4. 核心入口 — 单摄像头截图+判定
# ──────────────────────────────────────────────

def capture_single(host: str, port: int, app_key: str, app_secret: str,
                   camera_index_code: str, camera_name: str = "",
                   save_dir: str = "./screenshots", timeout: int = 15,
                   retry_count: int = 3, api_path: str = "/artemis",
                   warmup_seconds: float = 5.0) -> dict:
    """
    对单个摄像头执行完整的截图+质量判定流程。

    返回结构化结果:
    {
        "status": "online" | "offline" | "abnormal",
        "screenshotPath": "/path/to/file.jpg",
        "qualityScore": 0.0~1.0,
        "qualityDetail": {...},
        "errorMsg": str | null,
        "captureTime": "2026-06-24 15:01:23",
        "retryUsed": 0,
        "rtspUrl": "rtsp://..." | null,
        "totalFrames": 0~N  // 读取到的总帧数（0=真正离线）
    }
    """
    capture_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    # 1. 获取 RTSP 地址
    rtsp_url, api_error = get_rtsp_url(host, port, app_key, app_secret,
                                        camera_index_code, api_path)
    if api_error:
        return {
            "status": "offline",
            "screenshotPath": None,
            "qualityScore": 0.0,
            "qualityDetail": {},
            "errorMsg": f"API获取RTSP失败: {api_error}",
            "captureTime": capture_time,
            "retryUsed": 0,
            "rtspUrl": None,
            "totalFrames": 0,
        }

    # 2. RTSP 截图（带重试）
    success, frame, retries_used, cap_error, total_frames = capture_with_retry(
        rtsp_url, timeout, retry_count, warmup_seconds)

    if not success:
        return {
            "status": "offline",
            "screenshotPath": None,
            "qualityScore": 0.0,
            "qualityDetail": {},
            "errorMsg": f"RTSP截图失败(重试{retries_used}次): {cap_error}",
            "captureTime": capture_time,
            "retryUsed": retries_used,
            "rtspUrl": rtsp_url,
            "totalFrames": total_frames,
        }

    # 3. 质量检测
    quality = assess_quality(frame)

    # 4. 降级帧 vs 严格帧二次验证
    # 降级帧（capture_with_retry 返回 error!=None）：跳过二次质量过滤，不因评分低而触发额外重试
    # 严格帧（error==None，已通过内部质量门）：若 assess_quality 不通过，做一次二次验证
    if cap_error is not None:
        # 降级帧：已有画面，直接接受，节省二次重试的 ~10s
        pass
    elif not quality["isGood"]:
        # 严格帧质量不佳 → 二次抓取验证
        success2, frame2, _, cap_error2, _ = capture_with_retry(
            rtsp_url, timeout, retry_count, warmup_seconds)
        if success2:
            quality2 = assess_quality(frame2)
            if quality2["isGood"]:
                quality = quality2
                frame = frame2
                retries_used += 1

    # 5. 判定最终状态
    if cap_error is not None:
        # 降级帧：直接在线（已有画面，不因评分低而误判异常）
        status = "online"
    elif not quality["isGood"]:
        status = "abnormal"
    else:
        status = "online"

    # 6. 保存截图 — 按日期分目录，文件名为摄像头名称
    date_dir = datetime.now().strftime("%Y-%m-%d")
    save_path = os.path.join(save_dir, date_dir)
    os.makedirs(save_path, exist_ok=True)
    safe_name = camera_name.replace("/", "_").replace("\\", "_") if camera_name else camera_index_code
    filename = f"{safe_name}.jpg"
    filepath = os.path.join(save_path, filename)
    # Windows 下 cv2.imwrite 不支持 UTF-8 路径，用 imencode + 文件写绕过
    success, encoded = cv2.imencode('.jpg', frame)
    if success:
        with open(filepath, 'wb') as f:
            f.write(encoded.tobytes())
    else:
        return {
            "status": "abnormal",
            "screenshotPath": None,
            "qualityScore": quality["score"],
            "qualityDetail": {
                "laplacianScore": quality.get("laplacianScore", 0),
                "brightnessScore": quality.get("brightnessScore", 0),
                "colorDiversityScore": quality.get("colorDiversityScore", 0),
            },
            "errorMsg": "JPEG编码失败",
            "captureTime": capture_time,
            "retryUsed": retries_used,
            "rtspUrl": rtsp_url,
            "totalFrames": total_frames,
        }

    return {
        "status": status,
        "screenshotPath": filepath,
        "qualityScore": quality["score"],
        "qualityDetail": {
            "laplacianScore": quality.get("laplacianScore", 0),
            "brightnessScore": quality.get("brightnessScore", 0),
            "colorDiversityScore": quality.get("colorDiversityScore", 0),
        },
        "errorMsg": None if status == "online" else quality.get("reason", "画面质量差"),
        "captureTime": capture_time,
        "retryUsed": retries_used,
        "rtspUrl": rtsp_url,
        "totalFrames": total_frames,
    }


# ──────────────────────────────────────────────
# CLI 入口
# ──────────────────────────────────────────────

def main():
    import argparse
    parser = argparse.ArgumentParser(description="海康摄像头截图 + 质量检测")
    parser.add_argument("--host", required=True, help="海康平台地址")
    parser.add_argument("--port", type=int, default=443, help="端口(默认443)")
    parser.add_argument("--app-key", required=True, help="appKey")
    parser.add_argument("--app-secret", required=True, help="appSecret")
    parser.add_argument("--camera-code", required=True, help="摄像头编码")
    parser.add_argument("--camera-name", default="", help="摄像头名称(用于文件名)")
    parser.add_argument("--save-dir", default="./screenshots", help="截图保存目录")
    parser.add_argument("--timeout", type=int, default=15, help="RTSP超时秒数")
    parser.add_argument("--retry", type=int, default=3, help="重试次数")
    parser.add_argument("--api-path", default="/artemis", help="API前缀")
    parser.add_argument("--json", action="store_true", help="输出JSON格式结果")

    args = parser.parse_args()

    result = capture_single(
        host=args.host,
        port=args.port,
        app_key=args.app_key,
        app_secret=args.app_secret,
        camera_index_code=args.camera_code,
        camera_name=args.camera_name,
        save_dir=args.save_dir,
        timeout=args.timeout,
        retry_count=args.retry,
        api_path=args.api_path,
    )

    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print(f"状态: {result['status']}")
        print(f"截图路径: {result['screenshotPath']}")
        print(f"质量评分: {result['qualityScore']}")
        print(f"错误信息: {result['errorMsg']}")
        print(f"捕获时间: {result['captureTime']}")


if __name__ == "__main__":
    main()
