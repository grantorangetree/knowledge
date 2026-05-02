#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
╔══════════════════════════════════════════════════════╗
║   CZUR 스캔 이미지 → 문서 변환기                       ║
║   Gemma 4 E4B (Ollama) 비전 모델 사용                 ║
╚══════════════════════════════════════════════════════╝

사용법:
  # 전체 JPG 일괄 처리
  python3 scan_to_doc.py

  # 특정 이미지만 처리
  python3 scan_to_doc.py --image IMG_2023_12_04_23_40_24_026.jpg

  # 출력 파일 지정
  python3 scan_to_doc.py --out output.md

  # 모델 지정
  python3 scan_to_doc.py --model gemma4:e4b
"""

import argparse, base64, json, sys, os, time
from pathlib import Path
from urllib import request, error

# ── 설정 ──────────────────────────────────────────────
OLLAMA_URL  = "http://localhost:11434/api/generate"
DEFAULT_MODEL = "gemma4:e4b"
BASE_DIR    = Path(__file__).parent
OUTPUT_DIR  = BASE_DIR / "extracted_docs"

# ── ANSI 컬러 ─────────────────────────────────────────
RESET  = "\033[0m";  BOLD   = "\033[1m"
CYAN   = "\033[96m"; GREEN  = "\033[92m"
YELLOW = "\033[93m"; RED    = "\033[91m"
DIM    = "\033[2m";  PURPLE = "\033[95m"

def ok(m):   print(f"{GREEN}  ✓  {m}{RESET}")
def warn(m): print(f"{YELLOW}  ⚠  {m}{RESET}")
def err(m):  print(f"{RED}  ✗  {m}{RESET}")
def info(m): print(f"{CYAN}  →  {m}{RESET}")
def step(m): print(f"\n{BOLD}{CYAN}▸ {m}{RESET}")

PROMPT = """이 이미지는 스캔된 문서입니다. 다음을 수행해주세요:

1. 이미지에 보이는 모든 텍스트를 정확하게 추출하세요
2. 문서의 구조(제목, 소제목, 단락, 표, 목록 등)를 최대한 유지하세요
3. 한국어와 영어가 섞여 있다면 그대로 보존하세요
4. 읽기 어려운 부분은 [불명확] 으로 표시하세요
5. 이미지/그림에 대한 설명은 [이미지: 설명] 형식으로 작성하세요

추출된 내용만 출력하고 다른 설명은 하지 마세요."""


def image_to_base64(path: Path) -> str:
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode("utf-8")


def analyze_image(image_path: Path, model: str) -> str:
    """Ollama API로 이미지 분석"""
    b64 = image_to_base64(image_path)

    payload = {
        "model": model,
        "prompt": PROMPT,
        "images": [b64],
        "stream": False,
        "options": {
            "num_ctx": 8192,
            "temperature": 0.1,
        }
    }

    data = json.dumps(payload).encode("utf-8")
    req = request.Request(
        OLLAMA_URL,
        data=data,
        headers={"Content-Type": "application/json"},
        method="POST"
    )

    try:
        with request.urlopen(req, timeout=120) as resp:
            result = json.loads(resp.read().decode("utf-8"))
            return result.get("response", "").strip()
    except error.URLError as e:
        raise RuntimeError(f"Ollama 연결 실패: {e}")


def process_single(image_path: Path, model: str, out_path: Path = None) -> Path:
    """단일 이미지 처리"""
    step(f"분석 중: {image_path.name}")
    info(f"모델: {model}")

    start = time.time()
    text = analyze_image(image_path, model)
    elapsed = time.time() - start

    if not text:
        warn("추출된 텍스트 없음")
        return None

    # 출력 경로 결정
    if out_path is None:
        OUTPUT_DIR.mkdir(exist_ok=True)
        out_path = OUTPUT_DIR / (image_path.stem + ".md")

    # 마크다운 헤더 포함해서 저장
    content = f"# {image_path.name}\n\n"
    content += f"> 원본: `{image_path}`  \n"
    content += f"> 분석 모델: `{model}`  \n"
    content += f"> 처리 시간: {elapsed:.1f}초\n\n"
    content += "---\n\n"
    content += text + "\n"

    out_path.write_text(content, encoding="utf-8")
    ok(f"저장 완료: {out_path.name}  ({elapsed:.1f}초)")
    return out_path


def process_batch(image_paths: list, model: str, combined_output: Path = None):
    """여러 이미지 일괄 처리"""
    total = len(image_paths)
    failed = []
    results = []

    combined_lines = []

    for i, img_path in enumerate(image_paths, 1):
        print(f"\n{PURPLE}{BOLD}[{i}/{total}]{RESET} {img_path.name}")
        try:
            out = process_single(img_path, model)
            if out:
                results.append(out)
                # 통합 문서용 내용 추가
                text = out.read_text(encoding="utf-8")
                combined_lines.append(text)
        except Exception as e:
            err(f"실패: {e}")
            failed.append(img_path.name)

        # API 과부하 방지
        if i < total:
            time.sleep(1)

    # 통합 문서 저장
    if combined_output and combined_lines:
        combined_output.write_text("\n\n---\n\n".join(combined_lines), encoding="utf-8")
        ok(f"\n통합 문서 저장: {combined_output}")

    # 결과 요약
    print(f"\n{BOLD}{'═'*50}{RESET}")
    ok(f"완료: {len(results)}/{total}개")
    if failed:
        warn(f"실패 {len(failed)}개: {', '.join(failed[:5])}")


def main():
    print(f"""
{PURPLE}{BOLD}╔══════════════════════════════════════════════════════╗
║   CZUR 스캔 → 문서 변환기  (Gemma 4 Vision)          ║
╚══════════════════════════════════════════════════════╝{RESET}
""")

    parser = argparse.ArgumentParser(description="스캔 이미지를 AI로 분석해 문서로 변환")
    parser.add_argument("--image",  "-i", metavar="PATH", help="특정 이미지 파일")
    parser.add_argument("--dir",    "-d", metavar="PATH", help="이미지 폴더 (기본: 현재 폴더)")
    parser.add_argument("--out",    "-o", metavar="PATH", help="출력 파일 경로")
    parser.add_argument("--model",  "-m", metavar="MODEL", default=DEFAULT_MODEL, help=f"Ollama 모델 (기본: {DEFAULT_MODEL})")
    parser.add_argument("--combine","-c", action="store_true", help="모든 결과를 하나의 파일로 합치기")
    parser.add_argument("--limit",  "-l", type=int, metavar="N", help="처리할 최대 이미지 수")
    args = parser.parse_args()

    # Ollama 연결 확인
    step("Ollama 연결 확인")
    try:
        req = request.Request("http://localhost:11434/api/tags")
        with request.urlopen(req, timeout=5) as resp:
            ok("Ollama 연결 정상")
    except Exception:
        err("Ollama가 실행 중이 아닙니다. 'ollama serve' 를 먼저 실행하세요.")
        sys.exit(1)

    # ── 단일 이미지 모드 ───────────────────────────────
    if args.image:
        img_path = Path(args.image)
        if not img_path.is_absolute():
            img_path = BASE_DIR / img_path
        if not img_path.exists():
            err(f"파일 없음: {img_path}")
            sys.exit(1)
        out_path = Path(args.out) if args.out else None
        process_single(img_path, args.model, out_path)
        return

    # ── 일괄 처리 모드 ────────────────────────────────
    search_dir = Path(args.dir) if args.dir else BASE_DIR
    images = sorted([
        p for p in search_dir.glob("*.jpg")
        if not p.name.startswith(".")
    ])
    images += sorted([
        p for p in search_dir.glob("*.JPG")
        if not p.name.startswith(".")
    ])

    if not images:
        err(f"JPG 이미지를 찾지 못했습니다: {search_dir}")
        sys.exit(1)

    if args.limit:
        images = images[:args.limit]

    info(f"처리 대상: {len(images)}개 이미지")
    info(f"출력 폴더: {OUTPUT_DIR}")
    info(f"사용 모델: {args.model}")

    print(f"\n  계속하려면 Enter, 취소하려면 Ctrl+C ...")
    try:
        input()
    except KeyboardInterrupt:
        print(f"\n{YELLOW}취소됨{RESET}\n")
        sys.exit(0)

    combined_out = None
    if args.combine or args.out:
        combined_out = Path(args.out) if args.out else (BASE_DIR / "전체_문서.md")

    process_batch(images, args.model, combined_out)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print(f"\n{YELLOW}  취소되었습니다.{RESET}\n")
        sys.exit(0)
