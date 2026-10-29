#!/usr/bin/env python3
"""
쇼핑몰 상품 이미지 깨짐 검사 도구
------------------------------------
지정한 쇼핑몰 페이지를 크롤링하면서 페이지에 올라간 이미지를 모두 수집하고,
각 이미지를 실제로 내려받아 다음을 검사합니다.

  - HTTP 오류 (404, 403, 500 등) / 연결 실패 / 타임아웃
  - 이미지가 아닌 응답 (HTML 오류 페이지 등이 이미지 자리에 오는 경우)
  - 0바이트 파일, Content-Length와 실제 크기 불일치 (전송 중 잘림)
  - 파일 손상: 디코딩 실패, 중간에 잘린(truncated) 이미지
  - 비정상 크기 (가로/세로 1px 등, 선택)

결과는 CSV와 HTML 리포트로 저장되며, 깨진 이미지가 어느 페이지에서 쓰였는지도 함께 표시합니다.

설치:
    pip install requests beautifulsoup4 pillow

사용 예:
    # 메인부터 같은 도메인 페이지를 최대 500개까지 크롤링
    python image_checker.py https://myshop.com --max-pages 500

    # 상품 상세 페이지 URL 패턴만 검사 대상으로 (크롤링은 전체를 따라감)
    python image_checker.py https://myshop.com --page-pattern "/product/"

    # sitemap.xml에 있는 페이지만 검사
    python image_checker.py https://myshop.com --sitemap https://myshop.com/sitemap.xml

    # 이미지 URL 목록 파일만 바로 검사 (크롤링 생략)
    python image_checker.py --image-list urls.txt
"""

import argparse
import csv
import html
import io
import re
import sys
import threading
import time
from collections import defaultdict, deque
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from urllib.parse import urljoin, urlparse, urldefrag

import requests
from bs4 import BeautifulSoup
from PIL import Image, ImageFile

# 잘린 이미지를 "그냥 로드"하지 않도록 명시 (기본값이지만 확실히)
ImageFile.LOAD_TRUNCATED_IMAGES = False
Image.MAX_IMAGE_PIXELS = None  # 큰 상세 이미지 경고 방지

try:  # AVIF 지원 (pillow-avif-plugin 설치 시)
    import pillow_avif  # noqa: F401
except ImportError:
    pass

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/130.0 Safari/537.36 ImageChecker/1.0")

IMG_EXT_RE = re.compile(r"\.(jpe?g|png|gif|webp|avif|bmp|svg|ico|tiff?)(\?|$)", re.I)
BG_URL_RE = re.compile(r"url\(\s*['\"]?([^'\")]+)['\"]?\s*\)", re.I)
LAZY_ATTRS = ("src", "data-src", "data-original", "data-lazy", "data-lazy-src",
              "data-url", "data-img", "data-image", "ec-data-src")
SKIP_PAGE_EXT = re.compile(r"\.(jpe?g|png|gif|webp|avif|bmp|svg|ico|pdf|zip|css|js|mp4|mp3|xml|json)(\?|$)", re.I)

_local = threading.local()


def session():
    s = getattr(_local, "s", None)
    if s is None:
        s = requests.Session()
        s.headers.update({"User-Agent": UA, "Accept-Language": "ko-KR,ko;q=0.9"})
        _local.s = s
    return s


def norm(url):
    url, _ = urldefrag(url.strip())
    return url


# ---------------------------------------------------------------- 이미지 수집
def extract_images(soup, base):
    found = set()

    def add(u):
        if not u:
            return
        u = u.strip()
        if u.startswith("data:") or u.startswith("javascript:") or not u:
            return
        if u.startswith("//"):
            u = urlparse(base).scheme + ":" + u
        found.add(norm(urljoin(base, u)))

    for tag in soup.find_all(["img", "source", "image"]):
        for a in LAZY_ATTRS:
            v = tag.get(a)
            if v and not (a == "src" and tag.name == "source"):
                add(v)
        for a in ("srcset", "data-srcset"):
            v = tag.get(a)
            if v:
                for part in v.split(","):
                    add(part.strip().split(" ")[0])
        if tag.name == "image":
            add(tag.get("href") or tag.get("xlink:href"))

    for m in soup.find_all("meta", attrs={"property": ["og:image", "og:image:url"]}):
        add(m.get("content"))
    for m in soup.find_all("meta", attrs={"name": ["twitter:image"]}):
        add(m.get("content"))

    for tag in soup.find_all(style=True):
        for u in BG_URL_RE.findall(tag["style"]):
            add(u)
    for st in soup.find_all("style"):
        for u in BG_URL_RE.findall(st.get_text() or ""):
            add(u)

    for a in soup.find_all("a", href=True):  # 원본 이미지로 직접 링크된 경우
        if IMG_EXT_RE.search(a["href"]):
            add(a["href"])
    return found


def extract_links(soup, base, domain):
    links = set()
    for a in soup.find_all("a", href=True):
        h = a["href"].strip()
        if h.startswith(("mailto:", "tel:", "javascript:", "#")):
            continue
        u = norm(urljoin(base, h))
        p = urlparse(u)
        if p.scheme in ("http", "https") and p.netloc == domain and not SKIP_PAGE_EXT.search(p.path):
            links.add(u)
    return links


def read_sitemap(url, limit):
    pages, todo, seen = [], deque([url]), set()
    while todo and len(pages) < limit:
        sm = todo.popleft()
        if sm in seen:
            continue
        seen.add(sm)
        try:
            r = session().get(sm, timeout=20)
            r.raise_for_status()
        except Exception as e:
            print(f"  [!] 사이트맵 읽기 실패 {sm}: {e}")
            continue
        locs = re.findall(r"<loc>\s*([^<]+?)\s*</loc>", r.text)
        for loc in locs:
            loc = html.unescape(loc)
            if loc.endswith(".xml") or "sitemap" in loc.lower() and loc.endswith((".xml", ".xml.gz")):
                todo.append(loc)
            else:
                pages.append(loc)
    return pages[:limit]


def crawl(args):
    """페이지를 돌며 {이미지 URL: {등장한 페이지들}} 을 만든다."""
    img_pages = defaultdict(set)
    domain = urlparse(args.url).netloc
    pattern = re.compile(args.page_pattern) if args.page_pattern else None

    if args.sitemap:
        queue = deque(read_sitemap(args.sitemap, args.max_pages))
        follow = False
        print(f"사이트맵에서 페이지 {len(queue)}개를 찾았습니다.")
    else:
        queue = deque([norm(args.url)])
        follow = True

    seen = set(queue)
    visited = 0
    while queue and visited < args.max_pages:
        page = queue.popleft()
        visited += 1
        try:
            r = session().get(page, timeout=args.timeout)
            ctype = r.headers.get("Content-Type", "")
            if r.status_code >= 400 or "html" not in ctype:
                print(f"  [{visited}] 건너뜀 ({r.status_code}) {page}")
                continue
            r.encoding = r.apparent_encoding if not r.encoding or r.encoding.lower() == "iso-8859-1" else r.encoding
            soup = BeautifulSoup(r.text, "html.parser")
        except Exception as e:
            print(f"  [{visited}] 페이지 실패 {page}: {e}")
            continue

        if pattern is None or pattern.search(page):
            imgs = extract_images(soup, r.url)
            for u in imgs:
                img_pages[u].add(page)
            print(f"  [{visited}] 이미지 {len(imgs):>3}개  {page}")
        else:
            print(f"  [{visited}] (패턴 불일치, 링크만 수집) {page}")

        if follow:
            for link in extract_links(soup, r.url, domain):
                if link not in seen:
                    seen.add(link)
                    queue.append(link)
        if args.delay:
            time.sleep(args.delay)

    print(f"\n페이지 {visited}개 방문, 고유 이미지 {len(img_pages)}개 수집\n")
    return img_pages


# ---------------------------------------------------------------- 이미지 검사
def check_image(url, args):
    res = {"url": url, "status": "", "http": "", "content_type": "", "bytes": "",
           "format": "", "width": "", "height": "", "problem": ""}
    data = None
    for attempt in range(args.retries + 1):
        try:
            r = session().get(url, timeout=args.timeout, headers={"Referer": args.url or url})
            res["http"] = r.status_code
            res["content_type"] = r.headers.get("Content-Type", "").split(";")[0]
            data = r.content
            expected = r.headers.get("Content-Length")
            if r.status_code >= 500 and attempt < args.retries:
                time.sleep(1.5)
                continue
            break
        except requests.exceptions.Timeout:
            res["problem"] = "타임아웃"
        except requests.exceptions.ConnectionError as e:
            res["problem"] = f"연결 실패: {str(e)[:80]}"
        except Exception as e:
            res["problem"] = f"요청 오류: {str(e)[:80]}"
        if attempt < args.retries:
            time.sleep(1.5)

    if data is None:
        res["status"] = "BROKEN"
        return res

    res["bytes"] = len(data)
    if res["http"] >= 400:
        res["status"], res["problem"] = "BROKEN", f"HTTP {res['http']}"
        return res
    if len(data) == 0:
        res["status"], res["problem"] = "BROKEN", "0바이트 파일"
        return res
    if expected and expected.isdigit() and "gzip" not in r.headers.get("Content-Encoding", "") \
            and int(expected) != len(data):
        res["status"], res["problem"] = "BROKEN", f"전송 잘림 ({len(data)}/{expected} bytes)"
        return res

    head = data[:512].lstrip().lower()
    is_svg = "svg" in res["content_type"] or url.lower().split("?")[0].endswith(".svg")
    if is_svg:
        if b"<svg" in data[:4096].lower():
            res["status"], res["format"] = "OK", "SVG"
        else:
            res["status"], res["problem"] = "BROKEN", "SVG 내용 없음"
        return res
    if head.startswith((b"<!doctype", b"<html", b"<?xml", b"{")):
        res["status"], res["problem"] = "BROKEN", "이미지 대신 HTML/텍스트 응답"
        return res

    try:
        with Image.open(io.BytesIO(data)) as im:
            im.verify()  # 구조 검사
        with Image.open(io.BytesIO(data)) as im:
            res["format"] = im.format
            res["width"], res["height"] = im.size
            im.load()  # 실제 전체 디코딩 → 잘린 파일 검출
            if getattr(im, "n_frames", 1) > 1:  # 애니메이션 GIF/WebP 마지막 프레임까지
                im.seek(im.n_frames - 1)
                im.load()
    except Image.DecompressionBombError:
        res["status"], res["problem"] = "WARN", "매우 큰 이미지 (디코딩 생략)"
        return res
    except (OSError, SyntaxError, ValueError) as e:
        msg = str(e)
        if "cannot identify image file" in msg:
            if res["content_type"] in ("image/avif", "image/heic", "image/heif") or \
                    re.search(r"\.(avif|heic)(\?|$)", url, re.I):
                res["status"], res["problem"] = "WARN", "AVIF/HEIC (검사 미지원: pip install pillow-avif-plugin)"
                return res
            res["status"], res["problem"] = "BROKEN", "이미지로 인식 불가 (손상 또는 잘못된 형식)"
        elif "truncated" in msg.lower():
            res["status"], res["problem"] = "BROKEN", "파일이 중간에 잘림 (truncated)"
        else:
            res["status"], res["problem"] = "BROKEN", f"디코딩 실패: {msg[:80]}"
        return res

    if args.min_size and (res["width"] < args.min_size or res["height"] < args.min_size):
        res["status"], res["problem"] = "WARN", f"비정상적으로 작음 ({res['width']}x{res['height']})"
        return res
    if not res["content_type"].startswith("image/"):
        res["status"], res["problem"] = "WARN", f"Content-Type 이상: {res['content_type'] or '없음'}"
        return res
    res["status"] = "OK"
    return res


# ---------------------------------------------------------------- 리포트
def write_reports(results, img_pages, out_prefix, started):
    order = {"BROKEN": 0, "WARN": 1, "OK": 2}
    results.sort(key=lambda r: (order.get(r["status"], 3), r["url"]))
    csv_path, html_path = f"{out_prefix}.csv", f"{out_prefix}.html"

    with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["상태", "문제", "이미지 URL", "HTTP", "Content-Type", "크기(bytes)",
                    "포맷", "가로", "세로", "사용된 페이지 수", "사용된 페이지(최대 5개)"])
        for r in results:
            pages = sorted(img_pages.get(r["url"], []))
            w.writerow([r["status"], r["problem"], r["url"], r["http"], r["content_type"], r["bytes"],
                        r["format"], r["width"], r["height"], len(pages), " | ".join(pages[:5])])

    counts = {k: sum(1 for r in results if r["status"] == k) for k in ("BROKEN", "WARN", "OK")}
    rows = []
    for r in results:
        if r["status"] == "OK":
            continue
        pages = sorted(img_pages.get(r["url"], []))
        plinks = "<br>".join(f'<a href="{html.escape(p)}" target="_blank">{html.escape(p)}</a>' for p in pages[:5])
        if len(pages) > 5:
            plinks += f"<br><small>외 {len(pages) - 5}개</small>"
        u = html.escape(r["url"])
        rows.append(f"""<tr class="{r['status'].lower()}">
<td><span class="badge">{r['status']}</span></td><td>{html.escape(r['problem'])}</td>
<td><img src="{u}" loading="lazy" alt="" onerror="this.replaceWith('✕')"></td>
<td class="url"><a href="{u}" target="_blank">{u}</a></td><td>{r['http']}</td><td class="url">{plinks}</td></tr>""")

    body = "\n".join(rows) or '<tr><td colspan="6" style="text-align:center;padding:2rem">깨진 이미지가 없습니다 🎉</td></tr>'
    elapsed = (datetime.now() - started).seconds
    doc = f"""<!doctype html><html lang="ko"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><title>이미지 검사 리포트</title>
<style>
body{{font-family:-apple-system,"Malgun Gothic",sans-serif;margin:24px;color:#222;background:#fafafa}}
h1{{font-size:20px}} .sum{{display:flex;gap:12px;margin:16px 0}}
.card{{background:#fff;border:1px solid #ddd;border-radius:8px;padding:12px 18px}} .card b{{font-size:22px;display:block}}
table{{border-collapse:collapse;width:100%;background:#fff;font-size:13px}}
td,th{{border-bottom:1px solid #eee;padding:8px;text-align:left;vertical-align:top}}
th{{background:#f0f0f0;position:sticky;top:0}} .url{{word-break:break-all;max-width:420px}}
img{{max-width:80px;max-height:80px}} .badge{{padding:2px 8px;border-radius:4px;color:#fff;font-weight:600}}
.broken .badge{{background:#d93025}} .warn .badge{{background:#e8a000}}
</style></head><body>
<h1>상품 이미지 깨짐 검사 리포트</h1>
<div>검사 시각: {started:%Y-%m-%d %H:%M} · 소요 {elapsed}초</div>
<div class="sum"><div class="card">전체<b>{len(results)}</b></div>
<div class="card" style="color:#d93025">깨짐<b>{counts['BROKEN']}</b></div>
<div class="card" style="color:#b07800">주의<b>{counts['WARN']}</b></div>
<div class="card" style="color:#188038">정상<b>{counts['OK']}</b></div></div>
<table><tr><th>상태</th><th>문제</th><th>미리보기</th><th>이미지 URL</th><th>HTTP</th><th>사용된 페이지</th></tr>
{body}</table><p><small>정상 이미지를 포함한 전체 목록은 {html.escape(csv_path)} 에 있습니다.</small></p></body></html>"""
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(doc)
    return csv_path, html_path, counts


# ---------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser(description="쇼핑몰 상품 이미지 깨짐 검사")
    ap.add_argument("url", nargs="?", help="시작 URL (쇼핑몰 메인 또는 카테고리 페이지)")
    ap.add_argument("--max-pages", type=int, default=300, help="방문할 최대 페이지 수 (기본 300)")
    ap.add_argument("--page-pattern", help="이미지를 수집할 페이지 URL 정규식 (예: '/product/')")
    ap.add_argument("--sitemap", help="sitemap.xml URL (지정 시 사이트맵 페이지만 검사)")
    ap.add_argument("--image-list", help="이미지 URL 목록 텍스트 파일 (한 줄에 하나, 크롤링 생략)")
    ap.add_argument("--image-domain", action="append",
                    help="이 도메인의 이미지만 검사 (여러 번 지정 가능, 예: --image-domain img.myshop.com)")
    ap.add_argument("--workers", type=int, default=8, help="동시 이미지 검사 수 (기본 8)")
    ap.add_argument("--delay", type=float, default=0.3, help="페이지 요청 간 대기 초 (기본 0.3)")
    ap.add_argument("--timeout", type=float, default=20)
    ap.add_argument("--retries", type=int, default=1, help="실패 시 재시도 횟수")
    ap.add_argument("--min-size", type=int, default=2, help="이보다 작은 가로/세로는 주의 표시 (0=끄기)")
    ap.add_argument("--out", default=f"image_report_{datetime.now():%Y%m%d_%H%M}", help="리포트 파일 이름(확장자 제외)")
    args = ap.parse_args()

    if not args.url and not args.image_list:
        ap.error("시작 URL 또는 --image-list 중 하나는 필요합니다.")

    started = datetime.now()
    if args.image_list:
        with open(args.image_list, encoding="utf-8") as f:
            img_pages = {norm(l): set() for l in f if l.strip() and not l.startswith("#")}
        print(f"목록에서 이미지 {len(img_pages)}개를 읽었습니다.\n")
    else:
        print(f"크롤링 시작: {args.url}\n")
        img_pages = crawl(args)

    urls = list(img_pages)
    if args.image_domain:
        doms = set(args.image_domain)
        urls = [u for u in urls if urlparse(u).netloc in doms]
        print(f"이미지 도메인 필터 적용 → {len(urls)}개")

    print(f"이미지 {len(urls)}개 검사 중 (동시 {args.workers}개)...")
    results, done = [], 0
    with ThreadPoolExecutor(max_workers=args.workers) as ex:
        futs = {ex.submit(check_image, u, args): u for u in urls}
        for fut in as_completed(futs):
            r = fut.result()
            results.append(r)
            done += 1
            if r["status"] != "OK":
                print(f"  ✕ [{r['status']}] {r['problem']}  {r['url']}")
            elif done % 50 == 0:
                print(f"  ... {done}/{len(urls)}")

    csv_path, html_path, c = write_reports(results, img_pages, args.out, started)
    print(f"\n완료: 전체 {len(results)} / 깨짐 {c['BROKEN']} / 주의 {c['WARN']} / 정상 {c['OK']}")
    print(f"리포트: {html_path}\n전체 목록: {csv_path}")
    sys.exit(1 if c["BROKEN"] else 0)


if __name__ == "__main__":
    main()
