#!/usr/bin/env python3
"""
쇼핑몰 이미지 검사기 - 화면(GUI) 버전
-------------------------------------
shop_image_check.py 의 검사 기능을 버튼 하나로 쓸 수 있게 감싼 프로그램입니다.
윈도우용 실행 파일(.exe)로 만들면 파이썬이 없는 PC에서도 더블클릭으로 실행됩니다.
"""

import os
import queue
import re
import subprocess
import sys
import threading
import traceback
from datetime import datetime
from pathlib import Path

import tkinter as tk
from tkinter import ttk, messagebox

import shop_image_check as core

APP_TITLE = "쇼핑몰 이미지 검사기"


def default_out_dir():
    """결과 저장 폴더: 문서\\이미지검사 결과 (없으면 홈 폴더)."""
    home = Path.home()
    for name in ("Documents", "문서", "My Documents"):
        if (home / name).is_dir():
            d = home / name / "이미지검사 결과"
            break
    else:
        d = home / "이미지검사 결과"
    d.mkdir(parents=True, exist_ok=True)
    return d


def open_path(path):
    path = str(path)
    if sys.platform.startswith("win"):
        os.startfile(path)  # noqa
    elif sys.platform == "darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])


class _LineWriter:
    """print() 출력을 한 줄씩 콜백으로 넘긴다 (여러 스레드에서 호출돼도 안전)."""

    def __init__(self, on_line):
        self.on_line, self.buf, self.lock = on_line, "", threading.Lock()

    def write(self, s):
        with self.lock:
            self.buf += s
            while "\n" in self.buf:
                line, self.buf = self.buf.split("\n", 1)
                self.on_line(line)
        return len(s)

    def flush(self):
        pass

    def reconfigure(self, **kw):
        pass


def run_check(url, out_dir, on_line, limit=0):
    """검사를 실행하고 만들어진 엑셀 경로를 돌려준다. (화면 없이도 호출 가능)"""
    url = url.strip()
    if not url.startswith("http"):
        url = "https://" + url
    domain = core.urlparse(url).netloc
    name = domain.replace("www.", "").split(".")[0] or "shop"
    out = Path(out_dir) / f"{name}_이미지검사_{datetime.now():%Y%m%d_%H%M}.xlsx"

    argv = ["shop_image_check", url, "--out", str(out)]
    if limit:
        argv += ["--limit", str(limit)]
    old_argv, old_out, old_err = sys.argv, sys.stdout, sys.stderr
    writer = _LineWriter(on_line)
    sys.argv, sys.stdout, sys.stderr = argv, writer, writer
    try:
        core.main()
    except SystemExit as e:
        if e.code not in (0, None):
            raise RuntimeError("상품을 찾지 못했습니다. 쇼핑몰 주소를 확인해 주세요.")
    finally:
        sys.argv, sys.stdout, sys.stderr = old_argv, old_out, old_err
    return out if out.exists() else None


# ---------------------------------------------------------------------------- 화면
class App:
    def __init__(self, root):
        self.root = root
        self.q = queue.Queue()
        self.running = False
        self.result_path = None
        self.out_dir = default_out_dir()

        root.title(APP_TITLE)
        root.geometry("760x560")
        root.minsize(620, 460)
        try:
            root.tk.call("tk", "scaling", 1.3)
        except Exception:
            pass

        pad = {"padx": 16}
        f = ttk.Frame(root)
        f.pack(fill="both", expand=True, pady=12)

        ttk.Label(f, text=APP_TITLE, font=("Malgun Gothic", 16, "bold")).pack(anchor="w", **pad)
        ttk.Label(f, text="카페24 쇼핑몰 주소(또는 상품 1개 링크)를 넣고 [검사 시작]을 누르세요.",
                  foreground="#555").pack(anchor="w", pady=(2, 10), **pad)

        row = ttk.Frame(f)
        row.pack(fill="x", **pad)
        self.url = tk.StringVar()
        ent = ttk.Entry(row, textvariable=self.url, font=("Malgun Gothic", 11))
        ent.pack(side="left", fill="x", expand=True, ipady=4)
        ent.insert(0, "https://")
        ent.focus_set()
        ent.bind("<Return>", lambda e: self.start())
        self.btn = ttk.Button(row, text="검사 시작", command=self.start)
        self.btn.pack(side="left", padx=(8, 0), ipady=3)

        opt = ttk.Frame(f)
        opt.pack(fill="x", pady=(8, 0), **pad)
        self.trial = tk.BooleanVar(value=False)
        ttk.Checkbutton(opt, text="빠른 시험 검사 (상품 30개만)", variable=self.trial).pack(side="left")
        ttk.Label(opt, text=f"결과 저장 위치: {self.out_dir}", foreground="#777").pack(side="right")

        self.phase = tk.StringVar(value="대기 중")
        ttk.Label(f, textvariable=self.phase, font=("Malgun Gothic", 10, "bold")).pack(anchor="w", pady=(14, 4), **pad)
        self.bar = ttk.Progressbar(f, mode="determinate", maximum=100)
        self.bar.pack(fill="x", **pad)

        logf = ttk.Frame(f)
        logf.pack(fill="both", expand=True, pady=(10, 0), **pad)
        self.log = tk.Text(logf, height=12, wrap="none", font=("Consolas", 9), state="disabled",
                           background="#fafafa", relief="solid", borderwidth=1)
        sb = ttk.Scrollbar(logf, command=self.log.yview)
        self.log.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.log.pack(side="left", fill="both", expand=True)
        self.log.tag_configure("bad", foreground="#c62828")
        self.log.tag_configure("done", foreground="#1b5e20", font=("Consolas", 9, "bold"))

        bottom = ttk.Frame(f)
        bottom.pack(fill="x", pady=(10, 0), **pad)
        self.summary = tk.StringVar(value="")
        ttk.Label(bottom, textvariable=self.summary, font=("Malgun Gothic", 11, "bold")).pack(side="left")
        self.open_btn = ttk.Button(bottom, text="결과 엑셀 열기", command=self.open_result, state="disabled")
        self.open_btn.pack(side="right")
        ttk.Button(bottom, text="결과 폴더", command=lambda: open_path(self.out_dir)).pack(side="right", padx=6)

        root.protocol("WM_DELETE_WINDOW", self.on_close)
        root.after(100, self.poll)

    # -- 동작
    def start(self):
        if self.running:
            return
        url = self.url.get().strip()
        if not re.match(r"^(https?://)?[\w.-]+\.[a-z]{2,}", url, re.I) or url in ("https://", "http://"):
            messagebox.showwarning(APP_TITLE, "쇼핑몰 주소를 입력해 주세요.\n예) https://vivacious.kr/")
            return
        self.running, self.result_path = True, None
        self.btn.config(state="disabled")
        self.open_btn.config(state="disabled")
        self.summary.set("")
        self.bar.config(mode="indeterminate")
        self.bar.start(12)
        self.phase.set("상품 목록을 모으는 중...")
        self.log.config(state="normal")
        self.log.delete("1.0", "end")
        self.log.config(state="disabled")
        limit = 30 if self.trial.get() else 0
        threading.Thread(target=self.worker, args=(url, limit), daemon=True).start()

    def worker(self, url, limit):
        try:
            path = run_check(url, self.out_dir, lambda line: self.q.put(("line", line)), limit)
            self.q.put(("done", path))
        except Exception as e:
            self.q.put(("line", traceback.format_exc()))
            self.q.put(("error", str(e)))

    def poll(self):
        try:
            while True:
                kind, val = self.q.get_nowait()
                if kind == "line":
                    self.handle_line(val)
                elif kind == "done":
                    self.finish(val)
                elif kind == "error":
                    self.fail(val)
        except queue.Empty:
            pass
        self.root.after(100, self.poll)

    def handle_line(self, line):
        tag = "bad" if "✕" in line or "[깨짐]" in line else ("done" if line.startswith("완료") else None)
        self.log.config(state="normal")
        self.log.insert("end", line + "\n", tag)
        self.log.see("end")
        self.log.config(state="disabled")

        if line.startswith("[2/3]"):
            self.set_determinate("상품 페이지를 읽는 중")
        elif line.startswith("[3/3]"):
            self.set_determinate("이미지를 하나씩 검사하는 중")
        elif "다른 도메인의 깨진 이미지" in line:
            self.phase.set("교체할 수 있는 정상 파일을 찾는 중...")
        m = re.match(r"^\s+(\d+)/(\d+)\s*$", line)
        if m and self.bar["mode"] == "determinate":
            i, n = int(m.group(1)), int(m.group(2))
            self.bar["value"] = i * 100 / max(n, 1)
            self.phase.set(f"{self.phase_name} ({i:,}/{n:,})")
        m = re.search(r"누적 상품 (\d+)개", line)
        if m:
            self.phase.set(f"상품 목록을 모으는 중... (현재 {int(m.group(1)):,}개)")

    def set_determinate(self, name):
        self.phase_name = name
        self.bar.stop()
        self.bar.config(mode="determinate", value=0)
        self.phase.set(name)

    def finish(self, path):
        self.running = False
        self.bar.stop()
        self.bar.config(mode="determinate")
        self.bar["value"] = 100
        self.btn.config(state="normal")
        self.phase.set("검사 완료")
        text = self.log.get("1.0", "end")
        m = re.search(r"깨진 이미지 (\d+)개 \(상품 (\d+)개\), 주의 (\d+)개", text)
        if m:
            nb, np_, nw = map(int, m.groups())
            self.summary.set("깨진 이미지 없음 ✔" if nb == 0 else f"깨진 이미지 {nb}개 · 문제 상품 {np_}개")
        if path:
            self.result_path = path
            self.open_btn.config(state="normal")
            if messagebox.askyesno(APP_TITLE, "검사가 끝났습니다.\n결과 엑셀 파일을 지금 열까요?"):
                self.open_result()

    def fail(self, msg):
        self.running = False
        self.bar.stop()
        self.bar.config(mode="determinate", value=0)
        self.btn.config(state="normal")
        self.phase.set("검사 실패")
        messagebox.showerror(APP_TITLE, f"검사 중 문제가 생겼습니다.\n\n{msg}\n\n"
                                        "인터넷 연결과 쇼핑몰 주소를 확인한 뒤 다시 시도해 주세요.")

    def open_result(self):
        if self.result_path and Path(self.result_path).exists():
            open_path(self.result_path)

    def on_close(self):
        if self.running and not messagebox.askyesno(APP_TITLE, "검사가 진행 중입니다. 종료할까요?"):
            return
        self.root.destroy()
        os._exit(0)


def main():
    root = tk.Tk()
    try:
        ttk.Style().theme_use("vista" if sys.platform.startswith("win") else "clam")
    except Exception:
        pass
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
