import tkinter as tk
from tkinter import ttk, messagebox
import pygetwindow as gw
import time
import psutil

from PIL import Image, ImageTk

import win32gui
import win32con
import win32api
import win32process
import win32ui


# ========= 이동 로직 =========

def _get_work_area_rect_for_hwnd(hwnd):
    MONITOR_DEFAULTTONEAREST = 2
    hmonitor = win32api.MonitorFromWindow(hwnd, MONITOR_DEFAULTTONEAREST)
    mi = win32api.GetMonitorInfo(hmonitor)
    return mi.get("Work", mi.get("Monitor"))

def bring_window_to_front_by_hwnd(hwnd):
    try:
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(hwnd)
    except Exception:
        pass

def move_window_center_and_signal(hwnd):
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    w, h = right - left, bottom - top

    wk_left, wk_top, wk_right, wk_bottom = _get_work_area_rect_for_hwnd(hwnd)
    wk_w, wk_h = wk_right - wk_left, wk_bottom - wk_top

    x = wk_left + (wk_w - w) // 2
    y = wk_top + (wk_h - h) // 2

    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    bring_window_to_front_by_hwnd(hwnd)

    # (선택) 이동 시작 신호 흉내
    try:
        win32gui.PostMessage(hwnd, win32con.WM_ENTERSIZEMOVE, 0, 0)
    except Exception:
        pass

    flags = win32con.SWP_NOSIZE | win32con.SWP_NOACTIVATE
    win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, x, y, 0, 0, flags)

    time.sleep(0.05)
    try:
        win32gui.PostMessage(hwnd, win32con.WM_EXITSIZEMOVE, 0, 0)
    except Exception:
        pass


# ========= 아이콘 추출 =========

def _get_window_hicon(hwnd):
    # WM_GETICON으로 우선 시도
    for msg_wparam in (2, 0, 1):  # ICON_SMALL2=2, ICON_SMALL=0, ICON_BIG=1
        hicon = win32gui.SendMessage(hwnd, win32con.WM_GETICON, msg_wparam, 0)
        if hicon:
            return hicon
    # 클래스 아이콘으로 폴백
    GCL_HICON = -14
    GCL_HICONSM = -34
    for idx in (GCL_HICONSM, GCL_HICON):
        hicon = win32gui.GetClassLong(hwnd, idx)
        if hicon:
            return hicon
    return None

def _hicon_to_pil_image(hicon, size=(32, 32)):
    """HICON을 PIL Image로 변환"""
    if not hicon:
        return None
    width, height = size
    hdc = win32ui.CreateDCFromHandle(win32gui.GetDC(0))
    memdc = hdc.CreateCompatibleDC()
    bmp = win32ui.CreateBitmap()
    bmp.CreateCompatibleBitmap(hdc, width, height)
    old = memdc.SelectObject(bmp)

    # 아이콘을 DC에 그리기
    win32gui.DrawIconEx(memdc.GetHandleOutput(), 0, 0, hicon, width, height, 0, None, win32con.DI_NORMAL)

    # 비트맵 → bytes
    memdc.SelectObject(old)
    bmp_info = bmp.GetInfo()
    bmp_bytes = bmp.GetBitmapBits(True)

    # BGRX → PIL RGB
    img = Image.frombuffer(
        "RGB",
        (bmp_info["bmWidth"], bmp_info["bmHeight"]),
        bmp_bytes, "raw", "BGRX", 0, 1
    )

    # GDI 리소스 정리
    win32gui.DeleteObject(bmp.GetHandle())
    memdc.DeleteDC()
    hdc.DeleteDC()
    # NOTE: hicon은 시스템/클래스 소유일 수 있으므로 DestroyIcon은 생략

    return img

def get_hwnd_icon_image(hwnd, size=(24, 24)):
    try:
        hicon = _get_window_hicon(hwnd)
        if not hicon:
            return None
        img = _hicon_to_pil_image(hicon, size=(32, 32))
        if img is None:
            return None
        # 안티앨리어싱 리사이즈
        img = img.resize(size, Image.LANCZOS)
        return ImageTk.PhotoImage(img)
    except Exception:
        return None


# ========= 창 목록 수집 =========

def list_windows():
    wins = []
    for w in gw.getAllWindows():
        try:
            if not w.title or not getattr(w, "_hWnd", None):
                continue
            if not win32gui.IsWindowVisible(w._hWnd):
                continue
            wins.append(w)
        except Exception:
            continue
    return wins


# ========= Tk 앱 =========

def _tcl_safe(s: str) -> str:
    """Tk/Tcl 인자 토큰화를 망가뜨릴 수 있는 문자들을 비교적 안전하게 보정."""
    if s is None:
        return ""
    s = str(s)
    # 흔히 문제되는 대괄호/백슬래시/세미콜론 등 기본 이스케이프
    s = s.replace("\\", "\\\\").replace("[", "\\[").replace("]", "\\]").replace(";", "\\;")
    return s


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("창 중앙 이동기 (아이콘/검색 지원)")
        self.geometry("720x520")
        self.minsize(560, 420)

        # ttk 테마 & 스타일
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("Treeview", rowheight=28, font=("Segoe UI", 10))
        style.configure("TButton", font=("Segoe UI", 10))
        style.configure("TEntry", font=("Segoe UI", 10))

        # 상단바
        top = ttk.Frame(self, padding=(10, 10, 10, 0))
        top.pack(fill=tk.X)

        ttk.Label(top, text="검색:", font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT)
        self.search_var = tk.StringVar()
        search = ttk.Entry(top, textvariable=self.search_var, width=30)
        search.pack(side=tk.LEFT, padx=(6, 10))
        search.bind("<KeyRelease>", lambda e: self.refresh_tree())

        self.btn_refresh = ttk.Button(top, text="새로고침", command=self.refresh_tree)
        self.btn_refresh.pack(side=tk.LEFT)

        self.btn_center = ttk.Button(top, text="선택 창 중앙 이동", command=self.center_selected)
        self.btn_center.pack(side=tk.LEFT, padx=(8, 0))

        # 트리뷰 (아이콘 + 제목 + 프로세스 + 클래스)
        mid = ttk.Frame(self, padding=10)
        mid.pack(fill=tk.BOTH, expand=True)

        columns = ("title", "proc", "cls")
        self.tree = ttk.Treeview(mid, columns=columns, show="tree headings")
        self.tree.heading("#0", text="아이콘")
        self.tree.heading("title", text="창 제목")
        self.tree.heading("proc", text="프로세스")
        self.tree.heading("cls", text="클래스")

        self.tree.column("#0", width=36, stretch=False, anchor="center")
        self.tree.column("title", width=380, anchor="w")
        self.tree.column("proc", width=140, anchor="w")
        self.tree.column("cls", width=140, anchor="w")

        vsb = ttk.Scrollbar(mid, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=vsb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.LEFT, fill=tk.Y)

        # 하단 도움말
        bottom = ttk.Frame(self, padding=(10, 0, 10, 10))
        bottom.pack(fill=tk.X)
        ttk.Label(
            bottom,
            text="팁: 더블클릭으로도 중앙 이동 • 아이콘/제목/프로세스로 검색 가능",
            foreground="#666"
        ).pack(side=tk.LEFT)

        self.tree.bind("<Double-1>", lambda e: self.center_selected())

        # 데이터
        self.win_items = []     # [(hwnd, title, proc_name, class_name, tk_img), ...]
        self.tk_images = {}     # hwnd -> PhotoImage (가비지컬렉션 방지)
        self.refresh_tree()

    def refresh_tree(self):
        query = self.search_var.get().strip().lower()
        self.tree.delete(*self.tree.get_children())
        self.win_items.clear()

        for w in list_windows():
            hwnd = w._hWnd
            title = w.title
            class_name = win32gui.GetClassName(hwnd) if win32gui.IsWindow(hwnd) else ""
            proc_name = ""
            try:
                tid, pid = win32process.GetWindowThreadProcessId(hwnd)
                if pid:
                    proc = psutil.Process(pid)
                    proc_name = proc.name()
            except Exception:
                proc_name = ""

            # 검색 필터
            hay = f"{title} {proc_name} {class_name}".lower()
            if query and query not in hay:
                continue

            # 아이콘 생성 (없으면 인자 자체를 생략)
            img = self.tk_images.get(hwnd)
            if img is None:
                img = get_hwnd_icon_image(hwnd, size=(24, 24))
                if img:
                    self.tk_images[hwnd] = img

            # 값들을 안전 문자열로 변환
            vals = (_tcl_safe(title), _tcl_safe(proc_name), _tcl_safe(class_name))

            # insert 옵션 구성: image가 None이면 아예 옵션을 전달하지 않음
            insert_kwargs = {
                "text": "",
                "values": vals,
            }
            if img is not None:
                insert_kwargs["image"] = img

            iid = self.tree.insert("", tk.END, **insert_kwargs)
            # tags에 hwnd 저장
            self.tree.item(iid, tags=(str(hwnd),))
            self.win_items.append((hwnd, title, proc_name, class_name, img))

    def get_selected_hwnd(self):
        sel = self.tree.selection()
        if not sel:
            return None, None
        iid = sel[0]
        title = self.tree.set(iid, "title")
        tags = self.tree.item(iid, "tags")
        hwnd = int(tags[0]) if tags else None
        return hwnd, title

    def center_selected(self):
        hwnd, title = self.get_selected_hwnd()
        if not hwnd:
            messagebox.showwarning("경고", "창을 선택해주세요.")
            return
        try:
            move_window_center_and_signal(hwnd)
            messagebox.showinfo("성공", f"'{title}' 창을 중앙으로 이동했습니다.")
        except Exception as e:
            messagebox.showerror("오류", f"창을 이동할 수 없습니다:\n{e}")


if __name__ == "__main__":
    App().mainloop()
