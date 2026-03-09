# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, messagebox
import time
import psutil
import ctypes
from ctypes import wintypes

import pygetwindow as gw
from PIL import Image, ImageTk

import win32gui
import win32con
import win32api
import win32process
import win32ui

# ========= DPI 인식 (고해상도에서 흐림 방지) =========
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)  # PROCESS_SYSTEM_DPI_AWARE
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

# ========= DWM 확장 프레임 보정 =========
DWMWA_EXTENDED_FRAME_BOUNDS = 9

def get_extended_frame_bounds(hwnd):
    """DWM이 보고하는 '시각적 프레임'(그림자 제외) 사각형을 반환. 실패 시 GetWindowRect."""
    rect = wintypes.RECT()
    try:
        hr = ctypes.windll.dwmapi.DwmGetWindowAttribute(
            wintypes.HWND(hwnd),
            ctypes.c_uint(DWMWA_EXTENDED_FRAME_BOUNDS),
            ctypes.byref(rect),
            ctypes.sizeof(rect),
        )
        if hr == 0:
            return rect.left, rect.top, rect.right, rect.bottom
    except Exception:
        pass
    l, t, r, b = win32gui.GetWindowRect(hwnd)
    return l, t, r, b

def get_frame_padding(hwnd):
    """
    바깥(그림자 포함) 사각형과 '시각적 프레임' 차이를 계산.
    반환: (pad_left, pad_top, pad_right, pad_bottom, outer_w, outer_h, frame_w, frame_h)
    """
    ol, ot, or_, ob = win32gui.GetWindowRect(hwnd)              # Outer (shadow 포함)
    fl, ft, fr, fb = get_extended_frame_bounds(hwnd)            # Frame (shadow 제외)
    pad_left   = fl - ol
    pad_top    = ft - ot
    pad_right  = or_ - fr
    pad_bottom = ob - fb
    OW = or_ - ol
    OH = ob - ot
    FW = fr - fl
    FH = fb - ft
    return (pad_left, pad_top, pad_right, pad_bottom, OW, OH, FW, FH)

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

    try:
        win32gui.PostMessage(hwnd, win32con.WM_ENTERSIZEMOVE, 0, 0)
    except Exception:
        pass

    win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, x, y, 0, 0,
                          win32con.SWP_NOSIZE | win32con.SWP_NOACTIVATE)

    time.sleep(0.05)
    try:
        win32gui.PostMessage(hwnd, win32con.WM_EXITSIZEMOVE, 0, 0)
    except Exception:
        pass

def get_window_size(hwnd):
    """창의 크기(너비, 높이)를 반환합니다."""
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    return (right - left, bottom - top)

def get_window_position(hwnd):
    """창의 위치(x, y)를 반환합니다. DWM 시각적 프레임 기준."""
    fl, ft, _, _ = get_extended_frame_bounds(hwnd)
    return (fl, ft)

def apply_window_position(hwnd, target_x, target_y):
    """
    창을 지정된 위치로 이동합니다. 크기는 유지합니다.
    target_x, target_y는 DWM 시각적 프레임 기준 좌표이므로
    실제 SetWindowPos에는 그림자 패딩만큼 보정하여 전달합니다.
    """
    pad_left, pad_top, _, _, _, _, _, _ = get_frame_padding(hwnd)
    outer_x = target_x - pad_left
    outer_y = target_y - pad_top

    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    try:
        win32gui.PostMessage(hwnd, win32con.WM_ENTERSIZEMOVE, 0, 0)
    except Exception:
        pass

    win32gui.SetWindowPos(
        hwnd, win32con.HWND_TOP,
        int(outer_x), int(outer_y), 0, 0,
        win32con.SWP_NOSIZE | win32con.SWP_NOACTIVATE
    )

    time.sleep(0.05)
    try:
        win32gui.PostMessage(hwnd, win32con.WM_EXITSIZEMOVE, 0, 0)
    except Exception:
        pass

def apply_window_size(hwnd, width, height):
    """
    창에 지정된 크기를 적용합니다.
    현재 위치는 유지하고 크기만 변경합니다.
    """
    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    
    # WM_ENTERSIZEMOVE / WM_EXITSIZEMOVE 를 보내면 일부 앱이 크기 변경을 더 잘 인식함
    try:
        win32gui.PostMessage(hwnd, win32con.WM_ENTERSIZEMOVE, 0, 0)
    except Exception:
        pass
    
    # SWP_NOMOVE: 위치는 변경하지 않음, 크기만 변경
    win32gui.SetWindowPos(
        hwnd, win32con.HWND_TOP,
        0, 0,  # x, y는 SWP_NOMOVE로 인해 무시됨
        int(width), int(height),
        win32con.SWP_NOMOVE | win32con.SWP_NOACTIVATE
    )
    
    time.sleep(0.05)
    try:
        win32gui.PostMessage(hwnd, win32con.WM_EXITSIZEMOVE, 0, 0)
    except Exception:
        pass

def move_window_to_corner(hwnd, corner="top-left", margin=0):
    """
    모니터 '전체 좌표'(작업표시줄 포함) 기준으로,
    시각적 프레임(그림자 제외)이 모서리에 '딱' 닿도록 배치.
    corner: 'top-left' | 'bottom-left' | 'top-right' | 'bottom-right'
    margin: 가장자리 여백(px)
    """
    # 프레임 패딩/크기
    pad_left, pad_top, pad_right, pad_bottom, OW, OH, FW, FH = get_frame_padding(hwnd)

    # 창이 걸친 모니터 좌표
    MONITOR_DEFAULTTONEAREST = 2
    hmon = win32api.MonitorFromWindow(hwnd, MONITOR_DEFAULTTONEAREST)
    mi = win32api.GetMonitorInfo(hmon)
    m_left, m_top, m_right, m_bottom = mi["Monitor"]

    # 목표: frame.* 가 모니터 경계(m_* ± margin)에 정확히 닿도록
    if corner == "top-left":
        x = m_left + margin - pad_left                       # frame.left = x + pad_left
        y = m_top  + margin - pad_top                        # frame.top  = y + pad_top
    elif corner == "bottom-left":
        x = m_left + margin - pad_left
        y = m_bottom - margin - (OH - pad_bottom)            # frame.bottom = y + OH - pad_bottom
    elif corner == "top-right":
        x = m_right - margin - (OW - pad_right)              # frame.right  = x + OW - pad_right
        y = m_top   + margin - pad_top
    elif corner == "bottom-right":
        x = m_right - margin - (OW - pad_right)
        y = m_bottom - margin - (OH - pad_bottom)
    else:
        raise ValueError("corner must be one of: top-left, bottom-left, top-right, bottom-right")

    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    win32gui.SetWindowPos(
        hwnd, win32con.HWND_TOP,
        int(x), int(y), 0, 0,
        win32con.SWP_NOSIZE | win32con.SWP_NOACTIVATE
    )
    try:
        win32gui.PostMessage(hwnd, win32con.WM_EXITSIZEMOVE, 0, 0)
    except Exception:
        pass

def move_window_to_edge(hwnd, direction="top", margin=0):
    """
    창을 한 축만 이동시켜 모니터 가장자리에 배치합니다.
    direction: 'top' | 'bottom' | 'left' | 'right'
      - 'top': X축 유지, Y축을 화면 맨 위로
      - 'bottom': X축 유지, Y축을 화면 맨 아래로
      - 'left': Y축 유지, X축을 화면 맨 왼쪽으로
      - 'right': Y축 유지, X축을 화면 맨 오른쪽으로
    margin: 가장자리 여백(px)
    """
    # 현재 창의 바깥 사각형 (그림자 포함)
    ol, ot, or_, ob = win32gui.GetWindowRect(hwnd)
    current_x, current_y = ol, ot
    OW = or_ - ol
    OH = ob - ot
    
    # 프레임 패딩 (그림자 보정용)
    pad_left, pad_top, pad_right, pad_bottom, _, _, FW, FH = get_frame_padding(hwnd)
    
    # 창이 속한 모니터의 전체 좌표
    MONITOR_DEFAULTTONEAREST = 2
    hmon = win32api.MonitorFromWindow(hwnd, MONITOR_DEFAULTTONEAREST)
    mi = win32api.GetMonitorInfo(hmon)
    m_left, m_top, m_right, m_bottom = mi["Monitor"]
    
    # 방향에 따라 한 축만 변경
    if direction == "top":
        # X축 유지, Y를 화면 맨 위로 (frame.top이 m_top + margin에 닿도록)
        x = current_x
        y = m_top + margin - pad_top
    elif direction == "bottom":
        # X축 유지, Y를 화면 맨 아래로 (frame.bottom이 m_bottom - margin에 닿도록)
        x = current_x
        y = m_bottom - margin - (OH - pad_bottom)
    elif direction == "left":
        # Y축 유지, X를 화면 맨 왼쪽으로 (frame.left가 m_left + margin에 닿도록)
        x = m_left + margin - pad_left
        y = current_y
    elif direction == "right":
        # Y축 유지, X를 화면 맨 오른쪽으로 (frame.right가 m_right - margin에 닿도록)
        x = m_right - margin - (OW - pad_right)
        y = current_y
    else:
        raise ValueError("direction must be one of: top, bottom, left, right")
    
    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    win32gui.SetWindowPos(
        hwnd, win32con.HWND_TOP,
        int(x), int(y), 0, 0,
        win32con.SWP_NOSIZE | win32con.SWP_NOACTIVATE
    )
    try:
        win32gui.PostMessage(hwnd, win32con.WM_EXITSIZEMOVE, 0, 0)
    except Exception:
        pass

# ========= 아이콘 추출 =========
def _get_window_hicon(hwnd):
    for msg_wparam in (2, 0, 1):  # ICON_SMALL2, ICON_SMALL, ICON_BIG
        hicon = win32gui.SendMessage(hwnd, win32con.WM_GETICON, msg_wparam, 0)
        if hicon:
            return hicon
    GCL_HICON = -14
    GCL_HICONSM = -34
    for idx in (GCL_HICONSM, GCL_HICON):
        hicon = win32gui.GetClassLong(hwnd, idx)
        if hicon:
            return hicon
    return None

def _hicon_to_pil_image(hicon, size=(32, 32)):
    if not hicon:
        return None
    width, height = size
    hdc = win32ui.CreateDCFromHandle(win32gui.GetDC(0))
    memdc = hdc.CreateCompatibleDC()
    bmp = win32ui.CreateBitmap()
    bmp.CreateCompatibleBitmap(hdc, width, height)
    old = memdc.SelectObject(bmp)
    win32gui.DrawIconEx(memdc.GetHandleOutput(), 0, 0, hicon, width, height, 0, None, win32con.DI_NORMAL)
    memdc.SelectObject(old)

    bmp_info = bmp.GetInfo()
    bmp_bytes = bmp.GetBitmapBits(True)
    img = Image.frombuffer("RGB", (bmp_info["bmWidth"], bmp_info["bmHeight"]),
                           bmp_bytes, "raw", "BGRX", 0, 1)

    win32gui.DeleteObject(bmp.GetHandle())
    memdc.DeleteDC()
    hdc.DeleteDC()
    return img

def get_hwnd_icon_image(hwnd, size=(20, 20)):
    try:
        hicon = _get_window_hicon(hwnd)
        if not hicon:
            return None
        img = _hicon_to_pil_image(hicon, size=(32, 32))
        if img is None:
            return None
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

# ========= 유틸 =========
def _tcl_safe(s: str) -> str:
    if s is None:
        return ""
    s = str(s)
    return s.replace("\\", "\\\\").replace("[", "\\[").replace("]", "\\]").replace(";", "\\;")

# ========= 메인 앱 =========
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("창 중앙 이동기  •  Light ✦ Clean")
        self.geometry("1200x620")          # ✅ 가로 1200
        self.minsize(800, 520)

        self._build_style_light()
        self._build_ui()

        self.tk_images = {}  # hwnd -> PhotoImage
        self.saved_size = None  # (width, height) - 기억된 창 크기
        self.saved_size_title = None  # 크기를 기억한 창의 제목 (UI 표시용)
        self.saved_position = None  # (x, y) - 기억된 창 위치
        self.saved_position_title = None  # 위치를 기억한 창의 제목 (UI 표시용)
        self.refresh_tree()

        # 단축키
        self.bind("<Return>", lambda e: self.center_selected())
        self.bind("<F5>", lambda e: self.refresh_tree())
        self.bind("<Delete>", lambda e: self.close_selected())
        self.bind("<Control-l>", lambda e: (self.search_entry.focus_set(), self.search_entry.select_range(0, "end")))
        # 모서리 이동 단축키 (모니터 좌표 기준, margin=0)
        self.bind("<Alt-1>", lambda e: self.move_selected_top_left())
        self.bind("<Alt-2>", lambda e: self.move_selected_bottom_left())
        self.bind("<Alt-3>", lambda e: self.move_selected_top_right())
        self.bind("<Alt-4>", lambda e: self.move_selected_bottom_right())
        # 크기 복사 단축키
        self.bind("<Control-Shift-C>", lambda e: self.remember_window_size())
        self.bind("<Control-Shift-V>", lambda e: self.apply_remembered_size())
        # 위치 복사 단축키
        self.bind("<Control-Alt-c>", lambda e: self.remember_window_position())
        self.bind("<Control-Alt-v>", lambda e: self.apply_remembered_position())
        # 가장자리 이동 단축키 (한 축만 이동)
        self.bind("<Alt-Up>", lambda e: self.move_selected_to_top())
        self.bind("<Alt-Down>", lambda e: self.move_selected_to_bottom())
        self.bind("<Alt-Left>", lambda e: self.move_selected_to_left())
        self.bind("<Alt-Right>", lambda e: self.move_selected_to_right())

    # ----- 라이트 테마 -----
    def _build_style_light(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass

        bg      = "#F7F9FC"  # window
        panel   = "#FFFFFF"  # cards
        txt     = "#1B2430"
        subtxt  = "#6B7280"
        selbg   = "#E6EFFB"
        header  = "#EEF2F7"
        row1    = "#FFFFFF"
        row2    = "#F4F6FA"

        self.configure(bg=bg)

        style.configure(".", background=bg, foreground=txt)
        style.configure("Light.TFrame", background=panel)
        style.configure("Naked.TFrame", background=bg)
        style.configure("Light.TLabel", background=panel, foreground=txt)
        style.configure("Hint.TLabel", background=bg, foreground=subtxt)

        style.configure("TButton", padding=6)
        style.map("TButton", background=[("active", "#F1F5F9")])
        style.configure("TEntry", fieldbackground="#FFFFFF", foreground=txt)

        style.configure("Treeview",
                        background=row1, fieldbackground=row1, foreground=txt, rowheight=28,
                        borderwidth=0)
        style.configure("Treeview.Heading",
                        background=header, foreground=txt, relief="flat")
        style.map("Treeview",
                  background=[("selected", selbg)],
                  foreground=[("selected", txt)])
        style.layout("Treeview", [("Treeview.treearea", {"sticky": "nswe"})])

        style.configure("odd.Treeview", background=row1)
        style.configure("even.Treeview", background=row2)

    # ----- UI 구성 -----
    def _build_ui(self):
        # 상단 바
        top_wrap = ttk.Frame(self, style="Naked.TFrame", padding=(12, 12, 12, 6))
        top_wrap.pack(fill=tk.X)

        top = ttk.Frame(top_wrap, style="Light.TFrame", padding=(12, 10))
        top.pack(fill=tk.X)

        title = ttk.Label(top, text="열려있는 창", style="Light.TLabel",
                          font=("Segoe UI Semibold", 12))
        title.pack(side=tk.LEFT)

        ttk.Label(top, text="  ", style="Light.TLabel").pack(side=tk.LEFT)

        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(top, textvariable=self.search_var, width=40)
        self.search_entry.pack(side=tk.LEFT, padx=(6, 8))
        self.search_entry.bind("<KeyRelease>", lambda e: self.refresh_tree())

        self.btn_center = ttk.Button(top, text="선택 창 중앙 이동", command=self.center_selected)
        self.btn_center.pack(side=tk.LEFT, padx=(0, 6))
        self.btn_front = ttk.Button(top, text="전면으로", command=self.bring_to_front_selected)
        self.btn_front.pack(side=tk.LEFT, padx=(0, 6))
        self.btn_refresh = ttk.Button(top, text="새로고침 (F5)", command=self.refresh_tree)
        self.btn_refresh.pack(side=tk.LEFT)

        # 중간: Treeview (아이콘 칼럼 포함)
        mid_wrap = ttk.Frame(self, style="Naked.TFrame", padding=(12, 6, 12, 6))
        mid_wrap.pack(fill=tk.BOTH, expand=True)

        mid = ttk.Frame(mid_wrap, style="Light.TFrame", padding=(8, 8))
        mid.pack(fill=tk.BOTH, expand=True)

        columns = ("title", "proc", "cls", "hwnd")
        # show="tree headings" + #0 칼럼을 아이콘 표시용으로 사용
        self.tree = ttk.Treeview(mid, columns=columns, show="tree headings")
        self.tree.heading("#0", text="")
        self.tree.heading("title", text="창 제목")
        self.tree.heading("proc", text="프로세스")
        self.tree.heading("cls", text="클래스")
        self.tree.heading("hwnd", text="HWND")

        self.tree.column("#0", width=40, stretch=False, anchor="center")  # ✅ 아이콘 칼럼 40px
        self.tree.column("title", width=640, anchor="w")
        self.tree.column("proc", width=200, anchor="w")
        self.tree.column("cls", width=200, anchor="w")
        self.tree.column("hwnd", width=120, anchor="e")

        vsb = ttk.Scrollbar(mid, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=vsb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.LEFT, fill=tk.Y)

        # 컨텍스트 메뉴 (모서리 이동 포함)
        self.menu = tk.Menu(self, tearoff=False)
        self.menu.add_command(label="중앙으로 이동", command=self.center_selected)
        self.menu.add_command(label="전면으로 가져오기", command=self.bring_to_front_selected)
        self.menu.add_separator()
        self.menu.add_command(label="최소화", command=self.minimize_selected)
        self.menu.add_command(label="최대화", command=self.maximize_selected)
        self.menu.add_command(label="복원", command=self.restore_selected)
        self.menu.add_separator()
        self.always_on_top_var = tk.BooleanVar(value=False)
        self.menu.add_checkbutton(label="항상 위", onvalue=True, offvalue=False,
                                  variable=self.always_on_top_var, command=self.toggle_topmost_selected)
        self.menu.add_separator()
        # 모서리 이동 서브메뉴 (모니터 좌표 기준, margin=0)
        self.corner_menu = tk.Menu(self.menu, tearoff=False)
        self.corner_menu.add_command(label="왼쪽 위 (Alt+1)", command=self.move_selected_top_left)
        self.corner_menu.add_command(label="왼쪽 아래 (Alt+2)", command=self.move_selected_bottom_left)
        self.corner_menu.add_command(label="오른쪽 위 (Alt+3)", command=self.move_selected_top_right)
        self.corner_menu.add_command(label="오른쪽 아래 (Alt+4)", command=self.move_selected_bottom_right)
        self.menu.add_cascade(label="모서리로 이동", menu=self.corner_menu)
        # 가장자리 이동 서브메뉴 (한 축만 이동)
        self.edge_menu = tk.Menu(self.menu, tearoff=False)
        self.edge_menu.add_command(label="맨 위로 (Alt+Up)", command=self.move_selected_to_top)
        self.edge_menu.add_command(label="맨 아래로 (Alt+Down)", command=self.move_selected_to_bottom)
        self.edge_menu.add_command(label="맨 왼쪽으로 (Alt+Left)", command=self.move_selected_to_left)
        self.edge_menu.add_command(label="맨 오른쪽으로 (Alt+Right)", command=self.move_selected_to_right)
        self.menu.add_cascade(label="가장자리로 이동", menu=self.edge_menu)
        self.menu.add_separator()
        # 크기 복사 메뉴
        self.size_menu = tk.Menu(self.menu, tearoff=False)
        self.size_menu.add_command(label="이 창의 크기 기억 (Ctrl+Shift+C)", command=self.remember_window_size)
        self.size_menu.add_command(label="기억된 크기 적용 (Ctrl+Shift+V)", command=self.apply_remembered_size)
        self.menu.add_cascade(label="창 크기 복사", menu=self.size_menu)
        # 위치 복사 메뉴
        self.position_menu = tk.Menu(self.menu, tearoff=False)
        self.position_menu.add_command(label="이 창의 위치 기억 (Ctrl+Alt+C)", command=self.remember_window_position)
        self.position_menu.add_command(label="기억된 위치 적용 (Ctrl+Alt+V)", command=self.apply_remembered_position)
        self.menu.add_cascade(label="창 위치 복사", menu=self.position_menu)
        self.menu.add_separator()
        self.menu.add_command(label="닫기 (Del)", command=self.close_selected)

        self.tree.bind("<Button-3>", self._on_right_click)
        self.tree.bind("<Double-1>", lambda e: self.center_selected())

        # 하단 상태바
        bottom = ttk.Frame(self, style="Naked.TFrame", padding=(12, 0, 12, 12))
        bottom.pack(fill=tk.X)
        self.status_label = ttk.Label(bottom, text="준비됨", style="Hint.TLabel")
        self.status_label.pack(side=tk.LEFT)

    # ----- 데이터 로드 -----
    def refresh_tree(self):
        query = self.search_var.get().strip().lower()
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        self.tk_images.clear()

        count = 0
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

            hay = f"{title} {proc_name} {class_name}".lower()
            if query and query not in hay:
                continue

            img = self.tk_images.get(hwnd)
            if img is None:
                img = get_hwnd_icon_image(hwnd, size=(18, 18))
                if img:
                    self.tk_images[hwnd] = img

            vals = (_tcl_safe(title), _tcl_safe(proc_name), _tcl_safe(class_name), str(hwnd))
            tag = "even" if count % 2 else "odd"

            insert_kwargs = {"text": "", "values": vals, "tags": (tag,)}
            if img is not None:
                insert_kwargs["image"] = img  # #0 칼럼 아이콘

            self.tree.insert("", tk.END, **insert_kwargs)
            count += 1

        self.status_label.config(text=f"표시된 창: {count}개  (F5 새로고침)")

    # ----- 선택 유틸 -----
    def _get_selected_hwnd_and_title(self):
        sel = self.tree.selection()
        if not sel:
            return None, None
        iid = sel[0]
        vals = self.tree.item(iid, "values")
        if not vals:
            return None, None
        title, _, _, hwnd_str = vals
        try:
            hwnd = int(hwnd_str)
        except Exception:
            hwnd = None
        return hwnd, title

    def _ensure_selection_at(self, event):
        iid = self.tree.identify_row(event.y)
        if iid:
            self.tree.selection_set(iid)
            self.tree.focus(iid)

    # ----- 컨텍스트 메뉴 -----
    def _on_right_click(self, event):
        self._ensure_selection_at(event)
        try:
            self.menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.menu.grab_release()

    # ----- 액션 -----
    def center_selected(self, *args):
        hwnd, title = self._get_selected_hwnd_and_title()
        if not hwnd:
            messagebox.showwarning("경고", "창을 선택해주세요.")
            return
        try:
            move_window_center_and_signal(hwnd)
            self._notify(f"'{title}' 창을 중앙으로 이동했습니다.")
        except Exception as e:
            messagebox.showerror("오류", f"창을 이동할 수 없습니다:\n{e}")

    def bring_to_front_selected(self, *args):
        hwnd, title = self._get_selected_hwnd_and_title()
        if not hwnd:
            return
        bring_window_to_front_by_hwnd(hwnd)
        self._notify(f"'{title}' 창을 전면으로 가져왔습니다.")

    def minimize_selected(self, *args):
        hwnd, title = self._get_selected_hwnd_and_title()
        if not hwnd:
            return
        try:
            win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
            self._notify(f"'{title}' 창을 최소화했습니다.")
        except Exception as e:
            messagebox.showerror("오류", f"최소화 실패:\n{e}")

    def maximize_selected(self, *args):
        hwnd, title = self._get_selected_hwnd_and_title()
        if not hwnd:
            return
        try:
            win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
            self._notify(f"'{title}' 창을 최대화했습니다.")
        except Exception as e:
            messagebox.showerror("오류", f"최대화 실패:\n{e}")

    def restore_selected(self, *args):
        hwnd, title = self._get_selected_hwnd_and_title()
        if not hwnd:
            return
        try:
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            self._notify(f"'{title}' 창을 복원했습니다.")
        except Exception as e:
            messagebox.showerror("오류", f"복원 실패:\n{e}")

    def toggle_topmost_selected(self, *args):
        hwnd, title = self._get_selected_hwnd_and_title()
        if not hwnd:
            return
        try:
            if self.always_on_top_var.get():
                win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                                      win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
                self._notify(f"'{title}' 창을 항상 위로 설정.")
            else:
                win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                                      win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
                self._notify(f"'{title}' 창의 항상 위 해제.")
        except Exception as e:
            messagebox.showerror("오류", f"항상 위 토글 실패:\n{e}")

    def close_selected(self, *args):
        hwnd, title = self._get_selected_hwnd_and_title()
        if not hwnd:
            return
        try:
            win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
            self._notify(f"'{title}' 창을 닫으라고 요청했습니다.")
            self.after(300, self.refresh_tree)
        except Exception as e:
            messagebox.showerror("오류", f"닫기 실패:\n{e}")

    # ----- 모서리 이동 액션 (모니터 좌표계, margin=0) -----
    def move_selected_top_left(self, *args):
        hwnd, title = self._get_selected_hwnd_and_title()
        if not hwnd:
            messagebox.showwarning("경고", "창을 선택해주세요.")
            return
        move_window_to_corner(hwnd, "top-left", margin=0)
        self._notify(f"'{title}' 창을 좌상단(모니터 좌표)으로 이동했습니다.")

    def move_selected_bottom_left(self, *args):
        hwnd, title = self._get_selected_hwnd_and_title()
        if not hwnd:
            messagebox.showwarning("경고", "창을 선택해주세요.")
            return
        move_window_to_corner(hwnd, "bottom-left", margin=0)
        self._notify(f"'{title}' 창을 좌하단(모니터 좌표)으로 이동했습니다.")

    def move_selected_top_right(self, *args):
        hwnd, title = self._get_selected_hwnd_and_title()
        if not hwnd:
            messagebox.showwarning("경고", "창을 선택해주세요.")
            return
        move_window_to_corner(hwnd, "top-right", margin=0)
        self._notify(f"'{title}' 창을 우상단(모니터 좌표)으로 이동했습니다.")

    def move_selected_bottom_right(self, *args):
        hwnd, title = self._get_selected_hwnd_and_title()
        if not hwnd:
            messagebox.showwarning("경고", "창을 선택해주세요.")
            return
        move_window_to_corner(hwnd, "bottom-right", margin=0)
        self._notify(f"'{title}' 창을 우하단(모니터 좌표)으로 이동했습니다.")

    # ----- 가장자리 이동 액션 (한 축만 이동) -----
    def move_selected_to_top(self, *args):
        """X축 유지, 화면 맨 위로 이동"""
        hwnd, title = self._get_selected_hwnd_and_title()
        if not hwnd:
            messagebox.showwarning("경고", "창을 선택해주세요.")
            return
        move_window_to_edge(hwnd, "top", margin=0)
        self._notify(f"'{title}' 창을 맨 위로 이동했습니다. (X축 유지)")

    def move_selected_to_bottom(self, *args):
        """X축 유지, 화면 맨 아래로 이동"""
        hwnd, title = self._get_selected_hwnd_and_title()
        if not hwnd:
            messagebox.showwarning("경고", "창을 선택해주세요.")
            return
        move_window_to_edge(hwnd, "bottom", margin=0)
        self._notify(f"'{title}' 창을 맨 아래로 이동했습니다. (X축 유지)")

    def move_selected_to_left(self, *args):
        """Y축 유지, 화면 맨 왼쪽으로 이동"""
        hwnd, title = self._get_selected_hwnd_and_title()
        if not hwnd:
            messagebox.showwarning("경고", "창을 선택해주세요.")
            return
        move_window_to_edge(hwnd, "left", margin=0)
        self._notify(f"'{title}' 창을 맨 왼쪽으로 이동했습니다. (Y축 유지)")

    def move_selected_to_right(self, *args):
        """Y축 유지, 화면 맨 오른쪽으로 이동"""
        hwnd, title = self._get_selected_hwnd_and_title()
        if not hwnd:
            messagebox.showwarning("경고", "창을 선택해주세요.")
            return
        move_window_to_edge(hwnd, "right", margin=0)
        self._notify(f"'{title}' 창을 맨 오른쪽으로 이동했습니다. (Y축 유지)")

    # ----- 창 크기 복사 기능 -----
    def remember_window_size(self, *args):
        """선택한 창의 크기를 기억합니다."""
        hwnd, title = self._get_selected_hwnd_and_title()
        if not hwnd:
            messagebox.showwarning("경고", "크기를 기억할 창을 선택해주세요.")
            return
        
        if not win32gui.IsWindow(hwnd):
            messagebox.showerror("오류", "유효하지 않은 창입니다.")
            return
        
        try:
            width, height = get_window_size(hwnd)
            self.saved_size = (width, height)
            self.saved_size_title = title
            self._notify(f"'{title}' 창의 크기({width} x {height})를 기억했습니다. Ctrl+Shift+V로 다른 창에 적용하세요.")
        except Exception as e:
            messagebox.showerror("오류", f"창 크기를 가져올 수 없습니다:\n{e}")

    def apply_remembered_size(self, *args):
        """기억된 크기를 선택한 창에 적용합니다."""
        hwnd, title = self._get_selected_hwnd_and_title()
        if not hwnd:
            messagebox.showwarning("경고", "크기를 적용할 창을 선택해주세요.")
            return
        
        if self.saved_size is None:
            messagebox.showwarning("경고", "먼저 다른 창의 크기를 기억해주세요.\n(우클릭 -> 창 크기 복사 -> 이 창의 크기 기억)")
            return
        
        if not win32gui.IsWindow(hwnd):
            messagebox.showerror("오류", "유효하지 않은 창입니다.")
            return
        
        try:
            width, height = self.saved_size
            apply_window_size(hwnd, width, height)
            self._notify(f"'{title}' 창에 기억된 크기({width} x {height}, 원본: '{self.saved_size_title}')를 적용했습니다.")
        except Exception as e:
            messagebox.showerror("오류", f"창 크기를 적용할 수 없습니다:\n{e}")

    # ----- 창 위치 복사 기능 -----
    def remember_window_position(self, *args):
        """선택한 창의 위치를 기억합니다."""
        hwnd, title = self._get_selected_hwnd_and_title()
        if not hwnd:
            messagebox.showwarning("경고", "위치를 기억할 창을 선택해주세요.")
            return

        if not win32gui.IsWindow(hwnd):
            messagebox.showerror("오류", "유효하지 않은 창입니다.")
            return

        try:
            x, y = get_window_position(hwnd)
            self.saved_position = (x, y)
            self.saved_position_title = title
            self._notify(f"'{title}' 창의 위치({x}, {y})를 기억했습니다. Ctrl+Alt+V로 다른 창에 적용하세요.")
        except Exception as e:
            messagebox.showerror("오류", f"창 위치를 가져올 수 없습니다:\n{e}")

    def apply_remembered_position(self, *args):
        """기억된 위치를 선택한 창에 적용합니다."""
        hwnd, title = self._get_selected_hwnd_and_title()
        if not hwnd:
            messagebox.showwarning("경고", "위치를 적용할 창을 선택해주세요.")
            return

        if self.saved_position is None:
            messagebox.showwarning("경고", "먼저 다른 창의 위치를 기억해주세요.\n(우클릭 -> 창 위치 복사 -> 이 창의 위치 기억)")
            return

        if not win32gui.IsWindow(hwnd):
            messagebox.showerror("오류", "유효하지 않은 창입니다.")
            return

        try:
            x, y = self.saved_position
            apply_window_position(hwnd, x, y)
            self._notify(f"'{title}' 창에 기억된 위치({x}, {y}, 원본: '{self.saved_position_title}')를 적용했습니다.")
        except Exception as e:
            messagebox.showerror("오류", f"창 위치를 적용할 수 없습니다:\n{e}")

    def _notify(self, text):
        self.status_label.config(text=text)

if __name__ == "__main__":
    App().mainloop()
