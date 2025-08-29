# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, messagebox
import time
import psutil
import ctypes

import pygetwindow as gw
from PIL import Image, ImageTk

import win32gui
import win32con
import win32api
import win32process
import win32ui


# ========= DPI 인식 =========
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


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

    flags = win32con.SWP_NOSIZE | win32con.SWP_NOACTIVATE
    win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, x, y, 0, 0, flags)

    time.sleep(0.05)
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
    img = Image.frombuffer("RGB", (bmp_info["bmWidth"], bmp_info["bmHeight"]), bmp_bytes, "raw", "BGRX", 0, 1)
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
        self.geometry("1200x620")          # <-- 기본 가로 더 큼
        self.minsize(800, 520)

        self._build_style_light()
        self._build_ui()

        self.tk_images = {}  # hwnd -> PhotoImage
        self.refresh_tree()

        # 단축키
        self.bind("<Return>", lambda e: self.center_selected())
        self.bind("<F5>", lambda e: self.refresh_tree())
        self.bind("<Delete>", lambda e: self.close_selected())
        self.bind("<Control-l>", lambda e: (self.search_entry.focus_set(), self.search_entry.select_range(0, "end")))

    # ----- 라이트 테마 -----
    def _build_style_light(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass

        bg      = "#F7F9FC"  # window
        panel   = "#FFFFFF"  # panels/cards
        border  = "#DFE3EB"
        txt     = "#1B2430"
        subtxt  = "#6B7280"
        selbg   = "#E6EFFB"
        header  = "#EEF2F7"
        row1    = "#FFFFFF"
        row2    = "#F4F6FA"
        focus   = "#3B82F6"

        self.configure(bg=bg)

        style.configure(".", background=bg, foreground=txt)

        # Frames / Labels
        style.configure("Light.TFrame", background=panel, borderwidth=1, relief="solid")
        style.configure("Naked.TFrame", background=bg, borderwidth=0)
        style.configure("Light.TLabel", background=panel, foreground=txt)
        style.configure("Hint.TLabel", background=bg, foreground=subtxt)

        # Buttons / Entry
        style.configure("TButton", padding=6)
        style.map("TButton", background=[("active", "#F1F5F9")])
        style.configure("TEntry", fieldbackground="#FFFFFF", foreground=txt)

        # Treeview
        style.configure("Treeview",
                        background=row1, fieldbackground=row1, foreground=txt, rowheight=28,
                        borderwidth=0)
        style.configure("Treeview.Heading",
                        background=header, foreground=txt, relief="flat")
        style.map("Treeview",
                  background=[("selected", selbg)],
                  foreground=[("selected", txt)])
        style.layout("Treeview", [
            ("Treeview.treearea", {"sticky": "nswe"})
        ])

        # 스트라이프용 태그
        style.configure("odd.Treeview", background=row1)
        style.configure("even.Treeview", background=row2)

        # Scrollbar
        style.configure("Vertical.TScrollbar", background="#E5E7EB", troughcolor="#E5E7EB")
        style.map("Vertical.TScrollbar", background=[("active", "#D1D5DB")])

    # ----- UI 구성 -----
    def _build_ui(self):
        # 상단 바 (카드형)
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
        # 아이콘을 표시하려면 show="tree headings" + #0 사용
        self.tree = ttk.Treeview(mid, columns=columns, show="tree headings")
        self.tree.heading("#0", text="")  # 아이콘 칼럼 (헤더 텍스트 비움)
        self.tree.heading("title", text="창 제목")
        self.tree.heading("proc", text="프로세스")
        self.tree.heading("cls", text="클래스")
        self.tree.heading("hwnd", text="HWND")

        self.tree.column("#0", width=40, stretch=False, anchor="center")
        self.tree.column("title", width=540, anchor="w")
        self.tree.column("proc", width=180, anchor="w")
        self.tree.column("cls", width=180, anchor="w")
        self.tree.column("hwnd", width=120, anchor="e")

        vsb = ttk.Scrollbar(mid, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=vsb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.LEFT, fill=tk.Y)

        # 컨텍스트 메뉴
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

            # 아이콘 생성
            img = self.tk_images.get(hwnd)
            if img is None:
                img = get_hwnd_icon_image(hwnd, size=(18, 18))
                if img:
                    self.tk_images[hwnd] = img

            # 안전 문자열
            vals = (_tcl_safe(title), _tcl_safe(proc_name), _tcl_safe(class_name), str(hwnd))

            # 스트라이프 태그
            tag = "even" if count % 2 else "odd"

            # 아이콘을 #0 칼럼(image)에 세팅
            insert_kwargs = {"text": "", "values": vals, "tags": (tag,)}
            if img is not None:
                insert_kwargs["image"] = img

            iid = self.tree.insert("", tk.END, **insert_kwargs)

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

    def _notify(self, text):
        self.status_label.config(text=text)


if __name__ == "__main__":
    App().mainloop()
