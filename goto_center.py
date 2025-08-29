import tkinter as tk
from tkinter import messagebox
import pygetwindow as gw
import time

import win32gui
import win32con
import win32api

# ===== 유틸 =====

def get_window_list():
    """현재 열려있는 창 목록을 가져옵니다."""
    windows = gw.getAllWindows()
    window_list = [win for win in windows if win.title and win._hWnd]
    return window_list

def bring_window_to_front_by_hwnd(hwnd):
    """윈도우를 복원하고 전면으로 (활성화 여부는 앱 정책에 좌우됨)."""
    try:
        # 최소화된 경우 복원
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        # 포그라운드로
        win32gui.SetForegroundWindow(hwnd)
    except Exception:
        # 포그라운드가 시스템 정책상 막힐 수 있음 → 그래도 무시
        pass

def _get_work_area_rect_for_hwnd(hwnd):
    """창이 걸쳐있는 모니터의 '작업 영역(작업표시줄 제외)'을 반환."""
    MONITOR_DEFAULTTONEAREST = 2
    hmonitor = win32api.MonitorFromWindow(hwnd, MONITOR_DEFAULTTONEAREST)
    mi = win32api.GetMonitorInfo(hmonitor)
    # rcWork = (left, top, right, bottom)
    return mi["Work"] if "Work" in mi else mi["Monitor"]

def move_window_center_and_signal(hwnd):
    """
    창을 현재 위치한 모니터의 '작업 영역' 기준으로 가운데 배치하고,
    일부 앱이 위치를 저장하도록 이동 종료 신호(WM_EXITSIZEMOVE)를 보냅니다.
    """
    # 현재 창 사각형
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    w, h = right - left, bottom - top

    # 창이 위치한 모니터의 워크 에어리어
    wk_left, wk_top, wk_right, wk_bottom = _get_work_area_rect_for_hwnd(hwnd)
    wk_w, wk_h = wk_right - wk_left, wk_bottom - wk_top

    # 중앙 좌표 계산
    x = wk_left + (wk_w - w) // 2
    y = wk_top + (wk_h - h) // 2

    # 복원 및 전면으로
    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    bring_window_to_front_by_hwnd(hwnd)

    # (선택) 이동 시작 신호 흉내 — 일부 앱은 ENTER/EXIT 쌍을 좋아함
    try:
        win32gui.PostMessage(hwnd, win32con.WM_ENTERSIZEMOVE, 0, 0)
    except Exception:
        pass

    # 위치 이동 (크기 변경 없음, 활성화는 강제하지 않음)
    flags = win32con.SWP_NOSIZE | win32con.SWP_NOACTIVATE
    win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, x, y, 0, 0, flags)

    # 약간의 지연으로 내부 처리 여유
    time.sleep(0.05)

    # 이동 종료 신호 → 일부 앱이 이 타이밍에 위치 저장
    try:
        win32gui.PostMessage(hwnd, win32con.WM_EXITSIZEMOVE, 0, 0)
    except Exception:
        pass

def center_window_by_pygetwindow_window(win):
    """pygetwindow의 Window 객체를 받아 가운데로 이동."""
    try:
        hwnd = win._hWnd  # pygetwindow 내부 보관 중인 윈도우 핸들
        if not hwnd:
            raise RuntimeError("유효한 윈도우 핸들을 찾을 수 없습니다.")
        move_window_center_and_signal(hwnd)
        messagebox.showinfo("성공", f"'{win.title}' 창을 중앙으로 이동했습니다.")
    except Exception as e:
        messagebox.showerror("오류", f"창을 이동할 수 없습니다:\n{e}")

# ===== Tk 앱 =====

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("창 중앙 이동기")
        self.geometry("500x400")
        self.configure(bg="#f0f0f0")

        label = tk.Label(
            self, text="열려있는 창 목록", font=("Helvetica", 14, "bold"), bg="#f0f0f0"
        )
        label.pack(pady=10)

        self.listbox = tk.Listbox(self, font=("Helvetica", 12), selectmode=tk.SINGLE)
        self.listbox.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        self.listbox.bind("<Double-Button-1>", self.on_double_click)

        btn_frame = tk.Frame(self, bg="#f0f0f0")
        btn_frame.pack(pady=10)

        refresh_btn = tk.Button(
            btn_frame,
            text="새로고침",
            font=("Helvetica", 12),
            command=self.update_window_list,
        )
        refresh_btn.pack(side=tk.LEFT, padx=10)

        center_btn = tk.Button(
            btn_frame,
            text="중앙으로 이동",
            font=("Helvetica", 12),
            command=self.center_selected_window,
        )
        center_btn.pack(side=tk.LEFT, padx=10)

        exit_btn = tk.Button(
            btn_frame, text="종료", font=("Helvetica", 12), command=self.destroy
        )
        exit_btn.pack(side=tk.LEFT, padx=10)

        self.window_list = []
        self.update_window_list()

    def update_window_list(self):
        """창 목록을 업데이트합니다."""
        self.listbox.delete(0, tk.END)
        self.window_list = get_window_list()
        for win in self.window_list:
            self.listbox.insert(tk.END, win.title)

    def get_selected_window(self):
        """리스트박스에서 선택한 창을 반환합니다."""
        selection = self.listbox.curselection()
        if selection:
            index = selection[0]
            return self.window_list[index]
        return None

    def center_selected_window(self):
        """선택한 창을 화면 중앙으로 이동."""
        win = self.get_selected_window()
        if win:
            center_window_by_pygetwindow_window(win)
        else:
            messagebox.showwarning("경고", "창을 선택해주세요.")

    def on_double_click(self, event):
        """리스트박스 더블클릭 시 창 중앙 이동 실행."""
        self.center_selected_window()

if __name__ == "__main__":
    app = App()
    app.mainloop()
