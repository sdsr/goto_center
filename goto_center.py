import tkinter as tk
from tkinter import messagebox
import pygetwindow as gw
import pyautogui
import time
import win32gui
import win32con


def get_window_list():
    """현재 열려있는 창 목록을 가져옵니다."""
    windows = gw.getAllWindows()
    window_list = [win for win in windows if win.title]
    return window_list


def bring_window_to_front(window):
    """선택한 창을 최상위로 올리고 활성화합니다."""
    try:
        hwnd = win32gui.FindWindow(None, window.title)
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)

        # 창의 위치 정보를 가져옵니다.
        win_x, win_y, win_width, win_height = (
            window.left,
            window.top,
            window.width,
            window.height,
        )

        # 창의 제목 표시줄 중앙(위쪽 약간)으로 마우스를 이동 후 클릭하여 활성화합니다.
        pyautogui.moveTo(win_x + win_width // 2, win_y + 10, duration=0.3)
        pyautogui.click()
        time.sleep(0.5)
        win32gui.SetForegroundWindow(hwnd)
    except Exception as e:
        messagebox.showerror("오류", f"창을 맨 앞으로 가져올 수 없습니다:\n{e}")


def drag_window_to_center(window):
    """선택한 창을 마우스 드래그를 이용해 화면 중앙으로 이동합니다."""
    screen_width, screen_height = pyautogui.size()
    try:
        win_x, win_y = window.left, window.top
        win_width, win_height = window.width, window.height
        new_x = (screen_width - win_width) // 2
        new_y = (screen_height - win_height) // 2

        bring_window_to_front(window)
        pyautogui.moveTo(win_x + win_width // 2, win_y + 10, duration=0.5)
        pyautogui.mouseDown(button="left")
        pyautogui.moveTo(new_x + win_width // 2, new_y + 10, duration=0.5)
        pyautogui.mouseUp(button="left")

        messagebox.showinfo("성공", f"'{window.title}' 창을 중앙으로 이동했습니다.")
    except Exception as e:
        messagebox.showerror("오류", f"창을 이동할 수 없습니다:\n{e}")


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("창 중앙 이동기")
        self.geometry("500x400")
        self.configure(bg="#f0f0f0")

        # 타이틀 레이블
        label = tk.Label(
            self, text="열려있는 창 목록", font=("Helvetica", 14, "bold"), bg="#f0f0f0"
        )
        label.pack(pady=10)

        # 창 목록을 표시할 리스트박스
        self.listbox = tk.Listbox(self, font=("Helvetica", 12), selectmode=tk.SINGLE)
        self.listbox.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        self.listbox.bind("<Double-Button-1>", self.on_double_click)

        # 버튼 프레임: 새로고침, 중앙 이동, 종료 버튼
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
        """선택한 창을 화면 중앙으로 이동하는 함수 호출."""
        window = self.get_selected_window()
        if window:
            drag_window_to_center(window)
        else:
            messagebox.showwarning("경고", "창을 선택해주세요.")

    def on_double_click(self, event):
        """리스트박스 더블클릭 시 창 중앙 이동 실행."""
        self.center_selected_window()


if __name__ == "__main__":
    app = App()
    app.mainloop()
