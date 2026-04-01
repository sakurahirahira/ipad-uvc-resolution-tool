"""
iPad UVC 解像度設定ツール
iPadをUVC外部ディスプレイとして使う際に、ぴったりの解像度を設定するツール
"""

import ctypes
import ctypes.wintypes
import tkinter as tk
from tkinter import ttk, messagebox
import struct

# Windows API 定義
user32 = ctypes.windll.user32
DM_PELSWIDTH = 0x80000
DM_PELSHEIGHT = 0x100000
DM_DISPLAYFREQUENCY = 0x400000
CDS_UPDATEREGISTRY = 0x01
CDS_TEST = 0x02
DISP_CHANGE_SUCCESSFUL = 0
DISP_CHANGE_RESTART = 1
ENUM_CURRENT_SETTINGS = -1

# iPad 解像度プリセット
IPAD_PRESETS = {
    "iPad Pro 13\" (M4) ネイティブ": (2752, 2064),
    "iPad Pro 13\" (M4) 半分": (1376, 1032),
    "iPad Pro 12.9\" / Air 13\" ネイティブ": (2732, 2048),
    "iPad Pro 12.9\" / Air 13\" 半分": (1366, 1024),
    "iPad 4:3 (2048x1536)": (2048, 1536),
    "iPad 4:3 半分 (1024x768)": (1024, 768),
    "XGA 4:3 (1280x960)": (1280, 960),
    "SXGA 4:3 (1400x1050)": (1400, 1050),
}


class DEVMODEW(ctypes.Structure):
    _fields_ = [
        ("dmDeviceName", ctypes.c_wchar * 32),
        ("dmSpecVersion", ctypes.c_ushort),
        ("dmDriverVersion", ctypes.c_ushort),
        ("dmSize", ctypes.c_ushort),
        ("dmDriverExtra", ctypes.c_ushort),
        ("dmFields", ctypes.c_uint),
        ("dmPositionX", ctypes.c_int),
        ("dmPositionY", ctypes.c_int),
        ("dmDisplayOrientation", ctypes.c_uint),
        ("dmDisplayFixedOutput", ctypes.c_uint),
        ("dmColor", ctypes.c_short),
        ("dmDuplex", ctypes.c_short),
        ("dmYResolution", ctypes.c_short),
        ("dmTTOption", ctypes.c_short),
        ("dmCollate", ctypes.c_short),
        ("dmFormName", ctypes.c_wchar * 32),
        ("dmLogPixels", ctypes.c_ushort),
        ("dmBitsPerPel", ctypes.c_uint),
        ("dmPelsWidth", ctypes.c_uint),
        ("dmPelsHeight", ctypes.c_uint),
        ("dmDisplayFlags", ctypes.c_uint),
        ("dmDisplayFrequency", ctypes.c_uint),
        ("dmICMMethod", ctypes.c_uint),
        ("dmICMIntent", ctypes.c_uint),
        ("dmMediaType", ctypes.c_uint),
        ("dmDitherType", ctypes.c_uint),
        ("dmReserved1", ctypes.c_uint),
        ("dmReserved2", ctypes.c_uint),
        ("dmPanningWidth", ctypes.c_uint),
        ("dmPanningHeight", ctypes.c_uint),
    ]


class DISPLAY_DEVICEW(ctypes.Structure):
    _fields_ = [
        ("cb", ctypes.c_uint),
        ("DeviceName", ctypes.c_wchar * 32),
        ("DeviceString", ctypes.c_wchar * 128),
        ("StateFlags", ctypes.c_uint),
        ("DeviceID", ctypes.c_wchar * 128),
        ("DeviceKey", ctypes.c_wchar * 128),
    ]


def get_displays():
    """接続中のディスプレイ一覧を取得"""
    displays = []
    i = 0
    while True:
        dd = DISPLAY_DEVICEW()
        dd.cb = ctypes.sizeof(dd)
        if not user32.EnumDisplayDevicesW(None, i, ctypes.byref(dd), 0):
            break
        # アクティブなディスプレイのみ
        if dd.StateFlags & 0x01:  # DISPLAY_DEVICE_ATTACHED_TO_DESKTOP
            displays.append({
                "index": i,
                "name": dd.DeviceName.rstrip('\x00'),
                "description": dd.DeviceString.rstrip('\x00'),
                "flags": dd.StateFlags,
                "primary": bool(dd.StateFlags & 0x04),
            })
        i += 1
    return displays


def get_current_resolution(device_name):
    """現在の解像度を取得"""
    dm = DEVMODEW()
    dm.dmSize = ctypes.sizeof(dm)
    if user32.EnumDisplaySettingsW(device_name, ENUM_CURRENT_SETTINGS, ctypes.byref(dm)):
        return dm.dmPelsWidth, dm.dmPelsHeight, dm.dmDisplayFrequency
    return None


def get_supported_modes(device_name):
    """サポートされている全解像度モードを取得"""
    modes = set()
    i = 0
    while True:
        dm = DEVMODEW()
        dm.dmSize = ctypes.sizeof(dm)
        if not user32.EnumDisplaySettingsW(device_name, i, ctypes.byref(dm)):
            break
        modes.add((dm.dmPelsWidth, dm.dmPelsHeight, dm.dmDisplayFrequency, dm.dmBitsPerPel))
        i += 1
    return sorted(modes, key=lambda m: (m[0] * m[1], m[2]), reverse=True)


def find_best_match(modes, target_w, target_h):
    """ターゲット解像度に最も近いモードを見つける"""
    target_ratio = target_w / target_h
    best = None
    best_score = float('inf')

    for w, h, freq, bpp in modes:
        ratio = w / h
        # アスペクト比の差 + 解像度の差でスコアリング
        ratio_diff = abs(ratio - target_ratio)
        size_diff = abs(w - target_w) + abs(h - target_h)
        score = ratio_diff * 10000 + size_diff
        if score < best_score:
            best_score = score
            best = (w, h, freq, bpp)

    return best


def set_resolution(device_name, width, height, freq=None):
    """解像度を変更する"""
    dm = DEVMODEW()
    dm.dmSize = ctypes.sizeof(dm)
    dm.dmPelsWidth = width
    dm.dmPelsHeight = height
    dm.dmFields = DM_PELSWIDTH | DM_PELSHEIGHT
    if freq:
        dm.dmDisplayFrequency = freq
        dm.dmFields |= DM_DISPLAYFREQUENCY

    # まずテスト
    result = user32.ChangeDisplaySettingsExW(
        device_name, ctypes.byref(dm), None, CDS_TEST, None
    )
    if result != DISP_CHANGE_SUCCESSFUL:
        return False, f"テスト失敗 (コード: {result})。この解像度はサポートされていません。"

    # 実際に変更
    result = user32.ChangeDisplaySettingsExW(
        device_name, ctypes.byref(dm), None, CDS_UPDATEREGISTRY, None
    )
    if result == DISP_CHANGE_SUCCESSFUL:
        return True, "解像度を変更しました！"
    elif result == DISP_CHANGE_RESTART:
        return True, "解像度を変更しました。完全に反映するには再起動が必要です。"
    else:
        return False, f"変更失敗 (コード: {result})"


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("iPad UVC 解像度設定ツール")
        self.root.geometry("780x700")
        self.root.resizable(True, True)

        style = ttk.Style()
        style.configure("Header.TLabel", font=("Meiryo UI", 14, "bold"))
        style.configure("Info.TLabel", font=("Meiryo UI", 10))
        style.configure("Big.TButton", font=("Meiryo UI", 11), padding=8)

        self.create_widgets()
        self.refresh_displays()

    def create_widgets(self):
        main = ttk.Frame(self.root, padding=15)
        main.pack(fill=tk.BOTH, expand=True)

        # --- ディスプレイ選択 ---
        ttk.Label(main, text="iPad UVC 解像度設定ツール", style="Header.TLabel").pack(anchor=tk.W)
        ttk.Separator(main, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=(5, 10))

        disp_frame = ttk.LabelFrame(main, text="ディスプレイ選択", padding=10)
        disp_frame.pack(fill=tk.X, pady=5)

        self.display_var = tk.StringVar()
        self.display_combo = ttk.Combobox(disp_frame, textvariable=self.display_var,
                                          state="readonly", width=60)
        self.display_combo.pack(side=tk.LEFT, padx=(0, 10))
        self.display_combo.bind("<<ComboboxSelected>>", self.on_display_selected)

        ttk.Button(disp_frame, text="更新", command=self.refresh_displays).pack(side=tk.LEFT)

        # 現在の解像度表示
        self.current_res_var = tk.StringVar(value="ディスプレイを選択してください")
        ttk.Label(main, textvariable=self.current_res_var, style="Info.TLabel").pack(anchor=tk.W, pady=5)

        # --- iPadプリセット ---
        preset_frame = ttk.LabelFrame(main, text="iPad プリセット (ワンクリック適用)", padding=10)
        preset_frame.pack(fill=tk.X, pady=5)

        row = 0
        col = 0
        for name, (w, h) in IPAD_PRESETS.items():
            btn = ttk.Button(preset_frame, text=f"{name}\n({w}x{h})",
                             command=lambda w=w, h=h, n=name: self.apply_preset(w, h, n))
            btn.grid(row=row, column=col, padx=5, pady=3, sticky=tk.EW)
            col += 1
            if col >= 2:
                col = 0
                row += 1
        preset_frame.columnconfigure(0, weight=1)
        preset_frame.columnconfigure(1, weight=1)

        # --- カスタム解像度 ---
        custom_frame = ttk.LabelFrame(main, text="カスタム解像度", padding=10)
        custom_frame.pack(fill=tk.X, pady=5)

        input_row = ttk.Frame(custom_frame)
        input_row.pack(fill=tk.X)

        ttk.Label(input_row, text="幅:").pack(side=tk.LEFT)
        self.width_var = tk.StringVar(value="2732")
        ttk.Entry(input_row, textvariable=self.width_var, width=8).pack(side=tk.LEFT, padx=(2, 10))

        ttk.Label(input_row, text="高さ:").pack(side=tk.LEFT)
        self.height_var = tk.StringVar(value="2048")
        ttk.Entry(input_row, textvariable=self.height_var, width=8).pack(side=tk.LEFT, padx=(2, 10))

        ttk.Button(input_row, text="適用", command=self.apply_custom).pack(side=tk.LEFT, padx=10)
        ttk.Button(input_row, text="最も近い解像度を検索",
                   command=self.find_closest).pack(side=tk.LEFT)

        # --- サポート解像度一覧 ---
        modes_frame = ttk.LabelFrame(main, text="サポートされている解像度一覧（4:3に近い順）", padding=10)
        modes_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        columns = ("resolution", "freq", "bpp", "ratio")
        self.modes_tree = ttk.Treeview(modes_frame, columns=columns, show="headings", height=10)
        self.modes_tree.heading("resolution", text="解像度")
        self.modes_tree.heading("freq", text="周波数 (Hz)")
        self.modes_tree.heading("bpp", text="色深度")
        self.modes_tree.heading("ratio", text="アスペクト比")
        self.modes_tree.column("resolution", width=160)
        self.modes_tree.column("freq", width=100)
        self.modes_tree.column("bpp", width=80)
        self.modes_tree.column("ratio", width=120)

        scrollbar = ttk.Scrollbar(modes_frame, orient=tk.VERTICAL, command=self.modes_tree.yview)
        self.modes_tree.configure(yscrollcommand=scrollbar.set)
        self.modes_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.modes_tree.bind("<Double-1>", self.on_mode_double_click)

        ttk.Label(main, text="※ 一覧の項目をダブルクリックで適用 / UVCデバイスが報告する解像度のみ表示",
                  style="Info.TLabel").pack(anchor=tk.W)

        # --- ステータスバー ---
        self.status_var = tk.StringVar(value="準備完了")
        status_bar = ttk.Label(main, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(fill=tk.X, pady=(5, 0))

    def refresh_displays(self):
        self.displays = get_displays()
        values = []
        for d in self.displays:
            primary = " [プライマリ]" if d["primary"] else ""
            cur = get_current_resolution(d["name"])
            res_str = f" ({cur[0]}x{cur[1]}@{cur[2]}Hz)" if cur else ""
            values.append(f"{d['name']} - {d['description']}{primary}{res_str}")
        self.display_combo["values"] = values
        if values:
            # UVC / 非プライマリを優先選択
            idx = 0
            for i, d in enumerate(self.displays):
                if not d["primary"]:
                    idx = i
                    break
            self.display_combo.current(idx)
            self.on_display_selected(None)
        self.status_var.set(f"{len(self.displays)} 個のディスプレイを検出")

    def on_display_selected(self, event):
        idx = self.display_combo.current()
        if idx < 0:
            return
        d = self.displays[idx]
        cur = get_current_resolution(d["name"])
        if cur:
            w, h = cur[0], cur[1]
            ratio = self._ratio_str(w, h)
            self.current_res_var.set(
                f"現在の解像度: {w} x {h} @ {cur[2]}Hz  (アスペクト比: {ratio})"
            )
        self._refresh_modes(d["name"])

    def _ratio_str(self, w, h):
        from math import gcd
        g = gcd(w, h)
        return f"{w // g}:{h // g}"

    def _refresh_modes(self, device_name):
        for item in self.modes_tree.get_children():
            self.modes_tree.delete(item)
        modes = get_supported_modes(device_name)
        self.current_modes = modes

        # 4:3比率（1.333）に近い順にソート
        target_ratio = 4.0 / 3.0
        sorted_modes = sorted(modes, key=lambda m: (abs(m[0] / m[1] - target_ratio), -(m[0] * m[1])))

        for w, h, freq, bpp in sorted_modes:
            ratio = self._ratio_str(w, h)
            ratio_val = w / h
            # 4:3に近いものを強調
            tag = "match43" if abs(ratio_val - target_ratio) < 0.01 else ""
            self.modes_tree.insert("", tk.END,
                                   values=(f"{w} x {h}", f"{freq}", f"{bpp}bit", ratio),
                                   tags=(tag,))
        self.modes_tree.tag_configure("match43", background="#d4edda")

    def apply_preset(self, width, height, name):
        idx = self.display_combo.current()
        if idx < 0:
            messagebox.showwarning("警告", "ディスプレイを選択してください")
            return
        device_name = self.displays[idx]["name"]

        if messagebox.askyesno("確認", f"{name}\n{width}x{height} を適用しますか？"):
            success, msg = set_resolution(device_name, width, height)
            if success:
                self.status_var.set(f"✓ {msg}")
                messagebox.showinfo("成功", msg)
                self.on_display_selected(None)
            else:
                self.status_var.set(f"✗ {msg}")
                # サポートされていない場合、最も近い解像度を提案
                modes = get_supported_modes(device_name)
                best = find_best_match(modes, width, height)
                if best:
                    suggest = f"\n\n最も近いサポート解像度:\n{best[0]}x{best[1]} @ {best[2]}Hz"
                    if messagebox.askyesno("エラー",
                                           f"{msg}{suggest}\n\nこの解像度を適用しますか？"):
                        success2, msg2 = set_resolution(device_name, best[0], best[1], best[2])
                        if success2:
                            self.status_var.set(f"✓ {msg2}")
                            self.on_display_selected(None)
                        else:
                            messagebox.showerror("エラー", msg2)
                else:
                    messagebox.showerror("エラー", msg)

    def apply_custom(self):
        idx = self.display_combo.current()
        if idx < 0:
            messagebox.showwarning("警告", "ディスプレイを選択してください")
            return
        try:
            w = int(self.width_var.get())
            h = int(self.height_var.get())
        except ValueError:
            messagebox.showerror("エラー", "幅と高さに数値を入力してください")
            return

        device_name = self.displays[idx]["name"]
        if messagebox.askyesno("確認", f"{w}x{h} を適用しますか？"):
            success, msg = set_resolution(device_name, w, h)
            if success:
                self.status_var.set(f"✓ {msg}")
                messagebox.showinfo("成功", msg)
                self.on_display_selected(None)
            else:
                self.status_var.set(f"✗ {msg}")
                messagebox.showerror("エラー", msg)

    def find_closest(self):
        idx = self.display_combo.current()
        if idx < 0:
            messagebox.showwarning("警告", "ディスプレイを選択してください")
            return
        try:
            tw = int(self.width_var.get())
            th = int(self.height_var.get())
        except ValueError:
            messagebox.showerror("エラー", "幅と高さに数値を入力してください")
            return

        device_name = self.displays[idx]["name"]
        modes = get_supported_modes(device_name)
        best = find_best_match(modes, tw, th)
        if best:
            ratio = self._ratio_str(best[0], best[1])
            msg = (f"最も近い解像度:\n"
                   f"{best[0]} x {best[1]} @ {best[2]}Hz ({best[3]}bit)\n"
                   f"アスペクト比: {ratio}\n\n"
                   f"この解像度を適用しますか？")
            if messagebox.askyesno("検索結果", msg):
                success, msg2 = set_resolution(device_name, best[0], best[1], best[2])
                if success:
                    self.status_var.set(f"✓ {msg2}")
                    self.on_display_selected(None)
                else:
                    messagebox.showerror("エラー", msg2)
        else:
            messagebox.showinfo("結果", "サポートされている解像度が見つかりません")

    def on_mode_double_click(self, event):
        sel = self.modes_tree.selection()
        if not sel:
            return
        values = self.modes_tree.item(sel[0], "values")
        res_parts = values[0].replace(" ", "").split("x")
        w, h = int(res_parts[0]), int(res_parts[1])
        freq = int(values[1])

        idx = self.display_combo.current()
        if idx < 0:
            return
        device_name = self.displays[idx]["name"]

        if messagebox.askyesno("確認", f"{w}x{h} @ {freq}Hz を適用しますか？"):
            success, msg = set_resolution(device_name, w, h, freq)
            if success:
                self.status_var.set(f"✓ {msg}")
                messagebox.showinfo("成功", msg)
                self.on_display_selected(None)
            else:
                self.status_var.set(f"✗ {msg}")
                messagebox.showerror("エラー", msg)


def main():
    root = tk.Tk()
    # DPI対応
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    app = App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
