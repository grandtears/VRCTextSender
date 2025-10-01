import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import platform
import os
import sys
from typing import Optional

from osc_client import VRChatOSCClient
from vrc_monitor import is_vrchat_running
import config

def enable_dpi_awareness():
    """Windows DPI対応（見切れ防止）"""
    if platform.system() == "Windows":
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            try:
                ctypes.windll.user32.SetProcessDPIAware()
            except Exception:
                pass

def get_resource_path(filename):
    """リソースファイルのパスを取得（PyInstaller対応）"""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, filename)
    return os.path.join(os.path.dirname(__file__), filename)

def set_taskbar_icon(root):
    """タスクバーアイコンを設定"""
    # Windows環境でのタスクバー識別子設定
    if platform.system() == "Windows":
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("VRCTextSender.App")
        except Exception:
            pass
    
    # アイコンファイルの設定
    try:
        icon_path = get_resource_path(config.ICON_FILE)
        if os.path.exists(icon_path):
            # 基本的なPNGアイコン設定
            icon_img = tk.PhotoImage(file=icon_path)
            root.iconphoto(True, icon_img)  # タイトルバーとタスクバー両方に適用
            root._icon_ref = icon_img  # 参照保持
            
            # Pillowが利用可能な場合の高品質化（オプション）
            _enhance_icon_quality(root, icon_path)
            
    except Exception as e:
        print(f"アイコン設定エラー: {e}")

def _enhance_icon_quality(root, icon_path):
    """Pillowを使った高品質アイコン設定（オプション機能）"""
    if platform.system() != "Windows":
        return
        
    try:
        from PIL import Image
        
        # PNGをICOに変換
        pil_img = Image.open(icon_path)
        temp_ico = os.path.join(os.path.dirname(icon_path), "temp_icon.ico")
        
        # 複数サイズでICO作成
        ico_sizes = [(16, 16), (32, 32), (48, 48), (64, 64)]
        ico_images = []
        
        for size in ico_sizes:
            try:
                resample = Image.Resampling.LANCZOS
            except AttributeError:
                resample = Image.LANCZOS
            resized = pil_img.resize(size, resample)
            ico_images.append(resized)
        
        # ICOファイル作成と適用
        ico_images[0].save(
            temp_ico,
            format='ICO',
            sizes=ico_sizes,
            append_images=ico_images[1:]
        )
        
        root.iconbitmap(temp_ico)
        
        # 一時ファイルを遅延削除
        root.after(2000, lambda: _safe_remove_file(temp_ico))
        
    except ImportError:
        # Pillowが無い場合は何もしない
        pass
    except Exception:
        # ICO変換に失敗してもPNG版は動作
        pass

def _safe_remove_file(file_path):
    """ファイルを安全に削除"""
    try:
        if os.path.exists(file_path):
            os.remove(file_path)
    except Exception:
        pass

class MainWindow:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(config.APP_TITLE)
        self.root.geometry("560x420")
        self.root.minsize(520, 380)

        # DPI対応とタスクバーアイコン設定
        enable_dpi_awareness()
        set_taskbar_icon(self.root)

        # 状態管理
        self.client: Optional[VRChatOSCClient] = VRChatOSCClient(
            config.DEFAULT_IP, config.DEFAULT_PORT
        )
        self.vrchat_running = False
        self.monitoring = True

        self._build_ui()
        self._layout()
        self._update_counter()
        self._tick_monitor()

    def _build_ui(self):
        """UIウィジェットを構築"""
        self.main = ttk.Frame(self.root, padding=12)
        self.main.grid(sticky="nsew")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        self.main.columnconfigure(0, weight=1)
        self.main.rowconfigure(2, weight=1)

        # タイトル
        self.title_lbl = ttk.Label(
            self.main, 
            text="VRChat チャットボックス送信", 
            font=("Meiryo UI", 16, "bold")
        )

        # 接続設定フレーム
        self.conn_frame = ttk.LabelFrame(self.main, text="接続設定", padding=8)
        
        ttk.Label(self.conn_frame, text="IP Address:").grid(
            row=0, column=0, sticky="w"
        )
        self.ip_var = tk.StringVar(value=config.DEFAULT_IP)
        self.ip_entry = ttk.Entry(
            self.conn_frame, textvariable=self.ip_var, width=18
        )
        
        ttk.Label(self.conn_frame, text="Port:").grid(
            row=0, column=2, sticky="w", padx=(12, 0)
        )
        self.port_var = tk.StringVar(value=str(config.DEFAULT_PORT))
        self.port_entry = ttk.Entry(
            self.conn_frame, textvariable=self.port_var, width=8
        )
        
        self.update_btn = ttk.Button(
            self.conn_frame, text="接続更新", command=self._update_connection
        )

        # VRChat状態表示
        self.state_row = ttk.Frame(self.conn_frame)
        ttk.Label(self.state_row, text="VRChat状態:").pack(side="left")
        self.vrc_state_lbl = ttk.Label(
            self.state_row, 
            text="確認中...", 
            foreground="orange", 
            font=("Meiryo UI", 10, "bold")
        )
        self.vrc_state_lbl.pack(side="left", padx=(6, 0))

        # テキスト入力フレーム
        self.text_frame = ttk.LabelFrame(self.main, text="送信テキスト", padding=8)
        
        self.text_area = scrolledtext.ScrolledText(
            self.text_frame, height=8, font=("Meiryo UI", 11), wrap="word"
        )
        self.text_area.bind("<KeyRelease>", lambda e: self._update_counter())
        
        # オプション行
        self.options = ttk.Frame(self.text_frame)
        self.immediate_var = tk.BooleanVar(value=True)
        self.immediate_chk = ttk.Checkbutton(
            self.options, text="即座に送信", variable=self.immediate_var
        )
        self.counter_lbl = ttk.Label(
            self.options, text=f"0/{config.MAX_LENGTH}"
        )

        # ボタン行
        self.btn_row = ttk.Frame(self.main)
        self.send_btn = ttk.Button(
            self.btn_row, 
            text="VRChatに送信", 
            command=self._send_text, 
            state="disabled"
        )
        self.clear_btn = ttk.Button(
            self.btn_row, text="クリア", command=self._clear_text
        )

        # ステータスバー
        self.status_var = tk.StringVar(value="起動しました")
        self.status_bar = ttk.Label(
            self.main, 
            textvariable=self.status_var, 
            relief="sunken", 
            anchor="w",
            font=("Meiryo UI", 9)
        )

        # スタイル設定
        self._setup_styles()

        # Ctrl+Enterで送信
        self.text_area.bind("<Control-Return>", self._on_ctrl_enter)

    def _setup_styles(self):
        """ttk スタイルを設定"""
        style = ttk.Style()
        
        available_themes = style.theme_names()
        if 'vista' in available_themes:
            style.theme_use('vista')
        elif 'clam' in available_themes:
            style.theme_use('clam')
        
        style.configure('Accent.TButton', font=('Meiryo UI', 11, 'bold'))
        self.send_btn.configure(style='Accent.TButton')

    def _layout(self):
        """レイアウト設定"""
        self.title_lbl.grid(row=0, column=0, sticky="ew", pady=(0, 12))

        self.conn_frame.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        self.conn_frame.columnconfigure(1, weight=1)
        self.ip_entry.grid(row=0, column=1, sticky="ew", padx=(6, 6))
        self.port_entry.grid(row=0, column=3, padx=(6, 6))
        self.update_btn.grid(row=0, column=4, sticky="e")
        self.state_row.grid(row=1, column=0, columnspan=5, sticky="w", pady=(6, 0))

        self.text_frame.grid(row=2, column=0, sticky="nsew", pady=(0, 10))
        self.text_frame.columnconfigure(0, weight=1)
        self.text_frame.rowconfigure(0, weight=1)
        self.text_area.grid(row=0, column=0, sticky="nsew", pady=(0, 8))
        
        self.options.grid(row=1, column=0, sticky="ew")
        self.options.columnconfigure(0, weight=1)
        self.immediate_chk.grid(row=0, column=0, sticky="w")
        self.counter_lbl.grid(row=0, column=1, sticky="e")

        self.btn_row.grid(row=3, column=0, pady=(0, 10))
        self.send_btn.pack(side="left", padx=(0, 8))
        self.clear_btn.pack(side="left")

        self.status_bar.grid(row=4, column=0, sticky="ew")

    def _tick_monitor(self):
        """VRChat監視処理"""
        if not self.monitoring:
            return
            
        running = is_vrchat_running()
        if running != self.vrchat_running:
            self.vrchat_running = running
            if running:
                self.vrc_state_lbl.config(text="起動中 ✓", foreground="green")
                self.status_var.set("VRChat起動中 - 送信できます")
            else:
                self.vrc_state_lbl.config(text="未起動 ✗", foreground="red")
                self.status_var.set("VRChatを起動してください")
            self._update_send_btn_state()
        
        self.root.after(config.CHECK_INTERVAL_MS, self._tick_monitor)

    def _update_connection(self):
        """接続設定更新"""
        try:
            ip = self.ip_var.get().strip()
            port = int(self.port_var.get())
            
            if not ip:
                raise ValueError("IPアドレスが空です")
            
            if not (1 <= port <= 65535):
                raise ValueError("ポート番号は1-65535の範囲で入力してください")
            
            self.client.set_target(ip, port)
            self.status_var.set(f"接続先を更新: {ip}:{port}")
        except ValueError as e:
            messagebox.showerror("設定エラー", str(e))
            self.status_var.set("接続設定エラー")
        except Exception as e:
            messagebox.showerror("接続エラー", f"接続設定の更新に失敗しました\n{str(e)}")
            self.status_var.set("接続エラー")

    def _update_counter(self):
        """文字数カウンター更新"""
        text = self.text_area.get("1.0", "end-1c").strip()
        n = len(text)
        self.counter_lbl.config(
            text=f"{n}/{config.MAX_LENGTH}",
            foreground=(
                "red" if n > config.MAX_LENGTH 
                else ("orange" if n > 120 else "black")
            )
        )
        self._update_send_btn_state()

    def _update_send_btn_state(self):
        """送信ボタン状態制御"""
        text = self.text_area.get("1.0", "end-1c").strip()
        enable = (
            self.vrchat_running and 
            (0 < len(text) <= config.MAX_LENGTH)
        )
        self.send_btn.config(state=("normal" if enable else "disabled"))

    def _send_text(self):
        """テキスト送信"""
        if not self.vrchat_running:
            messagebox.showwarning(
                "VRChat未起動", 
                "VRChatを起動してOSCを有効にしてください。"
            )
            return
        
        text = self.text_area.get("1.0", "end-1c").strip()
        if not text:
            self.status_var.set("テキストが空です")
            return
        
        if len(text) > config.MAX_LENGTH:
            messagebox.showwarning(
                "文字数超過", 
                f"{config.MAX_LENGTH}文字以内で入力してください"
            )
            return
        
        try:
            self.client.send_chat(text, enter=self.immediate_var.get())
            if self.immediate_var.get():
                self._clear_text()
            
            preview = text[:25] + ("..." if len(text) > 25 else "")
            self.status_var.set(f"✓ 送信完了: {preview}")
            
        except Exception as e:
            error_msg = f"送信エラー: {str(e)}"
            self.status_var.set(error_msg)
            
            messagebox.showerror(
                "送信エラー", 
                "OSC送信に失敗しました。\n\n"
                "確認事項:\n"
                "• VRChatでOSCが有効になっているか\n"
                "• IP/ポート設定が正しいか\n"
                "• ファイアウォールの設定\n\n"
                f"エラー詳細: {str(e)}"
            )

    def _clear_text(self):
        """テキストクリア"""
        self.text_area.delete("1.0", "end")
        self._update_counter()
        self.status_var.set("テキストをクリアしました")

    def _on_ctrl_enter(self, event):
        """Ctrl+Enter送信"""
        if self.send_btn["state"] == "normal":
            self._send_text()
        return "break"

    def on_closing(self):
        """アプリケーション終了処理"""
        self.monitoring = False
        self.root.destroy()
