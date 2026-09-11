import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk
import pandas as pd
import random
import os
import json
import shutil
import logging
import threading
from collections import deque
from datetime import datetime
import openpyxl
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter
from tkinterdnd2 import DND_FILES, TkinterDnD

# ==================== 设计系统 Design Tokens ====================
# 所有颜色均为 (浅色模式, 深色模式) 元组
PALETTE = {
    "bg":           ("#F1F5F9", "#0F172A"),
    "sidebar":      ("#FFFFFF", "#1E293B"),
    "card":         ("#FFFFFF", "#1E293B"),
    "card_soft":    ("#F8FAFC", "#16203A"),
    "border":       ("#E2E8F0", "#334155"),
    "text":         ("#0F172A", "#F1F5F9"),
    "text_2":       ("#475569", "#94A3B8"),
    "text_3":       ("#94A3B8", "#64748B"),
    "primary":      ("#3B82F6", "#3B82F6"),
    "primary_hv":   ("#2563EB", "#2563EB"),
    "primary_soft": ("#EFF6FF", "#1E3A5F"),
    "success":      ("#10B981", "#10B981"),
    "success_hv":   ("#059669", "#059669"),
    "warn":         ("#D97706", "#F59E0B"),
    "danger":       ("#EF4444", "#F87171"),
    "danger_soft":  ("#FEF2F2", "#431818"),
    "row_fg":       ("#334155", "#E2E8F0"),
    "row_even":     ("#FFFFFF", "#1E293B"),
    "row_odd":      ("#F8FAFC", "#24324E"),
    "row_hover":    ("#EFF6FF", "#2C3D61"),
    "row_sel":      ("#DBEAFE", "#1D4ED8"),
    "row_sel_fg":   ("#1D4ED8", "#FFFFFF"),
    "row_error":    ("#FEF2F2", "#4C1D1D"),
    "row_error_fg": ("#DC2626", "#FCA5A5"),
}

MAX_HISTORY = 80          # 撤销栈上限，防止内存无限增长
CACHE_DEBOUNCE_MS = 800   # 缓存写入防抖间隔


class CustomDnDTk(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.TkdndVersion = TkinterDnD._require(self)


class SeatShufflerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Smart Seat Shuffler · 智能班级排座系统")
        self.root.geometry("1180x780")
        self.root.minsize(920, 620)
        self.root.configure(fg_color=PALETTE["bg"])

        self.root.attributes("-alpha", 0.0)
        self.fade_in()

        self.df = None
        self.current_sort_state = "未导入数据"
        self.history_stack = deque(maxlen=MAX_HISTORY)
        self.current_frame_name = "home"

        self._cache_job = None
        self._hover_iid = None
        self._loading_dots = 0
        self._loading_base_text = ""

        # --- 默认配置参数 ---
        self.config_file = "settings.json"
        self.auto_backup = True
        self.backup_path = os.path.join(os.getcwd(), "备份文件夹")
        self.cache_path = os.path.join(os.getcwd(), "缓存文件夹")
        self.log_path = os.path.join(os.getcwd(), "日志文件夹")

        self.row_height = 21.0
        self.font_size = 14
        self.col_w_a = 6.0
        self.col_w_b = 7.5
        self.col_w_c = 14.0
        self.col_w_d = 9.0
        self.col_w_e = 8.0
        self.align_mode = "居中"
        self.col_name_a = "序号"
        self.col_name_b = "班级"
        self.col_name_c = "学号"
        self.col_name_d = "姓名"
        self.col_name_e = "座位号"

        self.appearance_mode = "System"
        self.scroll_speed = 50

        # --- 快捷键配置 ---
        self.mod_undo = "Control"
        self.key_undo = "z"
        self.mod_export = "Control"
        self.key_export = "s"
        self.mod_shuffle = "Control"
        self.key_shuffle = "r"

        self._prev_mod_undo = None
        self._prev_key_undo = None
        self._prev_mod_export = None
        self._prev_key_export = None
        self._prev_mod_shuffle = None
        self._prev_key_shuffle = None

        # 按钮显隐状态
        self.show_download = True
        self.show_upload = True
        self.show_export = True
        self.show_shuffle = True
        self.show_undo = True
        self.show_sort_seat = True
        self.show_sort_info = True
        self.show_reset = True

        self.load_config()
        self.setup_logger()
        logging.info("=== 班级座位分配系统 启动 ===")

        ctk.set_appearance_mode(self.appearance_mode)
        ctk.set_default_color_theme("blue")

        # ==================== 整体布局 ====================
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(1, weight=1)

        self.build_sidebar()

        self.content_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        self.content_frame.grid(row=0, column=1, sticky="nsew")

        # 初始化三大主页面
        self.main_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        self.settings_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        self.help_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")

        self.build_main_ui()
        self.build_settings_ui()
        self.build_help_ui()

        # 初始化 Loading 动画遮罩层
        self.build_loading_overlay()

        # 默认选中主页
        self.select_frame_by_name("home")
        self.update_treeview_style()

        # 绑定快捷键
        self.bind_shortcuts()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ==================== 核心滚动引擎 ====================
    def apply_y_smooth(self, frame):
        frame._parent_canvas.configure(yscrollincrement=1)

        def _custom_y(event):
            delta = -event.delta if sys.platform in ["darwin", "apple"] else -int(event.delta / 120)
            if delta > 3:
                delta = 3
            elif delta < -3:
                delta = -3
            frame._parent_canvas.yview_scroll(int(delta * int(self.scroll_speed)), "units")

        frame._mouse_wheel_all = _custom_y

    # ==========================================================

    def bind_shortcuts(self):
        prev_combos = [
            (self._prev_mod_undo, self._prev_key_undo),
            (self._prev_mod_export, self._prev_key_export),
            (self._prev_mod_shuffle, self._prev_key_shuffle)
        ]
        for mod, key in prev_combos:
            if mod and key:
                try:
                    self.root.unbind(f'<{mod}-{key}>')
                    if mod == "Control":
                        self.root.unbind(f'<Command-{key}>')
                except Exception:
                    pass

        self.root.bind(f'<{self.mod_undo}-{self.key_undo}>', self.undo)
        if self.mod_undo == "Control":
            self.root.bind(f'<Command-{self.key_undo}>', self.undo)

        self.root.bind(f'<{self.mod_export}-{self.key_export}>', self.export_file)
        if self.mod_export == "Control":
            self.root.bind(f'<Command-{self.key_export}>', self.export_file)

        self.root.bind(f'<{self.mod_shuffle}-{self.key_shuffle}>', self.shuffle_seats)
        if self.mod_shuffle == "Control":
            self.root.bind(f'<Command-{self.key_shuffle}>', self.shuffle_seats)

        self._prev_mod_undo = self.mod_undo
        self._prev_key_undo = self.key_undo
        self._prev_mod_export = self.mod_export
        self._prev_key_export = self.key_export
        self._prev_mod_shuffle = self.mod_shuffle
        self._prev_key_shuffle = self.key_shuffle

    def fade_in(self):
        current_alpha = self.root.attributes("-alpha")
        if current_alpha < 1.0:
            current_alpha += 0.06
            self.root.attributes("-alpha", min(current_alpha, 1.0))
            self.root.after(15, self.fade_in)

    def setup_logger(self):
        if not os.path.exists(self.log_path):
            os.makedirs(self.log_path, exist_ok=True)
        log_file = os.path.join(self.log_path, 'seat_shuffler.log')
        logger = logging.getLogger()
        logger.setLevel(logging.INFO)
        for handler in logger.handlers[:]:
            logger.removeHandler(handler)
        fh = logging.FileHandler(log_file, encoding='utf-8')
        fh.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        logger.addHandler(fh)

    def build_sidebar(self):
        self.sidebar_frame = ctk.CTkFrame(self.root, width=228, corner_radius=0, fg_color=PALETTE["sidebar"])
        self.sidebar_frame.grid(row=0, column=0, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(4, weight=1)

        logo_frame = ctk.CTkFrame(self.sidebar_frame, fg_color="transparent")
        logo_frame.grid(row=0, column=0, padx=22, pady=(32, 24), sticky="ew")

        icon_label = ctk.CTkLabel(logo_frame, text="🎓", font=("Microsoft YaHei", 32))
        icon_label.pack(side="left", padx=(2, 12))

        text_frame = ctk.CTkFrame(logo_frame, fg_color="transparent")
        text_frame.pack(side="left", fill="y", expand=True)

        ctk.CTkLabel(text_frame, text="排座系统", font=("Microsoft YaHei UI", 20, "bold"),
                     text_color=PALETTE["text"]).pack(anchor="w")
        ctk.CTkLabel(text_frame, text="Smart Seating", font=("Arial", 11, "bold"),
                     text_color=PALETTE["text_3"]).pack(anchor="w", pady=(2, 0))

        def create_nav_btn(row, text, name):
            btn = ctk.CTkButton(self.sidebar_frame, text=text, corner_radius=9, height=44, border_spacing=12,
                                text_color=PALETTE["text_2"], fg_color="transparent",
                                hover_color=PALETTE["card_soft"], anchor="w",
                                font=("Microsoft YaHei", 14, "bold"), cursor="hand2",
                                command=lambda: self.select_frame_by_name(name))
            btn.grid(row=row, column=0, sticky="ew", padx=18, pady=5)
            return btn

        self.btn_nav_home = create_nav_btn(1, "📊  控制台", "home")
        self.btn_nav_settings = create_nav_btn(2, "⚙️  偏好设置", "settings")
        self.btn_nav_help = create_nav_btn(3, "❓  帮助与关于", "help")

        ctk.CTkLabel(self.sidebar_frame, text="Version 1.1", font=("Arial", 11),
                     text_color=PALETTE["text_3"]).grid(row=5, column=0, pady=22)

    def _set_nav_active(self, btn, active):
        if active:
            btn.configure(fg_color=PALETTE["primary_soft"], text_color=PALETTE["primary"],
                          hover_color=PALETTE["primary_soft"])
        else:
            btn.configure(fg_color="transparent", text_color=PALETTE["text_2"],
                          hover_color=PALETTE["card_soft"])

    def select_frame_by_name(self, name):
        if self.current_frame_name == "settings" and name != "settings":
            if self.has_unsaved_changes():
                res = messagebox.askyesnocancel("未保存提示",
                                                "检测到未保存的更改，是否立即保存？\n\n[是]：保存并离开\n[否]：放弃修改并离开\n[取消]：留在当前页面")
                if res is True:
                    self.save_settings_action(show_msg=False)
                elif res is False:
                    self.restore_ui_vars()
                else:
                    return

        self.current_frame_name = name

        self._set_nav_active(self.btn_nav_home, name == "home")
        self._set_nav_active(self.btn_nav_settings, name == "settings")
        self._set_nav_active(self.btn_nav_help, name == "help")

        if name == "home":
            self.settings_frame.pack_forget()
            self.help_frame.pack_forget()
            self.main_frame.pack(fill="both", expand=True)
            self.update_treeview_style()
        elif name == "settings":
            self.main_frame.pack_forget()
            self.help_frame.pack_forget()
            self.settings_frame.pack(fill="both", expand=True)
        elif name == "help":
            self.main_frame.pack_forget()
            self.settings_frame.pack_forget()
            self.help_frame.pack(fill="both", expand=True)

    # ==================== 构建主界面 ====================
    def build_main_ui(self):
        header = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        header.pack(fill="x", padx=32, pady=(24, 6))

        title_box = ctk.CTkFrame(header, fg_color="transparent")
        title_box.pack(side="left")
        ctk.CTkLabel(title_box, text="主控制台", font=("Microsoft YaHei UI", 26, "bold"),
                     text_color=PALETTE["text"]).pack(anchor="w")
        ctk.CTkLabel(title_box, text="导入名单 · 随机排座 · 一键导出", font=("Microsoft YaHei", 13),
                     text_color=PALETTE["text_2"]).pack(anchor="w", pady=(3, 0))

        self.status_badge = ctk.CTkLabel(header, text="  未导入数据  ", font=("Microsoft YaHei", 12, "bold"),
                                         fg_color=PALETTE["card_soft"], text_color=PALETTE["text_2"],
                                         corner_radius=15, height=30)
        self.status_badge.pack(side="right", pady=(14, 0))

        self.action_card = ctk.CTkFrame(self.main_frame, corner_radius=14, fg_color=PALETTE["card"])
        self.action_card.pack(fill="x", padx=32, pady=(10, 0))

        self.row_core = ctk.CTkFrame(self.action_card, fg_color="transparent")
        self.row_tools = ctk.CTkFrame(self.action_card, fg_color="transparent")

        def create_action_btn(parent, text, color_style, command, large):
            styles = {
                "primary": {"fg": PALETTE["primary"], "hover": PALETTE["primary_hv"], "txt": ("#FFFFFF", "#FFFFFF")},
                "secondary": {"fg": PALETTE["card_soft"], "hover": PALETTE["border"], "txt": PALETTE["text"]},
                "success": {"fg": PALETTE["success"], "hover": PALETTE["success_hv"], "txt": ("#FFFFFF", "#FFFFFF")},
                "danger": {"fg": PALETTE["danger_soft"], "hover": ("#FEE2E2", "#5C2020"), "txt": PALETTE["danger"]}
            }
            c = styles[color_style]

            return ctk.CTkButton(parent, text=text, command=command, cursor="hand2",
                                 font=("Microsoft YaHei", 14 if large else 13, "bold"),
                                 fg_color=c["fg"], hover_color=c["hover"], text_color=c["txt"],
                                 corner_radius=10, height=48 if large else 40)

        self.btn_upload = create_action_btn(self.row_core, "📁 导入名单", "primary", self.upload_file, True)
        self.btn_shuffle = create_action_btn(self.row_core, "🎲 随机打乱", "primary", self.shuffle_seats, True)
        self.btn_export = create_action_btn(self.row_core, "💾 导出表格", "success", self.export_file, True)
        self.btn_download = create_action_btn(self.row_tools, "📥 模板", "secondary", self.download_template, False)
        self.btn_undo = create_action_btn(self.row_tools, "⬅️ 撤销", "secondary", self.undo, False)
        self.btn_sort_seat = create_action_btn(self.row_tools, "🔢 座位排序", "secondary", self.sort_by_seat, False)
        self.btn_sort_info = create_action_btn(self.row_tools, "📝 学号排序", "secondary", self.sort_by_info, False)
        self.btn_reset = create_action_btn(self.row_tools, "🗑️ 重置", "danger", self.reset_data, False)

        self.refresh_button_layout()

        sf = ctk.CTkFrame(self.main_frame, corner_radius=14, fg_color=PALETTE["card"])
        sf.pack(fill="x", padx=32, pady=14)

        grid = ctk.CTkFrame(sf, fg_color="transparent")
        grid.pack(pady=14, padx=22, fill="x")
        grid.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(grid, text="座位范围", font=("Microsoft YaHei", 13, "bold"),
                     text_color=PALETTE["text_2"]).grid(row=0, column=0, sticky="w", pady=7, padx=(0, 16))

        range_box = ctk.CTkFrame(grid, fg_color="transparent")
        range_box.grid(row=0, column=1, sticky="w")
        self.start_seat_entry = ctk.CTkEntry(range_box, width=72, height=34, corner_radius=8, justify="center",
                                             border_color=PALETTE["border"])
        self.start_seat_entry.pack(side="left")
        ctk.CTkLabel(range_box, text="—", font=("Microsoft YaHei", 12),
                     text_color=PALETTE["text_3"]).pack(side="left", padx=8)
        self.end_seat_entry = ctk.CTkEntry(range_box, width=72, height=34, corner_radius=8, justify="center",
                                           border_color=PALETTE["border"])
        self.end_seat_entry.pack(side="left")

        self.is_compact_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(range_box, text="紧凑模式 (无缝补齐)", variable=self.is_compact_var,
                        font=("Microsoft YaHei", 12), text_color=PALETTE["text_2"],
                        border_color=PALETTE["border"], checkbox_width=20, checkbox_height=20,
                        hover_color=PALETTE["primary"]).pack(side="left", padx=(24, 0))

        ctk.CTkLabel(grid, text="排除座位", font=("Microsoft YaHei", 13, "bold"),
                     text_color=PALETTE["text_2"]).grid(row=1, column=0, sticky="w", pady=7, padx=(0, 16))
        self.excluded_seats_entry = ctk.CTkEntry(grid, height=34, corner_radius=8,
                                                 placeholder_text="选填 · 损坏的座位号，多个用逗号分隔，如: 12, 15",
                                                 border_color=PALETTE["border"])
        self.excluded_seats_entry.grid(row=1, column=1, sticky="ew")

        tc = ctk.CTkFrame(self.main_frame, corner_radius=14, fg_color=PALETTE["card"])
        tc.pack(fill="both", expand=True, padx=32, pady=(0, 8))

        self.tree = ttk.Treeview(tc, columns=("班级", "学号", "姓名", "座位号"), show="headings")
        scrollbar = ctk.CTkScrollbar(tc, orientation="vertical", command=self.tree.yview,
                                     button_color=PALETTE["border"], button_hover_color=PALETTE["text_3"])
        self.tree.configure(yscroll=scrollbar.set)

        scrollbar.pack(side="right", fill="y", padx=(0, 16), pady=16)
        self.tree.pack(side="left", fill="both", expand=True, padx=(16, 5), pady=16)

        self.tree.column("班级", width=120, anchor="center", stretch=True)
        self.tree.column("学号", width=180, anchor="center", stretch=True)
        self.tree.column("姓名", width=160, anchor="center", stretch=True)
        self.tree.column("座位号", width=120, anchor="center", stretch=True)

        self.tree.bind('<Double-1>', self.on_tree_double_click)
        self.tree.bind('<Motion>', self._on_tree_hover)
        self.tree.bind('<Leave>', self._on_tree_leave)

        def _tree_scroll(event):
            if sys.platform in ["darwin", "apple"]:
                delta = -event.delta
            else:
                delta = -int(event.delta / 120)
            if delta > 3:
                delta = 3
            elif delta < -3:
                delta = -3
            row_speed = max(1, int(int(self.scroll_speed) / 15))
            self.tree.yview_scroll(int(delta * row_speed), "units")
            return "break"

        self.tree.bind("<MouseWheel>", _tree_scroll)

        status_bar = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        status_bar.pack(fill="x", padx=36, pady=(0, 14))
        self.status_label = ctk.CTkLabel(status_bar, text="未导入数据 — 点击「导入」或将 Excel 文件拖入窗口",
                                         font=("Microsoft YaHei", 12), text_color=PALETTE["text_3"], anchor="w")
        self.status_label.pack(side="left")
        ctk.CTkLabel(status_bar, text="Smart Seat Shuffler v1.0", font=("Arial", 11),
                     text_color=PALETTE["text_3"]).pack(side="right")

        self.refresh_tree_headings()
        self.root.drop_target_register(DND_FILES)
        self.root.dnd_bind('<<Drop>>', self.handle_drop)

    def build_loading_overlay(self):
        self.loading_mask = ctk.CTkFrame(self.root, fg_color=("#D9E2EC", "#0A1120"), corner_radius=0)
        self.loading_card = ctk.CTkFrame(self.loading_mask, corner_radius=20, fg_color=PALETTE["card"],
                                         border_width=1, border_color=PALETTE["border"])
        self.loading_card.place(relx=0.5, rely=0.5, anchor="center")
        self.loading_label = ctk.CTkLabel(self.loading_card, text="正在处理中，请稍候",
                                          font=("Microsoft YaHei", 16, "bold"), text_color=PALETTE["text"])
        self.loading_label.pack(padx=64, pady=(34, 14))
        self.loading_progress = ctk.CTkProgressBar(self.loading_card, mode="indeterminate", width=240, height=8,
                                                   fg_color=PALETTE["border"], progress_color=PALETTE["primary"])
        self.loading_progress.pack(padx=64, pady=(0, 34))

    def show_loading(self, text="正在处理中，请稍候"):
        self._loading_base_text = text
        self.loading_label.configure(text=text)
        self.loading_mask.place(relx=0, rely=0, relwidth=1, relheight=1)
        self.loading_mask.lift()
        self.loading_progress.start()
        self.root.update_idletasks()
        self._animate_loading_dots()

    def _animate_loading_dots(self):
        if not self.loading_mask.winfo_ismapped():
            return
        self._loading_dots = (self._loading_dots + 1) % 4
        self.loading_label.configure(text=f"{self._loading_base_text} {'·' * self._loading_dots}")
        self.root.after(400, self._animate_loading_dots)

    def hide_loading(self):
        self.loading_progress.stop()
        self.loading_mask.place_forget()

    def _on_close(self):
        if self._cache_job is not None:
            try:
                self.root.after_cancel(self._cache_job)
            except Exception:
                pass
            self._cache_job = None
            self._flush_cache_now()
        self.root.destroy()

    def _recheck_errors(self):
        if self.df is not None and not self.df.empty:
            self.df["学号"] = self.df["学号"].fillna("").astype(str)
            self.df["姓名"] = self.df["姓名"].fillna("").astype(str)

            empty_name_mask = (self.df["姓名"].str.strip() == "") | (self.df["姓名"].str.lower() == "nan")
            valid_id_mask = self.df["学号"].str.strip() != ""
            dup_id_mask = self.df.duplicated(subset=["学号"], keep=False) & valid_id_mask

            self.df['_error'] = empty_name_mask | dup_id_mask

    def on_tree_double_click(self, event):
        if self.df is None or self.df.empty: return

        region = self.tree.identify_region(event.x, event.y)
        if region != "cell": return

        item_id = self.tree.identify_row(event.y)
        column_id = self.tree.identify_column(event.x)
        if not item_id or not column_id: return

        col_index = int(column_id.replace('#', '')) - 1
        col_names = ["班级", "学号", "姓名", "座位号"]
        col_name = col_names[col_index]

        x, y, width, height = self.tree.bbox(item_id, column_id)
        current_value = self.tree.item(item_id, 'values')[col_index]

        entry = tk.Entry(self.tree, relief="flat", highlightthickness=2, highlightcolor="#1f538d",
                         font=("Microsoft YaHei", 12))
        entry.place(x=x, y=y, width=width, height=height)
        entry.insert(0, current_value)
        entry.select_range(0, tk.END)
        entry.focus()

        def save_edit(e=None):
            new_value = entry.get()
            try:
                row_index = self.tree.index(item_id)
                actual_df_index = self.df.index[row_index]

                if str(self.df.at[actual_df_index, col_name]) != new_value:
                    self._push_history()
                    self.df.at[actual_df_index, col_name] = new_value
            except Exception:
                pass

            entry.destroy()
            self._recheck_errors()
            self.update_treeview()

        entry.bind('<Return>', save_edit)
        entry.bind('<FocusOut>', lambda e: entry.destroy())

    def refresh_button_layout(self):
        all_btns = [self.btn_download, self.btn_upload, self.btn_export, self.btn_shuffle,
                    self.btn_undo, self.btn_sort_seat, self.btn_sort_info, self.btn_reset]
        for btn in all_btns:
            btn.grid_forget()

        core = [b for b, on in [(self.btn_upload, self.show_upload), (self.btn_shuffle, self.show_shuffle),
                                (self.btn_export, self.show_export)] if on]
        tools = [b for b, on in [(self.btn_download, self.show_download), (self.btn_undo, self.show_undo),
                                 (self.btn_sort_seat, self.show_sort_seat),
                                 (self.btn_sort_info, self.show_sort_info),
                                 (self.btn_reset, self.show_reset)] if on]

        self.row_core.pack_forget()
        self.row_tools.pack_forget()

        if core:
            for i in range(3):
                self.row_core.columnconfigure(i, weight=1, uniform="core")
            for i, btn in enumerate(core):
                btn.grid(row=0, column=i, padx=6, sticky="ew")
        if tools:
            for i in range(5):
                self.row_tools.columnconfigure(i, weight=1, uniform="tools")
            for i, btn in enumerate(tools):
                btn.grid(row=0, column=i, padx=6, sticky="ew")

        if core and tools:
            self.row_core.pack(fill="x", padx=14, pady=(14, 8))
            self.row_tools.pack(fill="x", padx=14, pady=(0, 14))
        elif core:
            self.row_core.pack(fill="x", padx=14, pady=14)
        elif tools:
            self.row_tools.pack(fill="x", padx=14, pady=14)

    def update_treeview_style(self):
        mode = ctk.get_appearance_mode()
        idx = 1 if mode == "Dark" else 0
        style = ttk.Style()
        style.theme_use("default")

        style.configure("Treeview", font=("Microsoft YaHei", 12), rowheight=42, borderwidth=0, relief="flat",
                        background=PALETTE["row_even"][idx], fieldbackground=PALETTE["row_even"][idx],
                        foreground=PALETTE["row_fg"][idx])
        style.configure("Treeview.Heading", font=("Microsoft YaHei UI", 12, "bold"),
                        background=PALETTE["card_soft"][idx], foreground=PALETTE["text_2"][idx],
                        borderwidth=0, padding=12)
        style.map("Treeview",
                  background=[("selected", PALETTE["row_sel"][idx])],
                  foreground=[("selected", PALETTE["row_sel_fg"][idx])])
        self.tree.tag_configure('evenrow', background=PALETTE["row_even"][idx])
        self.tree.tag_configure('oddrow', background=PALETTE["row_odd"][idx])
        self.tree.tag_configure('hoverrow', background=PALETTE["row_hover"][idx])
        self.tree.tag_configure('errorrow', background=PALETTE["row_error"][idx],
                                foreground=PALETTE["row_error_fg"][idx])

    # ==================== 构建设置界面 ====================
    def build_settings_ui(self):
        header = ctk.CTkFrame(self.settings_frame, fg_color="transparent")
        header.pack(fill="x", padx=32, pady=(24, 10))
        title_box = ctk.CTkFrame(header, fg_color="transparent")
        title_box.pack(side="left")
        ctk.CTkLabel(title_box, text="系统偏好设置", font=("Microsoft YaHei UI", 26, "bold"),
                     text_color=PALETTE["text"]).pack(anchor="w")
        ctk.CTkLabel(title_box, text="外观 · 快捷键 · 数据安全 · 导出排版", font=("Microsoft YaHei", 13),
                     text_color=PALETTE["text_2"]).pack(anchor="w", pady=(3, 0))

        action_f = ctk.CTkFrame(header, fg_color="transparent")
        action_f.pack(side="right")
        ctk.CTkButton(action_f, text="恢复默认", command=self.reset_default_settings, width=100, height=36,
                      corner_radius=18,
                      fg_color="transparent", border_width=1.5, border_color=PALETTE["border"],
                      text_color=PALETTE["text_2"],
                      hover_color=PALETTE["card_soft"], font=("Microsoft YaHei", 13, "bold"), cursor="hand2").pack(
            side="left", padx=10)
        ctk.CTkButton(action_f, text="保存设置", command=self.save_settings_action, width=120, height=36,
                      corner_radius=18,
                      fg_color=PALETTE["primary"], hover_color=PALETTE["primary_hv"],
                      font=("Microsoft YaHei", 13, "bold"),
                      cursor="hand2").pack(side="left")

        self.settings_scroll_frame = ctk.CTkScrollableFrame(self.settings_frame, fg_color="transparent")
        self.settings_scroll_frame.pack(fill="both", expand=True, padx=(30, 38), pady=5)
        self.apply_y_smooth(self.settings_scroll_frame)

        self.backup_var = ctk.BooleanVar(value=self.auto_backup)
        self.path_var = tk.StringVar(value=self.backup_path)
        self.cache_path_var = tk.StringVar(value=self.cache_path)
        self.log_path_var = tk.StringVar(value=self.log_path)
        self.row_var = tk.StringVar(value=str(self.row_height))
        self.font_var = tk.StringVar(value=str(self.font_size))
        self.align_var = tk.StringVar(value=self.align_mode)

        self.scroll_speed_var = tk.StringVar(value=str(self.scroll_speed))

        self.var_mod_undo = tk.StringVar(value=self.mod_undo)
        self.var_key_undo = tk.StringVar(value=self.key_undo.upper())
        self.var_mod_export = tk.StringVar(value=self.mod_export)
        self.var_key_export = tk.StringVar(value=self.key_export.upper())
        self.var_mod_shuffle = tk.StringVar(value=self.mod_shuffle)
        self.var_key_shuffle = tk.StringVar(value=self.key_shuffle.upper())

        self.name_vars = []
        for v in [self.col_name_a, self.col_name_b, self.col_name_c, self.col_name_d, self.col_name_e]:
            self.name_vars.append(tk.StringVar(value=v))

        self.width_vars = []
        for v in [self.col_w_a, self.col_w_b, self.col_w_c, self.col_w_d, self.col_w_e]:
            self.width_vars.append(tk.StringVar(value=str(v)))

        # 【UI 修复】：移除违和的 Emoji，回归极简质感
        mode_map_rev = {"Light": "明亮", "Dark": "暗黑", "System": "跟随系统"}
        self.theme_var = tk.StringVar(value=mode_map_rev.get(self.appearance_mode, "跟随系统"))

        self.var_show_download = ctk.BooleanVar(value=self.show_download)
        self.var_show_upload = ctk.BooleanVar(value=self.show_upload)
        self.var_show_export = ctk.BooleanVar(value=self.show_export)
        self.var_show_shuffle = ctk.BooleanVar(value=self.show_shuffle)
        self.var_show_undo = ctk.BooleanVar(value=self.show_undo)
        self.var_show_sort_seat = ctk.BooleanVar(value=self.show_sort_seat)
        self.var_show_sort_info = ctk.BooleanVar(value=self.show_sort_info)
        self.var_show_reset = ctk.BooleanVar(value=self.show_reset)

        def create_setting_card(title, icon=""):
            card = ctk.CTkFrame(self.settings_scroll_frame, corner_radius=14, fg_color=PALETTE["card"],
                                border_width=1, border_color=PALETTE["border"])
            card.pack(fill="x", pady=(0, 18))
            ctk.CTkLabel(card, text=f"{icon} {title}", font=("Microsoft YaHei", 15, "bold"),
                         text_color=PALETTE["text"]).pack(anchor="w", padx=25, pady=(16, 10))
            return card

        # 1. 外观主题与交互 (恢复自然流式排版，不再死板)
        ui_card = create_setting_card("外观与交互设置", "🎨")

        r_ui_1 = ctk.CTkFrame(ui_card, fg_color="transparent")
        r_ui_1.pack(fill="x", padx=35, pady=(5, 10))
        ctk.CTkLabel(r_ui_1, text="色彩模式:", font=("Microsoft YaHei", 13, "bold")).pack(side="left", padx=(0, 15))

        def on_theme_change(choice):
            mode_map = {"明亮": "Light", "暗黑": "Dark", "跟随系统": "System"}
            ctk.set_appearance_mode(mode_map.get(choice, "System"))
            self.update_treeview_style()

        ctk.CTkSegmentedButton(r_ui_1, values=["明亮", "跟随系统", "暗黑"],
                               variable=self.theme_var, command=on_theme_change,
                               font=("Microsoft YaHei", 13, "bold"),
                               selected_color=PALETTE["primary"],
                               selected_hover_color=PALETTE["primary_hv"]).pack(side="left")

        r_ui_2 = ctk.CTkFrame(ui_card, fg_color="transparent")
        r_ui_2.pack(fill="x", padx=35, pady=(0, 20))
        ctk.CTkLabel(r_ui_2, text="滚动速度:", font=("Microsoft YaHei", 13, "bold")).pack(side="left", padx=(0, 15))
        ctk.CTkComboBox(r_ui_2, values=["10", "20", "30", "50", "80", "120", "150", "200"],
                        variable=self.scroll_speed_var,
                        font=("Microsoft YaHei", 12), width=105, corner_radius=8).pack(side="left")

        # 【UI 修复】：2. 键盘快捷键配置 (废弃生硬网格，采用紧凑自然的排版)
        shortcut_card = create_setting_card("键盘快捷键配置", "⌨️")
        s_f1 = ctk.CTkFrame(shortcut_card, fg_color="transparent")
        s_f1.pack(fill="x", padx=35, pady=(5, 20))

        mods_list = ["Control", "Alt", "Shift"]
        keys_list = [chr(i) for i in range(65, 91)]

        def add_shortcut_row(parent, label_text, mod_var, key_var):
            row = ctk.CTkFrame(parent, fg_color="transparent")
            row.pack(fill="x", pady=6)
            ctk.CTkLabel(row, text=label_text, font=("Microsoft YaHei", 13, "bold"), width=110, anchor="w").pack(
                side="left")
            ctk.CTkComboBox(row, values=mods_list, variable=mod_var, width=105, font=("Microsoft YaHei", 12),
                            corner_radius=8).pack(side="left", padx=(0, 5))
            ctk.CTkLabel(row, text="+", font=("Microsoft YaHei", 16, "bold")).pack(side="left", padx=5)
            ctk.CTkComboBox(row, values=keys_list, variable=key_var, width=70, font=("Microsoft YaHei", 12),
                            corner_radius=8).pack(side="left", padx=5)

        add_shortcut_row(s_f1, "撤销 (Undo):", self.var_mod_undo, self.var_key_undo)
        add_shortcut_row(s_f1, "导出 (Export):", self.var_mod_export, self.var_key_export)
        add_shortcut_row(s_f1, "打乱 (Shuffle):", self.var_mod_shuffle, self.var_key_shuffle)

        # 3. 界面功能开关
        toggle_card = create_setting_card("主界面控件显示", "🛠️")
        t_grid = ctk.CTkFrame(toggle_card, fg_color="transparent")
        t_grid.pack(fill="x", padx=35, pady=(5, 20))

        toggles = [
            ("下载模板", self.var_show_download), ("导入文件", self.var_show_upload),
            ("导出表格", self.var_show_export), ("随机打乱", self.var_show_shuffle),
            ("撤销功能", self.var_show_undo), ("按座位排序", self.var_show_sort_seat),
            ("按学号排序", self.var_show_sort_info), ("重置数据", self.var_show_reset)
        ]

        for c in range(4):
            t_grid.grid_columnconfigure(c, weight=1, uniform="sw")

        for idx, (text, var) in enumerate(toggles):
            r = idx // 4
            c = idx % 4
            sw = ctk.CTkSwitch(t_grid, text=text, variable=var, font=("Microsoft YaHei", 13, "bold"),
                               switch_width=38, switch_height=20)
            sw.grid(row=r, column=c, padx=8, pady=12, sticky="w")

        # 4. 数据与目录配置
        data_card = create_setting_card("数据与目录配置", "🗂️")

        ctk.CTkSwitch(data_card, text="开启导出时自动备份", variable=self.backup_var,
                      font=("Microsoft YaHei", 13, "bold"), switch_width=38, switch_height=20).pack(anchor="w", padx=35,
                                                                                                    pady=(5, 10))

        path_grid = ctk.CTkFrame(data_card, fg_color="transparent")
        path_grid.pack(fill="x", padx=35, pady=(0, 20))

        def add_path_row(parent, row_idx, label, var, open_cmd, clear_cmd, clear_text):
            parent.grid_columnconfigure(1, weight=1)
            ctk.CTkLabel(parent, text=label, width=70, anchor="w", font=("Microsoft YaHei", 13, "bold")).grid(
                row=row_idx, column=0, pady=8, sticky="w")
            ctk.CTkEntry(parent, textvariable=var, state="readonly", font=("Microsoft YaHei", 12)).grid(
                row=row_idx, column=1, padx=10, pady=8, sticky="ew")

            ctk.CTkButton(parent, text="浏览", width=60, fg_color=PALETTE["card_soft"], font=("Microsoft YaHei", 12),
                          text_color=PALETTE["text"], hover_color=PALETTE["border"], cursor="hand2",
                          command=lambda: var.set(filedialog.askdirectory() or var.get())).grid(row=row_idx, column=2,
                                                                                                padx=(0, 5), pady=8)

            ctk.CTkButton(parent, text="打开", width=60, fg_color=PALETTE["primary"],
                          hover_color=PALETTE["primary_hv"], font=("Microsoft YaHei", 12),
                          cursor="hand2", command=open_cmd).grid(row=row_idx, column=3, padx=5, pady=8)

            ctk.CTkButton(parent, text=clear_text, width=85, height=32, corner_radius=16, fg_color="transparent",
                          border_width=1.5, border_color=PALETTE["border"], text_color=PALETTE["text_2"],
                          hover_color=PALETTE["danger_soft"],
                          font=("Microsoft YaHei", 12, "bold"), cursor="hand2", command=clear_cmd).grid(row=row_idx,
                                                                                                        column=4,
                                                                                                        padx=5, pady=8)

        add_path_row(path_grid, 0, "备份路径:", self.path_var, self.open_backup_folder, self.clear_backup, "清理备份")
        add_path_row(path_grid, 1, "缓存路径:", self.cache_path_var, self.open_cache_folder, self.clear_cache,
                     "清理缓存")
        add_path_row(path_grid, 2, "日志路径:", self.log_path_var, self.open_log_file, self.clear_log, "清空日志")

        # 5. 导出排版
        format_card = create_setting_card("导出表格定制", "📊")

        row1 = ctk.CTkFrame(format_card, fg_color="transparent")
        row1.pack(fill="x", padx=35, pady=(5, 10))
        ctk.CTkLabel(row1, text="行高:", font=("Microsoft YaHei", 13, "bold")).pack(side="left")
        ctk.CTkEntry(row1, textvariable=self.row_var, width=60, justify="center", font=("Microsoft YaHei", 12)).pack(
            side="left", padx=(8, 25))
        ctk.CTkLabel(row1, text="字号:", font=("Microsoft YaHei", 13, "bold")).pack(side="left")
        ctk.CTkEntry(row1, textvariable=self.font_var, width=60, justify="center", font=("Microsoft YaHei", 12)).pack(
            side="left", padx=(8, 25))
        ctk.CTkLabel(row1, text="对齐:", font=("Microsoft YaHei", 13, "bold")).pack(side="left")
        ctk.CTkComboBox(row1, values=["居中", "靠左", "靠右"], variable=self.align_var, width=100,
                        font=("Microsoft YaHei", 12)).pack(side="left", padx=8)

        row2 = ctk.CTkFrame(format_card, fg_color="transparent")
        row2.pack(fill="x", padx=25, pady=(5, 25))

        left_lbl_f = ctk.CTkFrame(row2, fg_color="transparent")
        left_lbl_f.pack(side="left", padx=(10, 5), pady=(25, 0))
        ctk.CTkLabel(left_lbl_f, text="表头名称", font=("Microsoft YaHei", 13, "bold"), text_color=PALETTE["text_3"],
                     height=30).pack(pady=(0, 5))
        ctk.CTkLabel(left_lbl_f, text="列宽设定", font=("Microsoft YaHei", 13, "bold"), text_color=PALETTE["text_3"],
                     height=30).pack()

        labels = ["列A(序号)", "列B(班级)", "列C(学号)", "列D(姓名)", "列E(座位)"]
        for i, l in enumerate(labels):
            f = ctk.CTkFrame(row2, fg_color="transparent")
            f.pack(side="left", padx=8, expand=True, fill="x")
            ctk.CTkLabel(f, text=l, font=("Microsoft YaHei", 13, "bold"), text_color=PALETTE["text_2"]).pack(
                pady=(0, 5))
            ctk.CTkEntry(f, textvariable=self.name_vars[i], justify="center", font=("Microsoft YaHei", 12),
                         height=30).pack(fill="x", pady=(0, 5))
            ctk.CTkEntry(f, textvariable=self.width_vars[i], justify="center", font=("Microsoft YaHei", 12),
                         height=30).pack(fill="x")

    def save_settings_action(self, show_msg=True):
        try:
            mode_map = {"明亮": "Light", "暗黑": "Dark", "跟随系统": "System"}
            self.appearance_mode = mode_map.get(self.theme_var.get(), "System")

            self.scroll_speed = int(self.scroll_speed_var.get())

            self.mod_undo = self.var_mod_undo.get()
            self.key_undo = self.var_key_undo.get().lower()
            self.mod_export = self.var_mod_export.get()
            self.key_export = self.var_key_export.get().lower()
            self.mod_shuffle = self.var_mod_shuffle.get()
            self.key_shuffle = self.var_key_shuffle.get().lower()
            self.bind_shortcuts()

            self.show_download = self.var_show_download.get()
            self.show_upload = self.var_show_upload.get()
            self.show_export = self.var_show_export.get()
            self.show_shuffle = self.var_show_shuffle.get()
            self.show_undo = self.var_show_undo.get()
            self.show_sort_seat = self.var_show_sort_seat.get()
            self.show_sort_info = self.var_show_sort_info.get()
            self.show_reset = self.var_show_reset.get()

            self.row_height, self.font_size, self.align_mode = float(self.row_var.get()), int(
                self.font_var.get()), self.align_var.get()
            self.col_name_a, self.col_name_b, self.col_name_c, self.col_name_d, self.col_name_e = [v.get() for v in
                                                                                                   self.name_vars]
            self.col_w_a, self.col_w_b, self.col_w_c, self.col_w_d, self.col_w_e = [float(v.get()) for v in
                                                                                    self.width_vars]

            self.auto_backup = self.backup_var.get()
            self.backup_path = self.path_var.get()
            self.cache_path = self.cache_path_var.get()
            self.log_path = self.log_path_var.get()

            self.save_config()
            self.setup_logger()
            self.refresh_tree_headings()
            self.refresh_button_layout()

            logging.info("系统设置已更新。")
            if show_msg:
                messagebox.showinfo("成功", "所有偏好设置已成功保存！")
                self.select_frame_by_name("home")
        except Exception as e:
            logging.error(f"保存设置失败: {e}")
            messagebox.showerror("格式错误", "请检查数字格式是否正确！")

    def has_unsaved_changes(self):
        try:
            mode_map_rev = {"Light": "明亮", "Dark": "暗黑", "System": "跟随系统"}
            if self.theme_var.get() != mode_map_rev.get(self.appearance_mode, "跟随系统"): return True
            if int(self.scroll_speed_var.get()) != self.scroll_speed: return True

            if self.var_mod_undo.get() != self.mod_undo: return True
            if self.var_key_undo.get().lower() != self.key_undo: return True
            if self.var_mod_export.get() != self.mod_export: return True
            if self.var_key_export.get().lower() != self.key_export: return True
            if self.var_mod_shuffle.get() != self.mod_shuffle: return True
            if self.var_key_shuffle.get().lower() != self.key_shuffle: return True

            if self.var_show_download.get() != self.show_download: return True
            if self.var_show_upload.get() != self.show_upload: return True
            if self.var_show_export.get() != self.show_export: return True
            if self.var_show_shuffle.get() != self.show_shuffle: return True
            if self.var_show_undo.get() != self.show_undo: return True
            if self.var_show_sort_seat.get() != self.show_sort_seat: return True
            if self.var_show_sort_info.get() != self.show_sort_info: return True
            if self.var_show_reset.get() != self.show_reset: return True
            if self.backup_var.get() != self.auto_backup: return True
            if self.path_var.get() != self.backup_path: return True
            if self.cache_path_var.get() != self.cache_path: return True
            if self.log_path_var.get() != self.log_path: return True
            if float(self.row_var.get()) != self.row_height: return True
            if int(self.font_var.get()) != self.font_size: return True
            if self.align_var.get() != self.align_mode: return True
            if [v.get() for v in self.name_vars] != [self.col_name_a, self.col_name_b, self.col_name_c, self.col_name_d,
                                                     self.col_name_e]: return True
            if [float(v.get()) for v in self.width_vars] != [self.col_w_a, self.col_w_b, self.col_w_c, self.col_w_d,
                                                             self.col_w_e]: return True
            return False
        except ValueError:
            return True

    def restore_ui_vars(self):
        mode_map_rev = {"Light": "明亮", "Dark": "暗黑", "System": "跟随系统"}
        self.theme_var.set(mode_map_rev.get(self.appearance_mode, "跟随系统"))
        ctk.set_appearance_mode(self.appearance_mode)

        self.scroll_speed_var.set(str(self.scroll_speed))
        self.var_mod_undo.set(self.mod_undo)
        self.var_key_undo.set(self.key_undo.upper())
        self.var_mod_export.set(self.mod_export)
        self.var_key_export.set(self.key_export.upper())
        self.var_mod_shuffle.set(self.mod_shuffle)
        self.var_key_shuffle.set(self.key_shuffle.upper())

        self.var_show_download.set(self.show_download)
        self.var_show_upload.set(self.show_upload)
        self.var_show_export.set(self.show_export)
        self.var_show_shuffle.set(self.show_shuffle)
        self.var_show_undo.set(self.show_undo)
        self.var_show_sort_seat.set(self.show_sort_seat)
        self.var_show_sort_info.set(self.show_sort_info)
        self.var_show_reset.set(self.show_reset)

        self.backup_var.set(self.auto_backup)
        self.path_var.set(self.backup_path)
        self.cache_path_var.set(self.cache_path)
        self.log_path_var.set(self.log_path)

        self.row_var.set(str(self.row_height))
        self.font_var.set(str(self.font_size))
        self.align_var.set(self.align_mode)

        self.name_vars[0].set(self.col_name_a)
        self.name_vars[1].set(self.col_name_b)
        self.name_vars[2].set(self.col_name_c)
        self.name_vars[3].set(self.col_name_d)
        self.name_vars[4].set(self.col_name_e)

        self.width_vars[0].set(str(self.col_w_a))
        self.width_vars[1].set(str(self.col_w_b))
        self.width_vars[2].set(str(self.col_w_c))
        self.width_vars[3].set(str(self.col_w_d))
        self.width_vars[4].set(str(self.col_w_e))

    def reset_default_settings(self):
        if messagebox.askyesno("恢复默认", "确定要恢复所有设置到默认状态吗？"):
            self.theme_var.set("跟随系统")
            ctk.set_appearance_mode("System")

            self.scroll_speed_var.set("50")

            self.var_mod_undo.set("Control")
            self.var_key_undo.set("Z")
            self.var_mod_export.set("Control")
            self.var_key_export.set("S")
            self.var_mod_shuffle.set("Control")
            self.var_key_shuffle.set("R")

            for var in [self.var_show_download, self.var_show_upload, self.var_show_export, self.var_show_shuffle,
                        self.var_show_undo, self.var_show_sort_seat, self.var_show_sort_info, self.var_show_reset,
                        self.backup_var]:
                var.set(True)
            self.path_var.set(os.path.join(os.getcwd(), "备份文件夹"))
            self.cache_path_var.set(os.path.join(os.getcwd(), "缓存文件夹"))
            self.log_path_var.set(os.path.join(os.getcwd(), "日志文件夹"))
            self.row_var.set("21.0")
            self.font_var.set("14")
            self.align_var.set("居中")
            defaults_names = ["序号", "班级", "学号", "姓名", "座位号"]
            defaults_widths = ["6.0", "7.5", "14.0", "9.0", "8.0"]
            for i in range(5):
                self.name_vars[i].set(defaults_names[i])
                self.width_vars[i].set(defaults_widths[i])
            self.save_settings_action(show_msg=False)
            messagebox.showinfo("成功", "已成功恢复默认设置！")

    def build_help_ui(self):
        header = ctk.CTkFrame(self.help_frame, fg_color="transparent")
        header.pack(fill="x", padx=32, pady=(24, 10))
        title_box = ctk.CTkFrame(header, fg_color="transparent")
        title_box.pack(side="left")
        ctk.CTkLabel(title_box, text="帮助与说明", font=("Microsoft YaHei UI", 26, "bold"),
                     text_color=PALETTE["text"]).pack(anchor="w")
        ctk.CTkLabel(title_box, text="功能指南 · 个性化定制 · 数据安全", font=("Microsoft YaHei", 13),
                     text_color=PALETTE["text_2"]).pack(anchor="w", pady=(3, 0))

        self.help_scroll_frame = ctk.CTkScrollableFrame(self.help_frame, fg_color="transparent")
        self.help_scroll_frame.pack(fill="both", expand=True, padx=30, pady=(5, 20))

        self.apply_y_smooth(self.help_scroll_frame)

        def create_help_card(title, icon, title_color):
            card = ctk.CTkFrame(self.help_scroll_frame, corner_radius=14, fg_color=PALETTE["card"],
                                border_width=1, border_color=PALETTE["border"])
            card.pack(fill="x", pady=(0, 16), ipadx=8, ipady=8)

            header_frame = ctk.CTkFrame(card, fg_color="transparent")
            header_frame.pack(fill="x", padx=28, pady=(14, 5))

            ctk.CTkLabel(header_frame, text=icon, font=("Microsoft YaHei UI", 22)).pack(side="left", padx=(0, 10),
                                                                                        pady=(3, 0))
            ctk.CTkLabel(header_frame, text=title, font=("Microsoft YaHei UI", 16, "bold"),
                         text_color=title_color).pack(side="left")

            content_frame = ctk.CTkFrame(card, fg_color="transparent")
            content_frame.pack(fill="both", expand=True, padx=28, pady=(0, 14))
            return content_frame

        def add_bullet_text(parent, bold_title, text):
            row = ctk.CTkFrame(parent, fg_color="transparent")
            row.pack(fill="x", pady=6)
            row.grid_columnconfigure(1, weight=1)

            if bold_title:
                title_lbl = ctk.CTkLabel(row, text=f"• {bold_title}：", font=("Microsoft YaHei", 14, "bold"),
                                         text_color=PALETTE["text"])
                title_lbl.grid(row=0, column=0, sticky="nw")
                desc_lbl = ctk.CTkLabel(row, text=text, font=("Microsoft YaHei", 14), text_color=PALETTE["text_2"],
                                        justify="left", wraplength=760)
                desc_lbl.grid(row=0, column=1, sticky="nw", padx=(5, 0))
            else:
                desc_lbl = ctk.CTkLabel(row, text=f"• {text}", font=("Microsoft YaHei", 14),
                                        text_color=PALETTE["text_2"], justify="left", wraplength=820)
                desc_lbl.grid(row=0, column=0, columnspan=2, sticky="nw")

        card1 = create_help_card("核心功能指南", "🚀", "#2563EB")
        add_bullet_text(card1, "准备数据", "点击【下载模板】获取标准格式，或直接将您的 Excel 名单拖拽至软件界面。")
        add_bullet_text(card1, "智能排座",
                        "在【座位起止范围】输入号码（如 1 至 50）。如果机房有坏电脑，可在【排除座位】中填入（如：12, 15）。")
        add_bullet_text(card1, "分配模式", "勾选【紧凑模式】可确保座位连续无空缺；取消勾选则按实际起止号码分散排座。")
        add_bullet_text(card1, "脏数据预警",
                        "系统会自动识别空姓名、重复学号等异常并标红显示，双击任意单元格即可直接修改。")
        add_bullet_text(card1, "导出名单", "点击【导出表格】或使用快捷键，系统会自动按您的排版设置生成美观的 Excel 文件。")

        card2 = create_help_card("高级个性化定制", "🎨", "#059669")
        add_bullet_text(card2, "快捷键支持",
                        "支持在设置中自定义修饰键 (如 Control/Alt/Shift) + 字母键的组合（如撤销、导出、打乱），解放双手。")
        add_bullet_text(card2, "界面极简", "在【偏好设置】中，您可以自由隐藏不需要的按钮，打造专注的工作流。")
        add_bullet_text(card2, "视觉与交互", "支持主题切换，并可自由调节各界面的鼠标滚轮滑动速度。")

        card3 = create_help_card("终极数据安全防线", "🛡️", "#D97706")
        add_bullet_text(card3, "无限撤销", "无论排座结果不满意还是误操作，随时使用快捷键返回上一步，防呆且安心。")
        add_bullet_text(card3, "自动缓存", "您的每一次操作都会在后台自动保存，哪怕意外断电、闪退，数据也绝不丢失。")
        add_bullet_text(card3, "自动备份", "每次成功导出表格，系统都会在备份库中悄悄留存一份快照，做到有迹可循。")

        card4 = create_help_card("关于系统", "ℹ️", "#6B7280")
        add_bullet_text(card4, "当前版本", "Version 1.0 ")
        add_bullet_text(card4, "版权声明", "本项目禁止商用，后续会考虑在GitHub上开源")
        add_bullet_text(card4, "问题反馈", "如发现 bug 或有功能建议，欢迎联系开发者邮箱serein1346790@gmail.com ")

    def refresh_tree_headings(self):
        # 【完美对齐数据与表头】：强制让表头文字也完全居中
        self.tree.heading("班级", text=self.col_name_b, anchor=tk.CENTER)
        self.tree.heading("学号", text=self.col_name_c, anchor=tk.CENTER)
        self.tree.heading("姓名", text=self.col_name_d, anchor=tk.CENTER)
        self.tree.heading("座位号", text=self.col_name_e, anchor=tk.CENTER)

    def load_config(self):
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    for key in config:
                        if hasattr(self, key):
                            setattr(self, key, config[key])
            except Exception:
                pass

    def save_config(self):
        config_keys = [
            "appearance_mode", "scroll_speed",
            "mod_undo", "key_undo", "mod_export", "key_export", "mod_shuffle", "key_shuffle",
            "show_download", "show_upload", "show_export",
            "show_shuffle", "show_undo", "show_sort_seat", "show_sort_info", "show_reset",
            "auto_backup", "backup_path", "cache_path", "log_path", "row_height",
            "col_w_a", "col_w_b", "col_w_c", "col_w_d", "col_w_e", "font_size", "align_mode",
            "col_name_a", "col_name_b", "col_name_c", "col_name_d", "col_name_e"
        ]
        config = {k: getattr(self, k) for k in config_keys}
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=4)
        except Exception:
            pass

    def open_log_file(self):
        log_file = os.path.join(self.log_path_var.get(), "seat_shuffler.log")
        if os.path.exists(log_file):
            os.startfile(log_file)
        else:
            messagebox.showinfo("提示", "暂无日志。")

    def open_backup_folder(self):
        if os.path.exists(self.path_var.get()):
            os.startfile(self.path_var.get())
        else:
            messagebox.showwarning("提示", "路径不存在！")

    def open_cache_folder(self):
        if os.path.exists(self.cache_path_var.get()):
            os.startfile(self.cache_path_var.get())
        else:
            messagebox.showwarning("提示", "路径不存在！")

    def clear_backup(self):
        path = self.path_var.get()
        if os.path.exists(path) and messagebox.askyesno("警告", "确定要清理历史备份吗？"):
            count = 0
            for file_name in os.listdir(path):
                if file_name.startswith("备份_") and file_name.endswith(".xlsx"):
                    os.remove(os.path.join(path, file_name))
                    count += 1
            messagebox.showinfo("成功", f"清理了 {count} 个文件！")

    def clear_cache(self):
        cache_file = os.path.join(self.cache_path_var.get(), "实时数据缓存.xlsx")
        if os.path.exists(cache_file):
            try:
                os.remove(cache_file)
                messagebox.showinfo("成功", "缓存已清理！")
            except Exception as e:
                messagebox.showerror("错误", str(e))
        else:
            messagebox.showinfo("提示", "无缓存需清理。")

    def clear_log(self):
        if messagebox.askyesno("确认", "清空所有日志？"):
            logger = logging.getLogger()
            for handler in logger.handlers[:]:
                handler.close()
                logger.removeHandler(handler)
            log_file = os.path.join(self.log_path_var.get(), "seat_shuffler.log")
            if os.path.exists(log_file):
                os.remove(log_file)
            self.setup_logger()
            messagebox.showinfo("成功", "日志已清空！")

    def _push_history(self):
        self.history_stack.append(self.df.copy())

    def save_to_cache(self):
        if self.df is None or self.df.empty:
            return
        if self._cache_job is not None:
            try:
                self.root.after_cancel(self._cache_job)
            except Exception:
                pass
        self._cache_job = self.root.after(CACHE_DEBOUNCE_MS, self._do_save_cache)

    def _do_save_cache(self):
        self._cache_job = None
        self._flush_cache_now()

    def _flush_cache_now(self):
        if self.df is None or self.df.empty:
            return
        try:
            os.makedirs(self.cache_path, exist_ok=True)
            df_cache = self.df.copy().drop(columns=["_error"], errors='ignore')
            df_cache.to_excel(os.path.join(self.cache_path, "实时数据缓存.xlsx"), index=False)
        except Exception:
            pass

    def undo(self, event=None):
        if not self.history_stack:
            messagebox.showinfo("提示", "已撤销到最初状态。")
            return
        self.df = self.history_stack.pop()
        self.current_sort_state = "已撤销恢复"
        self._recheck_errors()
        self.update_treeview()

    def download_template(self):
        path = filedialog.asksaveasfilename(title="保存模板", initialfile="班级导入名单_模板.xlsx",
                                            defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        try:
            cols = [self.col_name_a, self.col_name_b, self.col_name_c, self.col_name_d, self.col_name_e]
            with pd.ExcelWriter(path, engine='openpyxl') as writer:
                pd.DataFrame(columns=cols).to_excel(writer, index=False, sheet_name='座位表')
                ws = writer.sheets['座位表']
                ws.row_dimensions[1].height = self.row_height
                for i, w in enumerate([self.col_w_a, self.col_w_b, self.col_w_c, self.col_w_d, self.col_w_e], 1):
                    ws.column_dimensions[get_column_letter(i)].width = w
            messagebox.showinfo("成功", "模板下载成功！已同步您的排版设置。")
        except Exception as e:
            messagebox.showerror("错误", str(e))

    def handle_drop(self, event):
        path = event.data.strip('{}')
        if path.endswith(('.xlsx', '.xls')):
            self.process_excel_file(path)
        else:
            messagebox.showerror("错误", "仅支持 Excel 文件！")

    def upload_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if path:
            self.process_excel_file(path)

    def process_excel_file(self, path):
        self.show_loading("⌛ 正在读取数据，请稍候...")
        threading.Thread(target=self._process_excel_thread, args=(path,), daemon=True).start()

    def _process_excel_thread(self, path):
        try:
            df = pd.read_excel(path)
            rename_map = {self.col_name_b: "班级", self.col_name_c: "学号", self.col_name_d: "姓名",
                          self.col_name_e: "座位号"}
            df.rename(columns=rename_map, inplace=True)

            if not all(k in df.columns for k in ["班级", "学号", "姓名"]):
                self.root.after(0, self._process_excel_error, "表头不匹配，请下载最新模板！")
                return

            df["座位号"] = df.get("座位号", "")
            df["学号"] = df["学号"].fillna("").astype(str)
            df["姓名"] = df["姓名"].fillna("").astype(str)

            self.root.after(0, self._process_excel_success, df)
        except Exception as e:
            self.root.after(0, self._process_excel_error, str(e))

    def _process_excel_success(self, df):
        self.df = df
        self.history_stack.clear()
        self.current_sort_state = "初始导入"
        self._recheck_errors()
        self.update_treeview()
        self.hide_loading()

        if '_error' in self.df.columns and self.df['_error'].any():
            messagebox.showwarning("数据警告",
                                   "检测到脏数据（学号重复 或 姓名为空），已标红显示。\n您可以直接双击红色单元格进行修改。")
        else:
            messagebox.showinfo("成功", f"成功导入 {len(self.df)} 条数据！")

    def _process_excel_error(self, err_msg):
        self.hide_loading()
        messagebox.showerror("错误", err_msg)

    def shuffle_seats(self, event=None):
        if self.df is None or self.df.empty:
            return
        try:
            total = len(self.df)
            start = int(self.start_seat_entry.get() or 1)
            end = int(self.end_seat_entry.get()) if self.end_seat_entry.get() else None
            exc = {int(x.strip()) for x in self.excluded_seats_entry.get().replace('，', ',').split(',') if x.strip()}
            valid = []
            if self.is_compact_var.get():
                cur = start
                while len(valid) < total:
                    if end and cur > end:
                        messagebox.showerror("错误", "指定区间座位数量不足！")
                        return
                    if cur not in exc:
                        valid.append(cur)
                    cur += 1
            else:
                if not end:
                    messagebox.showerror("错误", "分散模式须填写结束号码！")
                    return
                valid = [s for s in range(start, end + 1) if s not in exc]
                if len(valid) < total:
                    messagebox.showerror("错误", "除去排除的座位后，总座位不足！")
                    return
            self._push_history()
            self.df["座位号"] = random.sample(valid, total)
            self.current_sort_state = "已随机打乱"
            self.update_treeview()
        except Exception:
            messagebox.showerror("错误", "座位参数输入有误！")

    def sort_by_seat(self):
        if self.df is None or self.df.empty:
            return
        self._push_history()
        self.df = self.df.sort_values("座位号", key=lambda s: pd.to_numeric(s, errors="coerce"), na_position="last")
        self.current_sort_state = "按座位排序"
        self.update_treeview()

    def sort_by_info(self):
        if self.df is None or self.df.empty:
            return
        self._push_history()
        self.df = self.df.sort_values(["班级", "学号"])
        self.current_sort_state = "按学号排序"
        self.update_treeview()

    def export_file(self, event=None):
        if self.df is None:
            return
        default_name = f"座位表_{self.current_sort_state}_{datetime.now().strftime('%Y-%m-%d')}.xlsx"
        path = filedialog.asksaveasfilename(initialfile=default_name, defaultextension=".xlsx",
                                            filetypes=[("Excel", "*.xlsx")])
        if not path:
            return

        self.show_loading("⌛ 正在生成排版并导出...")
        threading.Thread(target=self._export_file_thread, args=(path,), daemon=True).start()

    def _export_file_thread(self, path):
        try:
            df_exp = self.df.copy().drop(columns=["序号", self.col_name_a, "_error"], errors='ignore')
            df_exp.insert(0, "序号", range(1, len(df_exp) + 1))
            df_exp.columns = [self.col_name_a, self.col_name_b, self.col_name_c, self.col_name_d, self.col_name_e]

            with pd.ExcelWriter(path, engine='openpyxl') as writer:
                df_exp.to_excel(writer, index=False, sheet_name='座位表')
                ws = writer.sheets['座位表']
                align = Alignment(
                    horizontal={"居中": "center", "靠左": "left", "靠右": "right"}.get(self.align_mode, "center"),
                    vertical='center')
                font = Font(size=self.font_size)

                for i, w in enumerate([self.col_w_a, self.col_w_b, self.col_w_c, self.col_w_d, self.col_w_e], 1):
                    ws.column_dimensions[get_column_letter(i)].width = w

                for row in ws.iter_rows():
                    ws.row_dimensions[row[0].row].height = self.row_height
                    for c in row:
                        c.alignment = align
                        c.font = font

            if self.auto_backup:
                os.makedirs(self.backup_path, exist_ok=True)
                shutil.copy2(path, os.path.join(self.backup_path, f"备份_{os.path.basename(path)}"))

            self.root.after(0, self._export_success)
        except Exception as e:
            self.root.after(0, self._export_error, str(e))

    def _export_success(self):
        self.hide_loading()
        messagebox.showinfo("成功", "数据导出成功！")

    def _export_error(self, err_msg):
        self.hide_loading()
        messagebox.showerror("错误", err_msg)

    def reset_data(self):
        if messagebox.askyesno("确认", "警告：这将清空当前列表的所有数据！是否继续？"):
            self.df = None
            self.current_sort_state = "未导入数据"
            self.history_stack.clear()
            self.update_treeview()

    def refresh_status(self):
        if self.df is None or self.df.empty:
            text, color = "未导入数据 — 点击「导入」或将 Excel 文件拖入窗口", PALETTE["text_3"]
        else:
            err_count = int(self.df["_error"].sum()) if "_error" in self.df.columns else 0
            text = f"共 {len(self.df)} 名学生 · {self.current_sort_state}"
            if err_count:
                text += f" · {err_count} 条数据待修正"
                color = PALETTE["warn"]
            else:
                color = PALETTE["success"]
        self.status_label.configure(text=text, text_color=color)
        self.status_badge.configure(text=f"  {self.current_sort_state}  ")

    def _on_tree_hover(self, event):
        iid = self.tree.identify_row(event.y)
        if iid == self._hover_iid:
            return
        if self._hover_iid and self.tree.exists(self._hover_iid):
            tags = tuple(t for t in self.tree.item(self._hover_iid, "tags") if t != "hoverrow")
            self.tree.item(self._hover_iid, tags=tags)
        self._hover_iid = iid
        if iid:
            self.tree.item(iid, tags=tuple(self.tree.item(iid, "tags")) + ("hoverrow",))

    def _on_tree_leave(self, event):
        if self._hover_iid and self.tree.exists(self._hover_iid):
            tags = tuple(t for t in self.tree.item(self._hover_iid, "tags") if t != "hoverrow")
            self.tree.item(self._hover_iid, tags=tags)
        self._hover_iid = None

    def update_treeview(self):
        children = self.tree.get_children()
        if children:
            self.tree.delete(*children)
        self._hover_iid = None
        if self.df is not None:
            records = self.df.to_dict("records")
            for i, r in enumerate(records):
                tag = "errorrow" if r.get("_error", False) else ("evenrow" if i % 2 == 0 else "oddrow")
                self.tree.insert("", "end", values=(r["班级"], r["学号"], r["姓名"], r["座位号"]), tags=(tag,))
        self.refresh_status()
        self.save_to_cache()


if __name__ == "__main__":
    root = CustomDnDTk()
    app = SeatShufflerApp(root)
    root.mainloop()