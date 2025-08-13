# Combined CMS Tool using CustomTkinter

import customtkinter as ctk
import tkinter as tk # Keep for messagebox and potentially Treeview styling access
from tkinter import ttk, messagebox, filedialog, simpledialog # ttk for Treeview
import requests
import pandas as pd
import threading
import json
import os
from typing import Dict, List, Optional

# --- Cookie Management Window ---
class CookieManagerWindow(ctk.CTkToplevel):
    def __init__(self, parent, current_cookie):
        super().__init__(parent)
        
        self.parent = parent
        self.current_cookie = current_cookie
        
        self.title("Cookie 管理")
        self.geometry("600x400")
        self.transient(parent)
        self.grab_set()
        
        # Center the window
        self.after(100, self.center_window)
        
        self.create_widgets()
        self.load_cookie()
    
    def center_window(self):
        self.update_idletasks()
        x = (self.winfo_screenwidth() // 2) - (self.winfo_width() // 2)
        y = (self.winfo_screenheight() // 2) - (self.winfo_height() // 2)
        self.geometry(f"+{x}+{y}")
    
    def create_widgets(self):
        # Title
        title_label = ctk.CTkLabel(self, text="Cookie 管理", font=ctk.CTkFont(size=24, weight="bold"))
        title_label.pack(pady=(20, 10))
        
        # Main frame
        main_frame = ctk.CTkFrame(self)
        main_frame.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        
        # Cookie input area
        cookie_frame = ctk.CTkFrame(main_frame)
        cookie_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        label = ctk.CTkLabel(cookie_frame, text="全局 Cookie (适用于所有功能):", 
                           font=ctk.CTkFont(size=14, weight="bold"))
        label.pack(anchor="w", padx=10, pady=(10, 5))
        
        # Text area for cookie
        self.cookie_text = ctk.CTkTextbox(cookie_frame, height=200, wrap="word")
        self.cookie_text.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
        # Buttons frame
        button_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        button_frame.pack(fill="x", padx=10, pady=(0, 10))
        
        save_btn = ctk.CTkButton(button_frame, text="保存Cookie", 
                                command=self.save_cookie, height=40,
                                font=ctk.CTkFont(weight="bold"))
        save_btn.pack(side="left", padx=(0, 10))
        
        cancel_btn = ctk.CTkButton(button_frame, text="取消", 
                                  command=self.destroy, height=40,
                                  fg_color="gray50", hover_color="gray30")
        cancel_btn.pack(side="left")
        
        reset_btn = ctk.CTkButton(button_frame, text="重置为默认", 
                                 command=self.reset_cookie, height=40,
                                 fg_color="orange", hover_color="darkorange")
        reset_btn.pack(side="right")
    
    def load_cookie(self):
        self.cookie_text.delete("1.0", "end")
        self.cookie_text.insert("1.0", self.current_cookie)
    
    def save_cookie(self):
        cookie_value = self.cookie_text.get("1.0", "end-1c").strip()
        if cookie_value:
            # Save to file
            self.parent.save_cookie_to_file(cookie_value)
            self.parent.current_cookie = cookie_value
            
            messagebox.showinfo("成功", "Cookie已保存！")
            self.destroy()
        else:
            messagebox.showwarning("警告", "Cookie值不能为空")
    
    def reset_cookie(self):
        if messagebox.askyesno("确认", "确定要重置Cookie为默认值吗？"):
            default_cookie = '_enj="YOUR_DEFAULT_COOKIE_HERE"'
            self.cookie_text.delete("1.0", "end")
            self.cookie_text.insert("1.0", default_cookie)

# --- Main Application Class ---
class UnifiedCMSTool(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("统一 CMS 处理工具")
        self.geometry("1400x800")

        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        self.projects = {
            "极智": "jz", "金胡桃": "kn", "测试": "abc", "宁夏": "nx",
            "福建": "fj", "江西": "jx", "陕西 线上": "sn", "甘肃OTT":"gs_ott",
            "甘肃OIPTV": "gs_iptv", "河南":"ha", "北京": "bj","陕西_银河少儿":"sn_ch"
        }
        
        # Initialize cookie
        self.current_cookie = self.load_cookie_from_file()
        
        self.is_processing = False
        self.file_handlers: Dict[str, Optional[pd.ExcelFile]] = {}
        self.dataframes: Dict[str, Optional[pd.DataFrame]] = {}

        # Menu structure
        self.menu_structure = {
            "总库": {
                "专辑": "master_album",
                "剧集": "master_episode", 
                "介质": "master_media"
            },
            "项目库": {
                "专辑": "project_album",
                "剧集": "project_episode"
            },
            "注入库": {
                "剧头": "inject_header",
                "子集": "inject_subset"
            }
        }

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.create_sidebar()
        self.create_main_content()
        self.create_all_frames()
        
        # Set default selection and expand the category
        self.selected_frame_key = None  # Reset to None first
        # Ensure the category is expanded and frame is shown properly
        self.after(100, lambda: self._initialize_default_view())
        self.appearance_mode_optionemenu.set("System")

    def _initialize_default_view(self):
        """Initialize the default view properly"""
        # Expand the 总库 category
        self.toggle_category("总库")
        # Show the master_album frame
        self.show_frame("master_album")

    def load_cookie_from_file(self):
        """Load cookie from file"""
        cookie_file = "cookie_config.txt"
        default_cookie = '_enj="2|1:0|10:1748223879|4:_enj|36:MjFfZW5qZnVjaGFveWlAZW5qb3ktdHYuY24=|cc3a292d75d6cd6d27c64bcb1dcb4fe14e4c19b903a457d650619c8a1ea85e3c"'
        
        try:
            if os.path.exists(cookie_file):
                with open(cookie_file, 'r', encoding='utf-8') as f:
                    cookie = f.read().strip()
                    return cookie if cookie else default_cookie
        except Exception as e:
            print(f"Error loading cookie: {e}")
        
        return default_cookie

    def save_cookie_to_file(self, cookie):
        """Save cookie to file"""
        cookie_file = "cookie_config.txt"
        try:
            with open(cookie_file, 'w', encoding='utf-8') as f:
                f.write(cookie)
        except Exception as e:
            print(f"Error saving cookie: {e}")
            messagebox.showerror("错误", f"保存Cookie失败: {e}")

    def create_sidebar(self):
        self.sidebar_frame = ctk.CTkFrame(self, width=250, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, rowspan=4, sticky="nsew")
        # Remove the weight from row 10 to prevent centering, buttons will align to top

        self.logo_label = ctk.CTkLabel(self.sidebar_frame, text="CMS 工具箱",
                                       font=ctk.CTkFont(size=24, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(30, 25))

        # Create expandable menu
        self.sidebar_buttons: Dict[str, ctk.CTkButton] = {}
        self.expanded_categories = set()  # Track which categories are expanded
        
        row_counter = 1
        
        for category, subcategories in self.menu_structure.items():
            # Category header button
            category_btn = ctk.CTkButton(
                self.sidebar_frame,
                text=f"▶ {category}",
                command=lambda cat=category: self.toggle_category(cat),
                                   height=40,
                                   corner_radius=8,
                font=ctk.CTkFont(weight="bold", size=14),
                fg_color="gray30",
                hover_color="gray20"
            )
            category_btn.grid(row=row_counter, column=0, padx=20, pady=3, sticky="ew")
            self.sidebar_buttons[f"category_{category}"] = category_btn
            row_counter += 1
            
            # Subcategory buttons (initially hidden)
            for subcat_name, frame_key in subcategories.items():
                subcat_btn = ctk.CTkButton(
                    self.sidebar_frame,
                    text=f"  {subcat_name}",
                    command=lambda k=frame_key: self.show_frame(k),
                    height=32,
                    corner_radius=6,
                    font=ctk.CTkFont(weight="normal"),
                    fg_color="transparent",
                    text_color=("gray10", "gray90")
                )
                subcat_btn.grid(row=row_counter, column=0, padx=(40, 20), pady=1, sticky="ew")
                subcat_btn.grid_remove()  # Hide initially
                self.sidebar_buttons[frame_key] = subcat_btn
                row_counter += 1

        # Add weight to push bottom elements down and keep menu buttons at top
        self.sidebar_frame.grid_rowconfigure(row_counter + 1, weight=1)
        
        # Calculate bottom element positions (use high row numbers to ensure they're at bottom)
        bottom_start_row = row_counter + 50  # Leave plenty of space for expansion
        
        # Cookie management button - positioned at bottom
        self.cookie_button = ctk.CTkButton(
            self.sidebar_frame, 
            text="Cookie 管理",
            command=self.open_cookie_manager,
                                            height=40,
                                            corner_radius=8,
                                            font=ctk.CTkFont(weight="bold"),
            fg_color="orange",
            hover_color="darkorange"
                                            )
        self.cookie_button.grid(row=bottom_start_row, column=0, padx=20, pady=10, sticky="sew")

        # Theme selection - positioned at very bottom
        self.appearance_mode_label = ctk.CTkLabel(self.sidebar_frame, text="主题模式:", anchor="w")
        self.appearance_mode_label.grid(row=bottom_start_row + 1, column=0, padx=20, pady=(25, 5), sticky="sw")
        
        self.appearance_mode_optionemenu = ctk.CTkOptionMenu(
            self.sidebar_frame,
                                                             values=["Light", "Dark", "System"],
            command=self.change_appearance_mode_event
        )
        self.appearance_mode_optionemenu.configure(width=160, height=35, corner_radius=6)
        self.appearance_mode_optionemenu.grid(row=bottom_start_row + 2, column=0, padx=20, pady=(0, 20), sticky="sew")

    def toggle_category(self, category):
        """Toggle category expansion"""
        category_btn = self.sidebar_buttons[f"category_{category}"]
        
        if category in self.expanded_categories:
            # Collapse
            self.expanded_categories.remove(category)
            category_btn.configure(text=f"▶ {category}")
            # Hide subcategory buttons
            for frame_key in self.menu_structure[category].values():
                self.sidebar_buttons[frame_key].grid_remove()
        else:
            # Expand
            self.expanded_categories.add(category)
            category_btn.configure(text=f"▼ {category}")
            # Show subcategory buttons
            for frame_key in self.menu_structure[category].values():
                self.sidebar_buttons[frame_key].grid()

    def create_main_content(self):
        self.main_content_frame = ctk.CTkFrame(self, corner_radius=5, fg_color="transparent")
        self.main_content_frame.grid(row=0, column=1, padx=(20, 20), pady=(20, 20), sticky="nsew")
        self.main_content_frame.grid_rowconfigure(0, weight=1)
        self.main_content_frame.grid_columnconfigure(0, weight=1)

    def create_all_frames(self):
        """Create all frames for different functionalities"""
        self.frames: Dict[str, ctk.CTkFrame] = {}
        
        # Flatten menu structure to get all frame keys
        all_frame_keys = []
        for category, subcategories in self.menu_structure.items():
            all_frame_keys.extend(subcategories.values())
        
        for frame_key in all_frame_keys:
            frame = ctk.CTkFrame(self.main_content_frame, corner_radius=5)
            self.frames[frame_key] = frame
            
            # Create frame content based on type
            if frame_key.startswith("master_"):
                self.create_master_frame(frame, frame_key)
            elif frame_key.startswith("project_"):
                self.create_project_frame(frame, frame_key)
            elif frame_key.startswith("inject_"):
                self.create_inject_frame(frame, frame_key)

    def create_master_frame(self, parent_frame, frame_key):
        """Create master database frames (总库)"""
        parent_frame.grid_columnconfigure(0, weight=1)
        parent_frame.grid_rowconfigure(2, weight=1)

        # Title mapping
        title_map = {
            "master_album": "总库专辑修改",
            "master_episode": "总库剧集修改", 
            "master_media": "总库介质修改"
        }
        
        title_label = ctk.CTkLabel(parent_frame, text=title_map.get(frame_key, "总库修改"), 
                                  font=ctk.CTkFont(size=20, weight="bold"))
        title_label.grid(row=0, column=0, padx=15, pady=(15, 20), sticky="w")

        # File Selection Area
        file_frame = ctk.CTkFrame(parent_frame, corner_radius=10, border_width=1)
        file_frame.grid(row=1, column=0, padx=15, pady=10, sticky="ew")
        file_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(file_frame, text="文件:").grid(row=0, column=0, padx=(15, 5), pady=10, sticky="w")
        
        # Create variables for this frame
        setattr(self, f"{frame_key}_path_var", tk.StringVar())
        setattr(self, f"{frame_key}_sheet_var", tk.StringVar())
        
        path_var = getattr(self, f"{frame_key}_path_var")
        sheet_var = getattr(self, f"{frame_key}_sheet_var")
        
        entry_path = ctk.CTkEntry(file_frame, textvariable=path_var, state='readonly', height=35, corner_radius=6)
        entry_path.grid(row=0, column=1, padx=5, pady=10, sticky="ew")
        
        btn_select_file = ctk.CTkButton(file_frame, text="选择文件", 
                                       command=lambda: self.select_file(frame_key), 
                                       height=35, width=100, corner_radius=6)
        btn_select_file.grid(row=0, column=2, padx=(5,15), pady=10)
        
        ctk.CTkLabel(file_frame, text="Sheet:").grid(row=0, column=3, padx=(20, 5), pady=10, sticky="w")
        
        sheet_combo = ctk.CTkComboBox(file_frame, variable=sheet_var, state="disabled", width=200, height=35, 
                                     corner_radius=6, command=lambda choice: self.load_preview(frame_key, choice))
        sheet_combo.grid(row=0, column=4, padx=(5,15), pady=10, sticky="w")
        setattr(self, f"{frame_key}_sheet_combo", sheet_combo)

        # Preview Area - 总库介质使用文本预览，其他使用表格预览
        self.create_preview_area(parent_frame, frame_key)
        
        # Controls and Progress
        self.create_controls_and_progress(parent_frame, frame_key)

    def create_project_frame(self, parent_frame, frame_key):
        """Create project database frames (项目库)"""
        parent_frame.grid_columnconfigure(0, weight=1)
        parent_frame.grid_rowconfigure(2, weight=1)

        # Title mapping
        title_map = {
            "project_album": "项目库专辑修改",
            "project_episode": "项目库剧集修改"
        }
        
        title_label = ctk.CTkLabel(parent_frame, text=title_map.get(frame_key, "项目库修改"), 
                                  font=ctk.CTkFont(size=20, weight="bold"))
        title_label.grid(row=0, column=0, padx=15, pady=(15, 20), sticky="w")

        # File, Project and Sheet Selection in one row
        selection_frame = ctk.CTkFrame(parent_frame, corner_radius=10, border_width=1)
        selection_frame.grid(row=1, column=0, padx=15, pady=10, sticky="ew")
        selection_frame.grid_columnconfigure(1, weight=1)

        # Create variables
        setattr(self, f"{frame_key}_path_var", tk.StringVar())
        setattr(self, f"{frame_key}_project_var", tk.StringVar(value="北京"))
        setattr(self, f"{frame_key}_sheet_var", tk.StringVar())
        
        path_var = getattr(self, f"{frame_key}_path_var")
        project_var = getattr(self, f"{frame_key}_project_var")
        sheet_var = getattr(self, f"{frame_key}_sheet_var")

        # File selection
        ctk.CTkLabel(selection_frame, text="文件:").grid(row=0, column=0, padx=(15, 5), pady=10, sticky="w")
        entry_path = ctk.CTkEntry(selection_frame, textvariable=path_var, state='readonly', height=35, corner_radius=6)
        entry_path.grid(row=0, column=1, padx=5, pady=10, sticky="ew")
        
        btn_select_file = ctk.CTkButton(selection_frame, text="选择文件", 
                                       command=lambda: self.select_file(frame_key), 
                                       height=35, width=100, corner_radius=6)
        btn_select_file.grid(row=0, column=2, padx=(5,15), pady=10)
        
        # Project selection
        ctk.CTkLabel(selection_frame, text="项目:").grid(row=0, column=3, padx=(20, 5), pady=10, sticky="w")
        project_combo = ctk.CTkComboBox(selection_frame, variable=project_var, values=list(self.projects.keys()), 
                                       width=150, height=35, corner_radius=6)
        project_combo.grid(row=0, column=4, padx=5, pady=10, sticky="w")
        
        # Sheet selection
        ctk.CTkLabel(selection_frame, text="Sheet:").grid(row=0, column=5, padx=(20, 5), pady=10, sticky="w")
        sheet_combo = ctk.CTkComboBox(selection_frame, variable=sheet_var, state="disabled", width=150, height=35, 
                                     corner_radius=6, command=lambda choice: self.load_preview(frame_key, choice))
        sheet_combo.grid(row=0, column=6, padx=(5,15), pady=10, sticky="w")
        setattr(self, f"{frame_key}_sheet_combo", sheet_combo)

        # Preview Area
        self.create_preview_area(parent_frame, frame_key)
        
        # Controls and Progress
        self.create_controls_and_progress(parent_frame, frame_key)

    def create_inject_frame(self, parent_frame, frame_key):
        """Create inject database frames (注入库)"""
        parent_frame.grid_columnconfigure(0, weight=1)
        parent_frame.grid_rowconfigure(2, weight=1)

        # Title mapping
        title_map = {
            "inject_header": "注入库剧头修改",
            "inject_subset": "注入库子集修改"
        }
        
        title_label = ctk.CTkLabel(parent_frame, text=title_map.get(frame_key, "注入库修改"), 
                                  font=ctk.CTkFont(size=20, weight="bold"))
        title_label.grid(row=0, column=0, padx=15, pady=(15, 20), sticky="w")

        # File, Project and Sheet Selection in one row
        selection_frame = ctk.CTkFrame(parent_frame, corner_radius=10, border_width=1)
        selection_frame.grid(row=1, column=0, padx=15, pady=10, sticky="ew")
        selection_frame.grid_columnconfigure(1, weight=1)

        # Create variables
        setattr(self, f"{frame_key}_path_var", tk.StringVar())
        setattr(self, f"{frame_key}_project_var", tk.StringVar(value="北京"))
        setattr(self, f"{frame_key}_sheet_var", tk.StringVar())
        
        path_var = getattr(self, f"{frame_key}_path_var")
        project_var = getattr(self, f"{frame_key}_project_var")
        sheet_var = getattr(self, f"{frame_key}_sheet_var")
        
        # File selection
        ctk.CTkLabel(selection_frame, text="文件:").grid(row=0, column=0, padx=(15, 5), pady=10, sticky="w")
        entry_path = ctk.CTkEntry(selection_frame, textvariable=path_var, state='readonly', height=35, corner_radius=6)
        entry_path.grid(row=0, column=1, padx=5, pady=10, sticky="ew")
        
        btn_select_file = ctk.CTkButton(selection_frame, text="选择文件", 
                                       command=lambda: self.select_file(frame_key), 
                                       height=35, width=100, corner_radius=6)
        btn_select_file.grid(row=0, column=2, padx=(5,15), pady=10)

        # Project selection
        ctk.CTkLabel(selection_frame, text="项目:").grid(row=0, column=3, padx=(20, 5), pady=10, sticky="w")
        project_combo = ctk.CTkComboBox(selection_frame, variable=project_var, values=list(self.projects.keys()), 
                                       width=150, height=35, corner_radius=6,
                                       command=lambda choice: self.on_project_change(frame_key, choice))
        project_combo.grid(row=0, column=4, padx=5, pady=10, sticky="w")
        
        # Sheet selection
        ctk.CTkLabel(selection_frame, text="Sheet:").grid(row=0, column=5, padx=(20, 5), pady=10, sticky="w")
        sheet_combo = ctk.CTkComboBox(selection_frame, variable=sheet_var, state="disabled", width=150, height=35, 
                                     corner_radius=6, command=lambda choice: self.load_preview(frame_key, choice))
        sheet_combo.grid(row=0, column=6, padx=(5,15), pady=10, sticky="w")
        setattr(self, f"{frame_key}_sheet_combo", sheet_combo)

        # Preview Area
        self.create_preview_area(parent_frame, frame_key)
        
        # Controls and Progress
        self.create_controls_and_progress(parent_frame, frame_key)

    def create_preview_area(self, parent_frame, frame_key):
        """Create preview area for a frame"""
        # Use Treeview for all frames to display table data
        preview_wrapper_frame = ctk.CTkFrame(parent_frame, corner_radius=10, border_width=1)
        preview_wrapper_frame.grid(row=2 if frame_key.startswith("master_") else 2, column=0, padx=15, pady=10, sticky="nsew")
        preview_wrapper_frame.grid_columnconfigure(0, weight=1)
        preview_wrapper_frame.grid_rowconfigure(1, weight=1)

        preview_label = ctk.CTkLabel(preview_wrapper_frame, text="数据预览", font=ctk.CTkFont(size=16, weight="bold"))
        preview_label.grid(row=0, column=0, padx=15, pady=(10,5), sticky="w")

        tree_frame = ctk.CTkFrame(preview_wrapper_frame)
        tree_frame.grid(row=1, column=0, padx=15, pady=(5,15), sticky="nsew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        # Configure Treeview style with dark theme to match background
        style = ttk.Style()
        # Dark theme colors to match CustomTkinter dark mode
        style.configure("Treeview", 
                       rowheight=25, 
                       fieldbackground="#2B2B2B",  # Dark background for cells
                       background="#2B2B2B",       # Dark background for tree
                       foreground="#DCE4EE",       # Light text
                       borderwidth=0,
                       relief="flat")
        style.configure("Treeview.Heading", 
                       font=('Segoe UI', 10, 'bold'), 
                       background="#1F1F1F",       # Darker background for headers
                       foreground="#DCE4EE",       # Light text for headers
                       relief="flat",
                       borderwidth=1)
        # Configure selection colors
        style.map("Treeview",
                 background=[('selected', '#1F538D')],  # Blue selection background
                 foreground=[('selected', '#FFFFFF')])  # White text when selected

        # Create Treeview with dark style
        tree = ttk.Treeview(tree_frame, show="headings", height=12)
        
        # Add scrollbars
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')

        setattr(self, f"{frame_key}_tree", tree)

    def create_controls_and_progress(self, parent_frame, frame_key):
        """Create control buttons and progress bars"""
        # Button text mapping
        button_text_map = {
            "master_album": "开始修改 (总库专辑)",
            "master_episode": "开始修改 (总库剧集)",
            "master_media": "开始修改 (总库介质)",
            "project_album": "开始修改 (项目库专辑)",
            "project_episode": "开始修改 (项目库剧集)",
            "inject_header": "开始修改 (注入库剧头)",
            "inject_subset": "开始修改 (注入库子集)"
        }

        # Controls Area
        control_frame = ctk.CTkFrame(parent_frame, fg_color="transparent")
        control_frame.grid(row=3, column=0, padx=15, pady=(10, 5), sticky="ew")
        control_frame.grid_columnconfigure(0, weight=1)

        modify_button = ctk.CTkButton(control_frame, text=button_text_map.get(frame_key, "开始修改"),
                                     command=lambda: self.start_processing(frame_key),
                                     state='disabled', height=40, width=220,
                                     font=ctk.CTkFont(weight="bold", size=14), corner_radius=8)
        modify_button.grid(row=0, column=0, pady=10)
        setattr(self, f"{frame_key}_modify_button", modify_button)

        # Progress Area
        progress_frame = ctk.CTkFrame(parent_frame, fg_color="transparent")
        progress_frame.grid(row=4, column=0, padx=15, pady=(5, 15), sticky="ew")
        progress_frame.grid_columnconfigure(0, weight=1)

        progress_var = tk.DoubleVar(value=0.0)
        progress_bar = ctk.CTkProgressBar(progress_frame, variable=progress_var, height=18, corner_radius=8)
        progress_bar.grid(row=0, column=0, padx=(0,10), sticky="ew")

        progress_label = ctk.CTkLabel(progress_frame, text="0%", width=45, font=ctk.CTkFont(size=12))
        progress_label.grid(row=0, column=1, padx=(0,5))

        setattr(self, f"{frame_key}_progress_var", progress_var)
        setattr(self, f"{frame_key}_progress", progress_bar)
        setattr(self, f"{frame_key}_progress_label", progress_label)

    # --- Helper Methods ---

    def select_file(self, mode):
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if not file_path:
            return

        path_var = getattr(self, f"{mode}_path_var", None)
        if path_var:
            path_var.set(file_path)

        try:
            excel_file = pd.ExcelFile(file_path)
            self.file_handlers[mode] = excel_file
            sheet_names = excel_file.sheet_names

            if not sheet_names:
                messagebox.showerror("错误", "Excel文件不包含任何Sheet")
                self.file_handlers[mode] = None
                self.update_ui_state_after_file_load(mode, success=False)
                return

            # Filter out sheets with "所有sheet" in the name (case insensitive)
            filtered_sheet_names = [name for name in sheet_names if "所有sheet" not in name.lower()]
            
            if not filtered_sheet_names:
                messagebox.showerror("错误", "Excel文件不包含有效的Sheet（排除'所有sheet'后）")
                self.file_handlers[mode] = None
                self.update_ui_state_after_file_load(mode, success=False)
                return

            sheet_combo = getattr(self, f"{mode}_sheet_combo", None)
            if sheet_combo:
                sheet_combo.configure(values=filtered_sheet_names, state="readonly")
                if filtered_sheet_names:
                    # Try to get the active sheet name from the Excel file
                    try:
                        # Get the active sheet name (the one that was last saved as active)
                        active_sheet = excel_file.book.active.title if hasattr(excel_file, 'book') and excel_file.book.active else None
                        if active_sheet and active_sheet in filtered_sheet_names:
                            default_sheet = active_sheet
                        else:
                            default_sheet = filtered_sheet_names[0]
                    except:
                        default_sheet = filtered_sheet_names[0]
                    
                    sheet_combo.set(default_sheet)
                    self.load_preview(mode, default_sheet)

            self.update_ui_state_after_file_load(mode, success=True)
            # Remove success popup - file loading is now silent

        except Exception as e:
            messagebox.showerror("错误", f"加载Excel文件失败: {str(e)}")
            self.file_handlers[mode] = None
            if path_var: path_var.set("")
            self.update_ui_state_after_file_load(mode, success=False)
            self.clear_preview(mode)

    def load_preview(self, mode, sheet_name):
        file_handler = self.file_handlers.get(mode)
        if not file_handler or not sheet_name:
            return

        try:
            df = file_handler.parse(sheet_name)
            
            # Use tree view for all modes
            tree = getattr(self, f"{mode}_tree", None)
            if tree:
                # Clear existing data
                for item in tree.get_children():
                    tree.delete(item)
                
                if not df.empty:
                    # Set up columns
                    columns = list(df.columns)
                    tree["columns"] = columns
                    tree["show"] = "headings"
                    
                    # Configure column headings and widths
                    for col in columns:
                        tree.heading(col, text=col)
                        # Calculate column width based on content
                        try:
                            max_len = max(
                                len(str(col)),  # Header length
                                df[col].astype(str).str.len().max() if not df[col].empty else 0
                            )
                            # Set reasonable width limits
                            width = min(max(max_len * 8 + 20, 80), 300)
                        except:
                            width = 120
                        tree.column(col, width=width, anchor='w')
                    
                    # Insert data (first 100 rows)
                    for index, row in df.head(100).iterrows():
                        values = []
                        for col_name in columns: # Changed 'col' to 'col_name' for clarity, though original 'col' was fine.
                            value = str(row[col_name]) if pd.notna(row[col_name]) else ""
                            values.append(value)
                        tree.insert("", "end", values=values)
                    
                    # Add indicator if there are more rows
                    if len(df) > 100:
                        indicator_values = ["..."] * len(columns)
                        indicator_values[0] = f"... 还有 {len(df) - 100} 行数据"
                        tree.insert("", "end", values=indicator_values)

        except Exception as e:
            messagebox.showerror("错误", f"预览Sheet '{sheet_name}' 失败: {str(e)}")
            self.clear_preview(mode)

    def clear_preview(self, mode):
        tree = getattr(self, f"{mode}_tree", None)
        if tree:
            for item in tree.get_children():
                tree.delete(item)
            tree["columns"] = []

    def update_ui_state_after_file_load(self, mode, success):
        button = getattr(self, f"{mode}_modify_button", None)
        combo = getattr(self, f"{mode}_sheet_combo", None)

        # Check if project is selected for non-master modes
        project_selected = True
        if not mode.startswith("master_"):
            project_var = getattr(self, f"{mode}_project_var", None)
            if project_var:
                project_selected = project_var.get() in self.projects

        state = 'normal' if success and project_selected else 'disabled'

        if button: 
            button.configure(state=state)
        if combo: 
            combo.configure(state='readonly' if success else 'disabled')

        if not success:
            path_var = getattr(self, f"{mode}_path_var", None)
            sheet_var = getattr(self, f"{mode}_sheet_var", None)
            if path_var: path_var.set("")
            if sheet_var: sheet_var.set("")
            if combo: combo.configure(values=[])
            self.clear_preview(mode)

    def open_cookie_manager(self):
        """Open cookie management window"""
        CookieManagerWindow(self, self.current_cookie)

    def get_current_cookie(self, mode):
        """Get cookie for specific mode"""
        return self.current_cookie

    # --- Backend Processing Logic ---

    def start_processing(self, mode):
        if self.is_processing:
            messagebox.showwarning("警告", "当前有任务正在处理中，请稍候。")
            return

        button = getattr(self, f"{mode}_modify_button", None)
        if button: 
            button.configure(state='disabled')

        self.is_processing = True
        thread = threading.Thread(target=self.process_task, args=(mode,), daemon=True)
        thread.start()

    def update_progress(self, mode, value, total):
        progress_var = getattr(self, f"{mode}_progress_var", None)
        label_widget = getattr(self, f"{mode}_progress_label", None)

        if progress_var and label_widget:
            percent = (value / total) * 100 if total > 0 else 0
            progress_var.set(percent / 100.0)
            label_widget.configure(text=f"{int(percent)}%")

    def processing_finished(self, mode, success_count, error_count, total_ops):
        self.is_processing = False
        button = getattr(self, f"{mode}_modify_button", None)
        
        # Check if should enable button
        file_handler = self.file_handlers.get(mode)
        should_enable = file_handler is not None
        
        # For non-master modes, also check project selection
        if not mode.startswith("master_"):
            project_var = getattr(self, f"{mode}_project_var", None)
            if project_var:
                should_enable = should_enable and (project_var.get() in self.projects)

        if button: 
            button.configure(state='normal' if should_enable else 'disabled')

        self.update_progress(mode, 0, 1)

        # Mode text mapping
        mode_text_map = {
            "master_album": "总库专辑修改",
            "master_episode": "总库剧集修改",
            "master_media": "总库介质修改",
            "project_album": "项目库专辑修改",
            "project_episode": "项目库剧集修改",
            "inject_header": "注入库剧头修改",
            "inject_subset": "注入库子集修改"
        }
        
        message = f"{mode_text_map.get(mode, '处理')}完成！\n总计：{total_ops}条\n成功：{success_count}条\n失败：{error_count}条"
        if error_count > 0:
            message += f"\n\n请检查控制台输出获取详细错误信息。"

        messagebox.showinfo("完成", message)

    def process_task(self, mode):
        success_count = 0
        error_count = 0
        total_operations = 0
        session = None
        log_entries = []

        try:
            session = requests.Session()
            current_cookie = self.get_current_cookie(mode)
            
                        # Set headers based on mode type - 为GET请求设置合适的请求头
            if mode.startswith("master_"):
                # 总库修改使用GET请求的请求头
                headers = {
                    'Accept': 'application/json, text/plain, */*',
                    'Accept-Encoding': 'gzip, deflate', 
                    'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6,zh-TW;q=0.5',
                    'Cache-Control': 'no-cache',
                    'Connection': 'keep-alive',
                    'Cookie': current_cookie,
                    'Host': 'cms.enjoy-tv.cn',
                    'Pragma': 'no-cache',
                    'Referer': 'http://cms.enjoy-tv.cn/',
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/138.0.0.0 Safari/537.36 Edg/138.0.0.0'
                }
            else:
                headers = {
                    'Accept': 'application/json, text/plain, */*',
                    'Content-Type': 'application/json',
                    'Cookie': current_cookie,
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/133.0.0.0 Safari/537.36 Edg/133.0.0.0',
                    'Host': 'cms.enjoy-tv.cn',
                    'Origin': 'http://cms.enjoy-tv.cn',
                    'Referer': 'http://cms.enjoy-tv.cn/',
                }
            
            session.headers.update(headers)

            # Process based on mode
            if mode == "master_media":
                success_count, error_count, total_operations = self.process_master_media(session, mode, log_entries)
            elif mode in ["master_album", "master_episode"]:
                success_count, error_count, total_operations = self.process_master_content(session, mode, log_entries)
            elif mode in ["project_album", "project_episode"]:
                success_count, error_count, total_operations = self.process_project_content(session, mode, log_entries)
            elif mode in ["inject_header", "inject_subset"]:
                success_count, error_count, total_operations = self.process_inject_content(session, mode, log_entries)
            else:
                success_count, error_count, total_operations = 0, 0, 0
                log_entries.append(f"未知的处理模式: {mode}")

        except Exception as e:
            log_entries.append(f"线程处理 ({mode}) 顶层异常: {e}")
            self.after(0, lambda: messagebox.showerror("处理异常", f"发生意外错误: {e}\n请查看控制台日志。"))
            success_count, error_count, total_operations = 0, 1, 1
        finally:
            print("\n--- Batch Processing Log ---")
            for entry in log_entries:
                print(entry)
            print("--- End of Log ---\n")
            self.after(0, self.processing_finished, mode, success_count, error_count, total_operations)

    def process_master_media(self, session, mode, log_entries):
        """Process master media modifications"""
        success_count = 0
        error_count = 0
        
        file_handler = self.file_handlers.get(mode)
        sheet_var = getattr(self, f"{mode}_sheet_var", None)
        selected_sheet = sheet_var.get() if sheet_var else None
        
        if not file_handler or not selected_sheet:
            self.after(0, lambda: messagebox.showwarning("警告", "请选择文件和Sheet"))
            log_entries.append("总库介质修改: 未选择文件或Sheet")
            return success_count, error_count, 0

        df = file_handler.parse(selected_sheet)
        required_cols = ['cid', 'vid', 'rate']
        if not all(col in df.columns for col in required_cols):
            missing = [col for col in required_cols if col not in df.columns]
            self.after(0, lambda: messagebox.showerror("错误", f"Sheet '{selected_sheet}' 缺少核心列: {', '.join(missing)}"))
            log_entries.append(f"总库介质修改: Sheet '{selected_sheet}' 缺少核心列: {', '.join(missing)}")
            return success_count, error_count, 0
            
        total_operations = len(df)
        url = "http://cms.enjoy-tv.cn/api/media/edit"

        for i, row in df.iterrows():
            if not self.is_processing: break
            try:
                cid_val = str(row['cid']).strip()
                vid_val = str(row['vid']).strip()
                rate_val = str(row['rate']).strip()

                if not cid_val or cid_val.lower() == 'nan' or \
                   not vid_val or vid_val.lower() == 'nan' or \
                   not rate_val or rate_val.lower() == 'nan':
                    log_entries.append(f"[总库介质 Skip] Invalid row (Sheet: {selected_sheet}, Index: {i}): CID/VID/Rate empty or NaN")
                    error_count += 1
                    self.after(0, self.update_progress, mode, i + 1, total_operations)
                    continue

                # 构建基础参数（必填字段）
                params = {'cid': cid_val, 'vid': vid_val, 'rate': rate_val}
                skip_columns = ['cid', 'vid', 'rate']
                
                # 添加所有其他字段作为修改字段
                has_data_to_update = False
                for field_name, value in row.items():
                    if field_name in skip_columns:
                        continue
                    value_str = str(value).strip()
                    if value_str and value_str.lower() != 'nan' and value_str.lower() != 'none':
                        params[field_name] = value_str
                        has_data_to_update = True

                # 如果有要修改的字段，发送请求
                if has_data_to_update:
                    try:
                        response = session.get(url, params=params, timeout=30)
                        log_entries.append(f"总库介质修改 - CID:{cid_val} VID:{vid_val} Status:{response.status_code}")
                        log_entries.append(f"总库介质修改 - CID:{cid_val} VID:{vid_val} 响应: {response.text}")
                        
                        if response.status_code == 200:
                            try:
                                resp_json = response.json()
                                if resp_json.get("code") == "A000000": 
                                    success_count += 1
                                    log_entries.append(f"总库介质修改 - CID:{cid_val} VID:{vid_val} ✅ 成功 - API返回: {resp_json.get('msg', '成功')}")
                                else:
                                    error_count += 1
                                    log_entries.append(f"总库介质修改 - CID:{cid_val} VID:{vid_val} ❌ 失败 - API错误: {resp_json.get('msg', '未知错误')} (代码: {resp_json.get('code', '未知')})")
                            except json.JSONDecodeError:
                                error_count += 1
                                log_entries.append(f"总库介质修改 - CID:{cid_val} VID:{vid_val} ❌ 失败 - 响应格式错误，非JSON格式")
                        else:
                            error_count += 1
                            log_entries.append(f"总库介质修改 - CID:{cid_val} VID:{vid_val} ❌ 失败 - HTTP状态码: {response.status_code}")
                    except Exception as e:
                        error_count += 1
                        log_entries.append(f"总库介质修改 - CID:{cid_val} VID:{vid_val} 异常: {e}")
                else:
                    log_entries.append(f"总库介质修改 - CID:{cid_val} 跳过: 无有效修改字段")

            except Exception as e:
                error_count += 1
                log_entries.append(f"[总库介质异常] Row {i}, CID={row.get('cid')}, Error: {e}")
            self.after(0, self.update_progress, mode, i + 1, total_operations)

        return success_count, error_count, total_operations

    def process_master_content(self, session, mode, log_entries):
        """Process master album/episode modifications"""
        success_count = 0
        error_count = 0
        
        file_handler = self.file_handlers.get(mode)
        sheet_var = getattr(self, f"{mode}_sheet_var", None)
        selected_sheet = sheet_var.get() if sheet_var else None
        
        if not file_handler or not selected_sheet:
            self.after(0, lambda: messagebox.showwarning("警告", "请选择文件和Sheet"))
            log_entries.append(f"总库{'专辑' if 'album' in mode else '剧集'}修改: 未选择文件或Sheet")
            return success_count, error_count, 0

        df = file_handler.parse(selected_sheet)
        if df.empty:
            self.after(0, lambda: messagebox.showwarning("警告", f"Sheet '{selected_sheet}' 为空"))
            log_entries.append(f"总库{'专辑' if 'album' in mode else '剧集'}修改: Sheet '{selected_sheet}' 为空")
            return success_count, error_count, 0
            
        # Check required columns based on mode
        if mode == "master_album":
            # 总库专辑只需要 cid
            if 'cid' not in df.columns and (df.columns[0].lower() != 'cid' if len(df.columns)>0 else True):
                self.after(0, lambda: messagebox.showerror("错误", f"Sheet '{selected_sheet}' 必须包含 'cid' 列"))
                log_entries.append(f"总库专辑修改: Sheet '{selected_sheet}' 缺少CID列")
                return success_count, error_count, 0
        else:  # master_episode
            # 总库剧集需要 cid 和 vid
            required_cols = ['cid', 'vid']
            if not all(col in df.columns for col in required_cols):
                missing = [col for col in required_cols if col not in df.columns]
                self.after(0, lambda: messagebox.showerror("错误", f"Sheet '{selected_sheet}' 缺少必需列: {', '.join(missing)}"))
                log_entries.append(f"总库剧集修改: Sheet '{selected_sheet}' 缺少必需列: {', '.join(missing)}")
                return success_count, error_count, 0

        cid_column_name = 'cid' if 'cid' in df.columns else df.columns[0]
        total_operations = len(df)
        
        # API endpoint based on mode - 只有总库专辑使用master_edit接口
        if mode == "master_album":
            # 总库专辑使用 /api/cover/master_edit 接口
            base_url = "http://cms.enjoy-tv.cn/api/cover/master_edit"
        else:  # master_episode
            # 总库剧集保持原来的 /api/video/edit 接口
            base_url = "http://cms.enjoy-tv.cn/api/video/edit"

        for i, row_series in df.iterrows():
            if not self.is_processing: break
            # 构建基础参数
            if mode == "master_album":
                # 只传cid和要修改的字段
                cid_str = str(row_series[cid_column_name]).strip()
                if not cid_str or cid_str.lower() == 'nan':
                    log_entries.append(f"[总库专辑 Skip] Invalid CID (Index: {i}): {cid_str}")
                    error_count += 1
                    self.after(0, self.update_progress, mode, i + 1, total_operations)
                    continue
                params = {'cid': cid_str}  # 只包含cid
                skip_columns = ['cid']
            else:  # master_episode
                # 只传vid和要修改的字段
                vid_str = str(row_series.get('vid', '')).strip()
                if not vid_str or vid_str.lower() == 'nan':
                    log_entries.append(f"[总库剧集 Skip] Invalid VID (Index: {i}): {vid_str}")
                    error_count += 1
                    self.after(0, self.update_progress, mode, i + 1, total_operations)
                    continue
                params = {'vid': vid_str}  # 只包含vid
                skip_columns = ['vid']
            has_data_to_update = False
            for field_name, value in row_series.items():
                if field_name in skip_columns:
                    continue
                value_str = str(value).strip()
                if value_str and value_str.lower() != 'nan' and value_str.lower() != 'none':
                    clean_field_name = str(field_name).strip()
                    # 只有总库专辑才进行特殊的数据类型处理
                    if mode == "master_album":
                        # 判断是否为单一数字(0,1,2等)，如果是则传入数值，否则传入字符串
                        if value_str.isdigit() and len(value_str) == 1:
                            params[clean_field_name] = int(value_str)
                        else:
                            params[clean_field_name] = value_str
                    else:
                        # 总库剧集保持原来的字符串处理
                        params[clean_field_name] = value_str
                    has_data_to_update = True
            log_entry_prefix = f"总库{'专辑' if 'album' in mode else '剧集'}修改 - "
            if mode == "master_album":
                log_entry_prefix += f"CID:{cid_str}"
            else:
                log_entry_prefix += f"VID:{vid_str}"
            log_entry_prefix += f" (Row:{i+1})"
            if has_data_to_update:
                try:
                    # 只有总库专辑使用手动构建URL的方式
                    if mode == "master_album":
                        # 手动构建URL，避免自动URL编码
                        url_parts = [base_url + "?"]
                        for key, value in params.items():
                            if isinstance(value, int):
                                url_parts.append(f"{key}={value}&")
                            else:
                                url_parts.append(f"{key}={value}&")
                        manual_url = "".join(url_parts).rstrip("&")
                        
                        # 使用完整URL进行GET请求，不使用params参数避免自动编码
                        response = session.get(manual_url, timeout=30)
                        log_entries.append(f"{log_entry_prefix} 请求URL: {manual_url}")
                    else:
                        # 总库剧集使用原来的POST请求方式
                        response = session.post(base_url, json=params, timeout=30)
                        log_entries.append(f"{log_entry_prefix} 请求URL: {base_url}")
                        log_entries.append(f"{log_entry_prefix} 请求数据: {params}")
                    log_entries.append(f"{log_entry_prefix} 状态码: {response.status_code}")
                    log_entries.append(f"{log_entry_prefix} 响应: {response.text}")
                    
                    # 严格按照用户要求：必须返回 "code": "A000000" 才算真正成功
                    if response.status_code == 200:
                        try:
                            resp_json = response.json()
                            if resp_json.get('code') == "A000000":
                                success_count += 1
                                log_entries.append(f"{log_entry_prefix} ✅ 成功 - API返回: {resp_json.get('msg', '成功')}")
                            else:
                                error_count += 1
                                log_entries.append(f"{log_entry_prefix} ❌ 失败 - API错误: {resp_json.get('msg', '未知错误')} (代码: {resp_json.get('code', '未知')})")
                        except json.JSONDecodeError:
                            error_count += 1
                            log_entries.append(f"{log_entry_prefix} ❌ 失败 - 响应格式错误，非JSON格式")
                    else:
                        error_count += 1
                        log_entries.append(f"{log_entry_prefix} ❌ 失败 - HTTP状态码: {response.status_code}")
                except Exception as e:
                    error_count += 1
                    log_entries.append(f"{log_entry_prefix} 异常: {e}")
            else:
                log_entries.append(f"{log_entry_prefix} 跳过: 无有效数据")
            self.after(0, self.update_progress, mode, i + 1, total_operations)

        return success_count, error_count, total_operations

    def process_project_content(self, session, mode, log_entries):
        """Process project album/episode modifications"""
        success_count = 0
        error_count = 0
        
        file_handler = self.file_handlers.get(mode)
        sheet_var = getattr(self, f"{mode}_sheet_var", None)
        project_var = getattr(self, f"{mode}_project_var", None)
        
        selected_sheet = sheet_var.get() if sheet_var else None
        selected_project = project_var.get() if project_var else None
        partner_code = self.projects.get(selected_project) if selected_project else None
        
        if not file_handler or not selected_sheet or not partner_code or selected_project not in self.projects:
            self.after(0, lambda: messagebox.showwarning("警告", "请选择文件、Sheet和项目"))
            log_entries.append(f"项目库{'专辑' if 'album' in mode else '剧集'}修改: 未完整选择文件、Sheet或项目")
            return success_count, error_count, 0

        df = file_handler.parse(selected_sheet)
        if df.empty:
            self.after(0, lambda: messagebox.showerror("错误", "数据为空"))
            log_entries.append(f"项目库{'专辑' if 'album' in mode else '剧集'}修改: 数据为空")
            return success_count, error_count, 0
        
        # Check required columns based on mode
        if mode == "project_album":
            # 项目库专辑: cid + partner_code + 要修改的字段
            if 'cid' not in df.columns:
                self.after(0, lambda: messagebox.showerror("错误", "缺少CID列"))
                log_entries.append(f"项目库专辑修改: 缺少CID列")
                return success_count, error_count, 0
        else:  # project_episode
            # 项目库剧集: cid + vid + partner_code + 要修改的字段
            required_cols = ['cid', 'vid']
            if not all(col in df.columns for col in required_cols):
                missing = [col for col in required_cols if col not in df.columns]
                self.after(0, lambda: messagebox.showerror("错误", f"缺少必需列: {', '.join(missing)}"))
                log_entries.append(f"项目库剧集修改: 缺少必需列: {', '.join(missing)}")
                return success_count, error_count, 0

        total_operations = len(df)
        
        # API endpoint based on mode
        if mode == "project_album":
            api_endpoint = "http://cms.enjoy-tv.cn/api/project/cover/edit"
        else:  # project_episode
            api_endpoint = "http://cms.enjoy-tv.cn/api/project/video/edit"

        for i, row in df.iterrows():
            if not self.is_processing: break
            cid_str = str(row.get('cid',"")).strip()
            if not cid_str or cid_str.lower() == 'nan':
                log_entries.append(f"[项目库{'专辑' if 'album' in mode else '剧集'} Skip] Invalid CID (Index: {i}): {cid_str}")
                error_count += 1
                self.after(0, self.update_progress, mode, i + 1, total_operations)
                continue

            # Build base parameters based on mode
            if mode == "project_album":
                # 项目库专辑: cid + partner_code + 要修改的字段
                data = {"cid": cid_str, "partner_code": partner_code}
                skip_columns = ['cid', 'partner_code']
            else:  # project_episode
                # 项目库剧集: cid + vid + partner_code + 要修改的字段
                vid_str = str(row.get('vid',"")).strip()
                if not vid_str or vid_str.lower() == 'nan':
                    log_entries.append(f"[项目库剧集 Skip] Invalid VID (Index: {i}): {vid_str}")
                    error_count += 1
                    self.after(0, self.update_progress, mode, i + 1, total_operations)
                    continue
                data = {"cid": cid_str, "vid": vid_str, "partner_code": partner_code}
                skip_columns = ['cid', 'vid', 'partner_code']
            
            has_update_data = False
            for column in df.columns:
                if column.lower() not in [col.lower() for col in skip_columns]:
                    value_str = str(row.get(column, "")).strip()
                    if value_str and value_str.lower() != 'nan' and value_str.lower() != 'none':
                        data[column] = value_str
                        has_update_data = True
                    
            log_prefix = f"项目库{'专辑' if 'album' in mode else '剧集'} - CID:{cid_str}"
            if mode == "project_episode":
                log_prefix += f" VID:{vid_str}"
            log_prefix += f" (Row:{i+1})"

            if has_update_data:
                try:
                    response = session.post(api_endpoint, json=data, timeout=30)
                    log_entries.append(f"{log_prefix} 状态码: {response.status_code}")
                    if response.status_code == 200:
                        try:
                            resp_json = response.json()
                            if resp_json.get('code') == "A000000": 
                                success_count += 1
                            else: 
                                error_count += 1
                                log_entries.append(f"{log_prefix} API错误: {resp_json.get('msg')}")
                        except json.JSONDecodeError:
                            error_count += 1
                            log_entries.append(f"{log_prefix} 响应非JSON格式")
                    else:
                        error_count += 1
                        log_entries.append(f"{log_prefix} HTTP错误: {response.status_code}")
                except Exception as e:
                    error_count += 1
                    log_entries.append(f"{log_prefix} 异常: {e}")
            else:
                log_entries.append(f"{log_prefix} 跳过: 除CID外无有效数据")

            self.after(0, self.update_progress, mode, i + 1, total_operations)

        return success_count, error_count, total_operations

    def process_inject_content(self, session, mode, log_entries):
        """Process inject header/subset modifications"""
        success_count = 0
        error_count = 0
        
        file_handler = self.file_handlers.get(mode)
        sheet_var = getattr(self, f"{mode}_sheet_var", None)
        project_var = getattr(self, f"{mode}_project_var", None)
        
        selected_sheet = sheet_var.get() if sheet_var else None
        selected_project = project_var.get() if project_var else None
        partner_code = self.projects.get(selected_project) if selected_project else None
        
        if not file_handler or not selected_sheet or not partner_code or selected_project not in self.projects:
            self.after(0, lambda: messagebox.showwarning("警告", "请选择文件、Sheet和项目"))
            log_entries.append(f"注入库{'剧头' if 'header' in mode else '子集'}修改: 未完整选择文件、Sheet或项目")
            return success_count, error_count, 0

        df = file_handler.parse(selected_sheet)
        if df.empty or 'cid' not in df.columns:
            self.after(0, lambda: messagebox.showerror("错误", "数据为空或缺少CID列"))
            log_entries.append(f"注入库{'剧头' if 'header' in mode else '子集'}修改: 数据为空或缺少CID列")
            return success_count, error_count, 0

        # 检查注入库子集是否有vid列
        if mode == "inject_subset" and 'vid' not in df.columns:
            self.after(0, lambda: messagebox.showerror("错误", "注入库子集缺少VID列"))
            log_entries.append(f"注入库子集修改: 缺少VID列")
            return success_count, error_count, 0

        total_operations = len(df)
        
        # 修正API端点
        if mode == "inject_header":
            api_endpoint = "http://cms.enjoy-tv.cn/api/inject/inject_cover"
        else:  # inject_subset
            api_endpoint = "http://cms.enjoy-tv.cn/api/inject/inject_video"

        for i, row in df.iterrows():
            if not self.is_processing: break
            cid_str = str(row.get('cid',"")).strip()
            if not cid_str or cid_str.lower() == 'nan':
                log_entries.append(f"[注入库{'剧头' if 'header' in mode else '子集'} Skip] Invalid CID (Index: {i}): {cid_str}")
                error_count += 1
                self.after(0, self.update_progress, mode, i + 1, total_operations)
                continue

            # 构建基础参数（必填字段）
            if mode == "inject_header":
                # 注入库剧头: cid + partner_code + task_status
                params = {"cid": cid_str, "partner_code": partner_code, "task_status": 1}
                skip_columns = ['cid', 'partner_code', 'task_status']
            else:  # inject_subset
                # 注入库子集: cid + vid + partner_code + task_status
                vid_str = str(row.get('vid',"")).strip()
                if not vid_str or vid_str.lower() == 'nan':
                    log_entries.append(f"[注入库子集 Skip] Invalid VID (Index: {i}): {vid_str}")
                    error_count += 1
                    self.after(0, self.update_progress, mode, i + 1, total_operations)
                    continue
                params = {"cid": cid_str, "vid": vid_str, "partner_code": partner_code, "task_status": 1}
                skip_columns = ['cid', 'vid', 'partner_code', 'task_status']
            
            # 添加所有其他字段作为修改字段（如果有的话）
            has_update_data = True  # 注入库总是有基础数据要更新
            for column in df.columns:
                if column.lower() not in [col.lower() for col in skip_columns]:
                    value_str = str(row.get(column, "")).strip()
                    if value_str and value_str.lower() != 'nan' and value_str.lower() != 'none':
                        params[column] = value_str
            
            log_prefix = f"注入库{'剧头' if 'header' in mode else '子集'} - CID:{cid_str}"
            if mode == "inject_subset":
                log_prefix += f" VID:{vid_str}"
            log_prefix += f" (Row:{i+1})"

            if has_update_data:
                try:
                    # 修正：使用GET请求和URL参数，而不是POST + JSON
                    response = session.get(api_endpoint, params=params, timeout=30)
                    log_entries.append(f"{log_prefix} 状态码: {response.status_code}")
                    if response.status_code == 200:
                        try:
                            resp_json = response.json()
                            if resp_json.get('code') == "A000000": 
                                success_count += 1
                            else: 
                                error_count += 1
                                log_entries.append(f"{log_prefix} API错误: {resp_json.get('msg')}")
                        except json.JSONDecodeError:
                            error_count += 1
                            log_entries.append(f"{log_prefix} 响应非JSON格式")
                    else: 
                        error_count += 1
                        log_entries.append(f"{log_prefix} HTTP错误: {response.status_code}")
                except Exception as e:
                    error_count += 1
                    log_entries.append(f"{log_prefix} 异常: {e}")
            else:
                log_entries.append(f"{log_prefix} 跳过: 无有效数据")

            self.after(0, self.update_progress, mode, i + 1, total_operations)
        
        return success_count, error_count, total_operations

    def show_frame(self, frame_key):
        # Don't return early if it's the same frame - always ensure it's displayed
        if hasattr(self, 'selected_frame_key') and self.selected_frame_key:
            current_frame_obj = self.frames.get(self.selected_frame_key)
            if current_frame_obj:
                current_frame_obj.grid_forget()

        frame = self.frames[frame_key]
        frame.grid(row=0, column=0, sticky="nsew", padx=0, pady=0)
        frame.tkraise()
        self.selected_frame_key = frame_key

        # Update button states
        for key, button in self.sidebar_buttons.items():
            if key == frame_key:
                button.configure(fg_color=button.cget("hover_color"))
            elif not key.startswith("category_"):
                button.configure(fg_color="transparent")

    def change_appearance_mode_event(self, new_appearance_mode: str):
        ctk.set_appearance_mode(new_appearance_mode)

        # Update Treeview styles based on theme - always use dark style for consistency
        style = ttk.Style()
        
        # Always use dark theme for table to match the overall dark interface
        style.configure("Treeview", 
                       rowheight=25, 
                       fieldbackground="#2B2B2B",  # Dark background for cells
                       background="#2B2B2B",       # Dark background for tree
                       foreground="#DCE4EE",       # Light text
                       borderwidth=0,
                       relief="flat")
        style.configure("Treeview.Heading", 
                       font=('Segoe UI', 10, 'bold'), 
                       background="#1F1F1F",       # Darker background for headers
                       foreground="#DCE4EE",       # Light text for headers
                       relief="flat",
                       borderwidth=1)
        # Configure selection colors
        style.map("Treeview",
                 background=[('selected', '#1F538D')],  # Blue selection background
                 foreground=[('selected', '#FFFFFF')])  # White text when selected

if __name__ == "__main__":
    app = UnifiedCMSTool()
    app.mainloop()