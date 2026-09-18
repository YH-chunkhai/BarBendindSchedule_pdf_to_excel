import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import threading
import glob
import os
import sys
from PIL import Image, ImageTk

# Import core parsing logic from convert_bbs and convert_bbs_with_dimension
from convert_bbs import parse_all_pdfs, export_to_excel
from convert_bbs_with_dimension import parse_all_pdfs_with_dimension

# Application Theme Settings
ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")

class BBSConverterApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Woh Hup — BBS PDF to Excel Converter")
        self.geometry("1280x820")
        self.minsize(1050, 650)

        self.selected_files = []
        self.extracted_data = []
        self.filtered_data = []
        self.tree_item_map = {}
        self.cell_editor = None
        self.editing_context = None

        # Mode Selection: "with_dimension" (Option 1) or "without_dimension" (Option 2)
        self.pdf_mode = ctk.StringVar(value="with_dimension")

        # Load Perfect Steel Logo
        self.logo_img_topbar = None
        self.logo_img_card = None
        logo_path = os.path.join(os.path.dirname(__file__), "perfect_steel_logo.png")
        if os.path.exists(logo_path):
            try:
                pil_img = Image.open(logo_path)
                self.logo_img_topbar = ctk.CTkImage(light_image=pil_img, dark_image=pil_img, size=(78, 38))
                self.logo_img_card = ctk.CTkImage(light_image=pil_img, dark_image=pil_img, size=(130, 64))
            except Exception as e:
                print(f"Error loading logo: {e}")

        # Configure Root Grid Layout
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # Main background container frame
        self.bg_frame = ctk.CTkFrame(self, fg_color=("#F8FAFC", "#0F172A"), corner_radius=0)
        self.bg_frame.grid(row=0, column=0, sticky="nsew")
        self.bg_frame.grid_columnconfigure(0, weight=1)
        self.bg_frame.grid_rowconfigure(0, weight=1)

        # Build the two view frames
        self.build_upload_view()
        self.build_workspace_view()

        # Show initial upload view
        self.show_upload_view()

    def build_upload_view(self):
        """Builds the centered card view matching Image 1 & 2"""
        self.upload_view = ctk.CTkFrame(self.bg_frame, fg_color="transparent")
        self.upload_view.grid_columnconfigure(0, weight=1)
        self.upload_view.grid_rowconfigure(0, weight=1)

        # Centered Soft Card
        self.card = ctk.CTkFrame(
            self.upload_view, 
            width=560, 
            corner_radius=20, 
            fg_color=("#FFFFFF", "#1E293B"),
            border_width=1,
            border_color=("#F1F5F9", "#334155")
        )
        self.card.place(relx=0.5, rely=0.48, anchor="center")

        # Card Title Stack
        if hasattr(self, 'logo_img_card') and self.logo_img_card:
            self.lbl_card_logo = ctk.CTkLabel(self.card, image=self.logo_img_card, text="")
            self.lbl_card_logo.pack(padx=40, pady=(24, 0))
            self.lbl_card_title = ctk.CTkLabel(
                self.card, 
                text="Convert PDF to Excel", 
                font=ctk.CTkFont(family="Segoe UI", size=22, weight="bold"),
                text_color=("#0F172A", "#F8FAFC")
            )
            self.lbl_card_title.pack(padx=40, pady=(8, 4))
        else:
            self.lbl_card_title = ctk.CTkLabel(
                self.card, 
                text="Convert PDF to Excel", 
                font=ctk.CTkFont(family="Segoe UI", size=24, weight="bold"),
                text_color=("#0F172A", "#F8FAFC")
            )
            self.lbl_card_title.pack(padx=40, pady=(36, 6))

        self.lbl_card_subtitle = ctk.CTkLabel(
            self.card, 
            text="Turn trapped data into editable spreadsheets", 
            font=ctk.CTkFont(family="Segoe UI", size=14),
            text_color=("#64748B", "#94A3B8")
        )
        self.lbl_card_subtitle.pack(padx=40, pady=(0, 14))

        # Mode Selection Section Header
        self.lbl_mode_title = ctk.CTkLabel(
            self.card,
            text="CHOOSE PDF FORMAT",
            font=ctk.CTkFont(family="Segoe UI", size=10, weight="bold"),
            text_color=("#94A3B8", "#64748B")
        )
        self.lbl_mode_title.pack(padx=40, anchor="w", pady=(0, 6))

        # 2 Nice Interactive Option Buttons Row
        self.mode_btn_row = ctk.CTkFrame(self.card, fg_color="transparent")
        self.mode_btn_row.pack(padx=40, pady=(0, 16), fill="x")
        self.mode_btn_row.grid_columnconfigure(0, weight=1)
        self.mode_btn_row.grid_columnconfigure(1, weight=1)

        self.btn_mode_dim = ctk.CTkButton(
            self.mode_btn_row,
            text="📄 PDF with Dimension\n(Header ABCD)",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            height=54,
            corner_radius=12,
            command=lambda: self.select_pdf_mode("with_dimension")
        )
        self.btn_mode_dim.grid(row=0, column=0, padx=(0, 6), sticky="ew")

        self.btn_mode_nodim = ctk.CTkButton(
            self.mode_btn_row,
            text="📐 PDF without Dimension\n(Legacy / Shape OCR)",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            height=54,
            corner_radius=12,
            command=lambda: self.select_pdf_mode("without_dimension")
        )
        self.btn_mode_nodim.grid(row=0, column=1, padx=(6, 0), sticky="ew")

        # Initialize button visual active state
        self.select_pdf_mode(self.pdf_mode.get())

        # Dropzone Container Box
        self.dropzone = ctk.CTkFrame(
            self.card, 
            width=480, 
            height=220, 
            corner_radius=16, 
            fg_color=("#F8FAFC", "#0F172A"),
            border_width=1,
            border_color=("#E2E8F0", "#334155")
        )
        self.dropzone.pack(padx=40, pady=(0, 24))
        self.dropzone.pack_propagate(False)

        # --- Dropzone State 1: Unselected (Cloud icon + Browse button) ---
        self.dz_unselected_frame = ctk.CTkFrame(self.dropzone, fg_color="transparent")
        self.dz_unselected_frame.place(relx=0.5, rely=0.5, anchor="center")

        # Cloud Icon Badge
        self.cloud_badge = ctk.CTkFrame(
            self.dz_unselected_frame, 
            width=54, 
            height=54, 
            corner_radius=27, 
            fg_color=("#F1F5F9", "#1E293B")
        )
        self.cloud_badge.pack(pady=(0, 10))
        self.cloud_badge.pack_propagate(False)

        self.cloud_icon = ctk.CTkLabel(
            self.cloud_badge, 
            text="☁️", 
            font=ctk.CTkFont(size=24)
        )
        self.cloud_icon.place(relx=0.5, rely=0.5, anchor="center")

        self.lbl_drag_drop = ctk.CTkLabel(
            self.dz_unselected_frame, 
            text="Drag & Drop your PDF here", 
            font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
            text_color=("#1E293B", "#F1F5F9")
        )
        self.lbl_drag_drop.pack(pady=(0, 2))

        self.lbl_or = ctk.CTkLabel(
            self.dz_unselected_frame, 
            text="or", 
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color=("#94A3B8", "#64748B")
        )
        self.lbl_or.pack(pady=(0, 10))

        # Button row: Browse Files & Select Folder
        self.btn_row = ctk.CTkFrame(self.dz_unselected_frame, fg_color="transparent")
        self.btn_row.pack()

        self.btn_browse = ctk.CTkButton(
            self.btn_row, 
            text="Browse Files", 
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            height=38,
            width=130,
            corner_radius=8,
            fg_color=("#FFFFFF", "#1E293B"),
            hover_color=("#F1F5F9", "#334155"),
            text_color=("#0F172A", "#F8FAFC"),
            border_width=1,
            border_color=("#E2E8F0", "#475569"),
            command=self.select_files
        )
        self.btn_browse.pack(side="left", padx=4)

        self.btn_browse_folder = ctk.CTkButton(
            self.btn_row, 
            text="Browse Folder", 
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            height=38,
            width=130,
            corner_radius=8,
            fg_color=("#FFFFFF", "#1E293B"),
            hover_color=("#F1F5F9", "#334155"),
            text_color=("#0F172A", "#F8FAFC"),
            border_width=1,
            border_color=("#E2E8F0", "#475569"),
            command=self.select_folder
        )
        self.btn_browse_folder.pack(side="left", padx=4)

        # --- Dropzone State 2: File Selected (Green file badge + filename + status) ---
        self.dz_selected_frame = ctk.CTkFrame(self.dropzone, fg_color="transparent")

        # Soft Green File Icon Badge
        self.file_badge = ctk.CTkFrame(
            self.dz_selected_frame, 
            width=54, 
            height=54, 
            corner_radius=27, 
            fg_color=("#DCFCE7", "#064E3B")
        )
        self.file_badge.pack(pady=(0, 10))
        self.file_badge.pack_propagate(False)

        self.file_icon = ctk.CTkLabel(
            self.file_badge, 
            text="📄", 
            font=ctk.CTkFont(size=24)
        )
        self.file_icon.place(relx=0.5, rely=0.5, anchor="center")

        self.lbl_selected_filename = ctk.CTkLabel(
            self.dz_selected_frame, 
            text="filename.pdf", 
            font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
            text_color=("#0F172A", "#F8FAFC")
        )
        self.lbl_selected_filename.pack(pady=(0, 4))

        self.lbl_upload_success = ctk.CTkLabel(
            self.dz_selected_frame, 
            text="✓ Successfully uploaded", 
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            text_color=("#16A34A", "#4ADE80")
        )
        self.lbl_upload_success.pack(pady=(0, 6))

        # Change file / folder button
        self.btn_change_file = ctk.CTkButton(
            self.dz_selected_frame,
            text="Choose different file",
            font=ctk.CTkFont(family="Segoe UI", size=11, underline=True),
            fg_color="transparent",
            hover_color=("#F1F5F9", "#334155"),
            text_color=("#64748B", "#94A3B8"),
            height=20,
            command=self.select_files
        )
        self.btn_change_file.pack()

        # Start Conversion Button (Full Width inside Card)
        self.btn_convert_now = ctk.CTkButton(
            self.card, 
            text="Start to Convert  ➔", 
            font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
            height=48,
            corner_radius=10,
            fg_color="#2563EB",
            hover_color="#1D4ED8",
            text_color="#FFFFFF",
            command=self.start_conversion
        )

        # Status & Progress bar inside card
        self.progress_container = ctk.CTkFrame(self.card, fg_color="transparent")

        self.progress_header = ctk.CTkFrame(self.progress_container, fg_color="transparent")
        self.progress_header.pack(fill="x", padx=40, pady=(0, 6))

        self.status_msg = ctk.CTkLabel(
            self.progress_header,
            text="Extracting BBS Data...",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color=("#2563EB", "#60A5FA")
        )
        self.status_msg.pack(side="left")

        self.lbl_progress_percent = ctk.CTkLabel(
            self.progress_header,
            text="0%",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            text_color=("#2563EB", "#60A5FA")
        )
        self.lbl_progress_percent.pack(side="right")

        self.progressbar = ctk.CTkProgressBar(self.progress_container, height=10, corner_radius=5, progress_color="#2563EB")
        self.progressbar.pack(fill="x", padx=40)
        self.progressbar.set(0)

        # Bottom padding in card
        self.card_padding = ctk.CTkFrame(self.card, height=20, fg_color="transparent")
        self.card_padding.pack(fill="x", side="bottom")

    def build_workspace_view(self):
        """Builds the soft, simple post-extraction workspace view"""
        self.workspace_view = ctk.CTkFrame(self.bg_frame, fg_color="transparent")
        self.workspace_view.grid_columnconfigure(0, weight=1)
        self.workspace_view.grid_rowconfigure(1, weight=1)

        # ------------------ TOP NAVIGATION BAR ------------------
        self.top_bar = ctk.CTkFrame(
            self.workspace_view, 
            corner_radius=0, 
            height=58, 
            fg_color=("#FFFFFF", "#1E293B"),
            border_width=1,
            border_color=("#F1F5F9", "#334155")
        )
        self.top_bar.grid(row=0, column=0, sticky="ew", padx=0, pady=0)
        self.top_bar.grid_columnconfigure(1, weight=1)

        # Left Icon & Title Box
        self.ws_left_box = ctk.CTkFrame(self.top_bar, fg_color="transparent")
        self.ws_left_box.grid(row=0, column=0, padx=20, pady=8, sticky="w")

        if hasattr(self, 'logo_img_topbar') and self.logo_img_topbar:
            self.ws_logo_lbl = ctk.CTkLabel(self.ws_left_box, image=self.logo_img_topbar, text="")
            self.ws_logo_lbl.pack(side="left", padx=(0, 12))
        else:
            self.ws_icon_badge = ctk.CTkFrame(
                self.ws_left_box, 
                width=34, 
                height=34, 
                corner_radius=8, 
                fg_color="#2563EB"
            )
            self.ws_icon_badge.pack(side="left", padx=(0, 10))
            self.ws_icon_badge.pack_propagate(False)

            self.ws_icon_lbl = ctk.CTkLabel(self.ws_icon_badge, text="📊", font=ctk.CTkFont(size=16))
            self.ws_icon_lbl.place(relx=0.5, rely=0.5, anchor="center")

        self.ws_title_stack = ctk.CTkFrame(self.ws_left_box, fg_color="transparent")
        self.ws_title_stack.pack(side="left")

        self.lbl_ws_subtitle = ctk.CTkLabel(
            self.ws_title_stack, 
            text="FILE WORKSPACE", 
            font=ctk.CTkFont(family="Segoe UI", size=9, weight="bold"),
            text_color="#94A3B8"
        )
        self.lbl_ws_subtitle.pack(anchor="w")

        self.lbl_ws_filename = ctk.CTkLabel(
            self.ws_title_stack, 
            text="MVR-TT-H312.pdf", 
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            text_color=("#0F172A", "#F8FAFC")
        )
        self.lbl_ws_filename.pack(anchor="w")

        # Center Search Box (Soft pill)
        self.search_box = ctk.CTkEntry(
            self.top_bar, 
            placeholder_text="🔍  Search Member, Bar Mark, Type, Size, Shape Code...", 
            font=ctk.CTkFont(family="Segoe UI", size=12),
            height=36,
            width=400,
            corner_radius=18,
            fg_color=("#F1F5F9", "#0F172A"),
            border_width=0
        )
        self.search_box.grid(row=0, column=1, padx=20, pady=10, sticky="w")
        self.search_box.bind("<KeyRelease>", self.on_search_change)

        # Right Status & Notification Indicator
        self.ws_right_box = ctk.CTkFrame(self.top_bar, fg_color="transparent")
        self.ws_right_box.grid(row=0, column=2, padx=20, pady=10, sticky="e")

        self.lbl_ws_saved = ctk.CTkLabel(
            self.ws_right_box, 
            text="●  All changes saved locally", 
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color="#16A34A"
        )
        self.lbl_ws_saved.pack(side="left", padx=10)



        # ------------------ CENTER DATA TABLE AREA ------------------
        self.center_frame = ctk.CTkFrame(self.workspace_view, fg_color="transparent")
        self.center_frame.grid(row=1, column=0, sticky="nsew", padx=20, pady=12)
        self.center_frame.grid_columnconfigure(0, weight=1)
        self.center_frame.grid_rowconfigure(2, weight=1)

        # Soft Toast Banner Notification (Replaces ugly native popup box)
        self.toast_banner = ctk.CTkFrame(
            self.center_frame, 
            corner_radius=10, 
            fg_color=("#DCFCE7", "#064E3B"),
            border_width=1,
            border_color=("#86EFAC", "#059669")
        )
        # Hidden by default, grid on post extraction / export
        self.lbl_toast_msg = ctk.CTkLabel(
            self.toast_banner, 
            text="✓ Extraction complete! 18 BBS rows ready for export.", 
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            text_color=("#14532D", "#F0FDF4")
        )
        self.lbl_toast_msg.pack(side="left", padx=16, pady=8)

        # Soft KPI Badges Bar
        self.stats_bar = ctk.CTkFrame(self.center_frame, fg_color="transparent", height=42)
        self.stats_bar.grid(row=1, column=0, sticky="ew", pady=(0, 10))

        self.badge_rows = self.create_metric_badge(self.stats_bar, "TOTAL ROWS", "0 rows", "📊", ("#FFFFFF", "#1E293B"), "#4F46E5")
        self.badge_rows.pack(side="left", padx=(0, 10))

        self.badge_members = self.create_metric_badge(self.stats_bar, "UNIQUE MEMBERS", "0 unique", "🏢", ("#FFFFFF", "#1E293B"), "#2563EB")
        self.badge_members.pack(side="left", padx=(0, 10))

        self.badge_shapes = self.create_metric_badge(self.stats_bar, "SHAPE CODES", "0 types", "📐", ("#FFFFFF", "#1E293B"), "#0D9488")
        self.badge_shapes.pack(side="left", padx=(0, 10))

        # Treeview Data Grid Frame
        self.grid_container = ctk.CTkFrame(
            self.center_frame, 
            corner_radius=12, 
            fg_color=("#FFFFFF", "#1E293B"),
            border_width=1,
            border_color=("#F1F5F9", "#334155")
        )
        self.grid_container.grid(row=2, column=0, sticky="nsew")
        self.grid_container.grid_columnconfigure(0, weight=1)
        self.grid_container.grid_rowconfigure(0, weight=1)

        # Treeview Data Grid Widget
        self.cols = ['Member', 'Bar Mark', 'Type', 'Size', 'No. of MBRS', 'No. of EACH', 'TOTAL', 'SHAPE CODE', 'A', 'B', 'C', 'D', 'E']
        self.tree = ttk.Treeview(self.grid_container, columns=self.cols, show='headings', style="Modern.Treeview")
        
        col_widths = {
            'Member': 240,
            'Bar Mark': 110,
            'Type': 80,
            'Size': 80,
            'No. of MBRS': 105,
            'No. of EACH': 105,
            'TOTAL': 95,
            'SHAPE CODE': 110,
            'A': 90,
            'B': 90,
            'C': 90,
            'D': 90,
            'E': 90
        }
        
        for col in self.cols:
            self.tree.heading(col, text=col)
            w = col_widths.get(col, 90)
            align = "w" if col in ['Member', 'Bar Mark'] else "center"
            self.tree.column(col, width=w, anchor=align)

        # Configure Treeview table styling (soft theme)
        self.update_tree_theme()

        # Bind events for inline cell editing
        self.tree.bind("<Double-1>", self.on_tree_double_click)
        self.tree.bind("<Return>", self.on_tree_double_click)
        self.tree.bind("<F2>", self.on_tree_double_click)
        self.tree.bind("<MouseWheel>", lambda e: self.destroy_cell_editor())

        # Sleek Modern Soft Scrollbars (Replacing retro Windows 95 scrollbars)
        self.tree_scroll_y = ctk.CTkScrollbar(
            self.grid_container, 
            orientation="vertical", 
            command=self.tree.yview,
            button_color="#CBD5E1",
            button_hover_color="#94A3B8",
            fg_color="transparent",
            width=12,
            corner_radius=6
        )
        self.tree_scroll_x = ctk.CTkScrollbar(
            self.grid_container, 
            orientation="horizontal", 
            command=self.tree.xview,
            button_color="#CBD5E1",
            button_hover_color="#94A3B8",
            fg_color="transparent",
            height=12,
            corner_radius=6
        )
        self.tree.configure(yscrollcommand=self.tree_scroll_y.set, xscrollcommand=self.tree_scroll_x.set)

        self.tree.grid(row=0, column=0, sticky="nsew", padx=(12, 4), pady=(12, 4))
        self.tree_scroll_y.grid(row=0, column=1, sticky="ns", pady=12, padx=(2, 8))
        self.tree_scroll_x.grid(row=1, column=0, sticky="ew", padx=12, pady=(2, 8))

        # ------------------ BOTTOM ACTION BAR ------------------
        self.bottom_bar = ctk.CTkFrame(
            self.workspace_view, 
            corner_radius=0, 
            height=66, 
            fg_color=("#FFFFFF", "#1E293B"),
            border_width=1,
            border_color=("#F1F5F9", "#334155")
        )
        self.bottom_bar.grid(row=2, column=0, sticky="ew", padx=0, pady=0)

        # Right Action Buttons
        self.bottom_btn_box = ctk.CTkFrame(self.bottom_bar, fg_color="transparent")
        self.bottom_btn_box.pack(side="right", padx=20, pady=10)

        self.btn_convert_next = ctk.CTkButton(
            self.bottom_btn_box, 
            text="Convert Next", 
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            height=42,
            width=130,
            corner_radius=8,
            fg_color=("#FFFFFF", "#0F172A"),
            hover_color=("#F1F5F9", "#334155"),
            text_color=("#0F172A", "#F8FAFC"),
            border_width=1,
            border_color=("#E2E8F0", "#475569"),
            command=self.reset_all
        )
        self.btn_convert_next.pack(side="left", padx=8)

        self.btn_export = ctk.CTkButton(
            self.bottom_btn_box, 
            text="📥  Convert to Excel", 
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            height=42,
            width=170,
            corner_radius=8,
            fg_color="#059669",
            hover_color="#047857",
            text_color="#FFFFFF",
            command=self.save_excel
        )
        self.btn_export.pack(side="left", padx=8)

    # ------------------ VIEW TOGGLING & LOGIC METHODS ------------------
    def show_upload_view(self):
        if hasattr(self, 'workspace_view'):
            self.workspace_view.grid_forget()
        self.upload_view.grid(row=0, column=0, sticky="nsew")

    def show_workspace_view(self):
        if hasattr(self, 'upload_view'):
            self.upload_view.grid_forget()
        self.workspace_view.grid(row=0, column=0, sticky="nsew")

    def show_toast(self, message, is_success=True):
        """Displays a soft non-intrusive notification banner inside the app"""
        bg_color = ("#DCFCE7", "#064E3B") if is_success else ("#FEE2E2", "#7F1D1D")
        border_color = ("#86EFAC", "#059669") if is_success else ("#FCA5A5", "#DC2626")
        text_color = ("#14532D", "#F0FDF4") if is_success else ("#7F1D1D", "#FEF2F2")

        self.toast_banner.configure(fg_color=bg_color, border_color=border_color)
        self.lbl_toast_msg.configure(text=message, text_color=text_color)
        self.toast_banner.grid(row=0, column=0, sticky="ew", pady=(0, 10))

    def hide_toast(self):
        if hasattr(self, 'toast_banner'):
            self.toast_banner.grid_forget()

    def create_metric_badge(self, parent, label_text, value_text, icon_emoji, bg_color, accent_color):
        frame = ctk.CTkFrame(parent, fg_color=bg_color, corner_radius=10, border_width=1, border_color=("#F1F5F9", "#334155"))
        
        icon_lbl = ctk.CTkLabel(frame, text=icon_emoji, font=ctk.CTkFont(size=15))
        icon_lbl.pack(side="left", padx=(10, 4), pady=5)

        text_box = ctk.CTkFrame(frame, fg_color="transparent")
        text_box.pack(side="left", padx=(0, 12), pady=4)

        lbl_title = ctk.CTkLabel(
            text_box, 
            text=label_text, 
            font=ctk.CTkFont(family="Segoe UI", size=9, weight="bold"),
            text_color="#94A3B8"
        )
        lbl_title.pack(anchor="w")
        
        lbl_val = ctk.CTkLabel(
            text_box, 
            text=value_text, 
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            text_color=accent_color
        )
        lbl_val.pack(anchor="w")
        frame.lbl_val = lbl_val
        return frame

    def update_tree_theme(self):
        """Soft minimalist styling for Treeview Data Table"""
        style = ttk.Style()
        style.theme_use("default")

        bg_main = "#FFFFFF"
        fg_main = "#1E293B"
        header_bg = "#F8FAFC"
        header_fg = "#475569"

        # Soft subtle selection color
        selected_bg = "#E0F2FE"
        selected_fg = "#0284C7"

        style.configure(
            "Modern.Treeview", 
            background=bg_main, 
            foreground=fg_main, 
            fieldbackground=bg_main, 
            rowheight=36,
            font=('Segoe UI', 11),
            borderwidth=0
        )
        style.map(
            "Modern.Treeview", 
            background=[("selected", selected_bg)],
            foreground=[("selected", selected_fg)]
        )
        style.configure(
            "Modern.Treeview.Heading", 
            background=header_bg, 
            foreground=header_fg, 
            relief="flat",
            font=('Segoe UI', 11, 'bold')
        )

        even_bg = "#FFFFFF"
        odd_bg = "#F8FAFC"
        
        if hasattr(self, 'tree'):
            self.tree.tag_configure('evenrow', background=even_bg)
            self.tree.tag_configure('oddrow', background=odd_bg)

    def select_files(self):
        files = filedialog.askopenfilenames(title="Select BBS PDF Files", filetypes=[("PDF files", "*.pdf")])
        if files:
            self.selected_files = list(files)
            self.update_upload_card()

    def select_folder(self):
        folder = filedialog.askdirectory(title="Select Folder containing PDF files")
        if folder:
            files = sorted(glob.glob(os.path.join(folder, "*.pdf")))
            if files:
                self.selected_files = files
                self.update_upload_card()
            else:
                messagebox.showwarning("No PDFs Found", "No PDF files were found in the selected folder.")

    def update_upload_card(self):
        """Switches Card view from Unselected (Image 1) to Selected (Image 2)"""
        count = len(self.selected_files)
        if count == 0:
            self.dz_selected_frame.place_forget()
            self.btn_convert_now.pack_forget()
            self.progress_container.pack_forget()
            self.dz_unselected_frame.place(relx=0.5, rely=0.5, anchor="center")
        else:
            self.dz_unselected_frame.place_forget()
            if count == 1:
                display_name = os.path.basename(self.selected_files[0])
            else:
                display_name = f"{count} PDF Files Selected"
                
            self.lbl_selected_filename.configure(text=display_name)
            self.dz_selected_frame.place(relx=0.5, rely=0.5, anchor="center")
            self.btn_convert_now.pack(padx=40, pady=(0, 24), fill="x")

    def select_pdf_mode(self, mode_val):
        """Updates selected PDF mode and updates button styling"""
        self.pdf_mode.set(mode_val)
        if mode_val == "with_dimension":
            self.btn_mode_dim.configure(
                fg_color=("#EFF6FF", "#1E3A8A"),
                border_color="#2563EB",
                border_width=2,
                text_color=("#1D4ED8", "#93C5FD"),
                text="📄 PDF with Dimension  ✓\n(Header ABCD)"
            )
            self.btn_mode_nodim.configure(
                fg_color=("#FFFFFF", "#0F172A"),
                border_color=("#E2E8F0", "#334155"),
                border_width=1,
                text_color=("#64748B", "#94A3B8"),
                text="📐 PDF without Dimension\n(Legacy / Shape OCR)"
            )
        else:
            self.btn_mode_dim.configure(
                fg_color=("#FFFFFF", "#0F172A"),
                border_color=("#E2E8F0", "#334155"),
                border_width=1,
                text_color=("#64748B", "#94A3B8"),
                text="📄 PDF with Dimension\n(Header ABCD)"
            )
            self.btn_mode_nodim.configure(
                fg_color=("#EFF6FF", "#1E3A8A"),
                border_color="#2563EB",
                border_width=2,
                text_color=("#1D4ED8", "#93C5FD"),
                text="📐 PDF without Dimension  ✓\n(Legacy / Shape OCR)"
            )

    def start_conversion(self):
        if not self.selected_files:
            cwd_pdfs = sorted(glob.glob("*.pdf"))
            if cwd_pdfs:
                self.selected_files = cwd_pdfs
                self.update_upload_card()
            else:
                messagebox.showwarning("No Files Selected", "Please select PDF file(s) or a folder first.")
                return

        self.btn_convert_now.pack_forget()
        self.progress_container.pack(padx=40, pady=(0, 24), fill="x")
        self.progressbar.set(0.0)
        self.lbl_progress_percent.configure(text="0%")

        threading.Thread(target=self.run_extraction_thread, daemon=True).start()

    def update_progress_cb(self, curr_page, total_pages, status_text):
        progress_val = curr_page / total_pages if total_pages > 0 else 0
        percent_str = f"{int(progress_val * 100)}%"
        self.after(0, lambda: self._apply_progress_update(progress_val, percent_str, status_text))

    def _apply_progress_update(self, progress_val, percent_str, status_text):
        self.progressbar.set(progress_val)
        self.lbl_progress_percent.configure(text=percent_str)
        self.status_msg.configure(text=status_text)

    def run_extraction_thread(self):
        try:
            if self.pdf_mode.get() == "with_dimension":
                self.extracted_data = parse_all_pdfs_with_dimension(self.selected_files, progress_callback=self.update_progress_cb)
            else:
                self.extracted_data = parse_all_pdfs(self.selected_files, progress_callback=self.update_progress_cb)
            self.filtered_data = list(self.extracted_data)
            self.after(0, self.update_ui_post_extraction)
        except Exception as e:
            self.after(0, lambda: messagebox.showerror("Extraction Error", str(e)))
            self.after(0, self.reset_buttons)

    def update_ui_post_extraction(self):
        self.progressbar.set(1.0)
        self.update_summary_metrics()
        
        mode_label = "PDF WITH DIMENSION" if self.pdf_mode.get() == "with_dimension" else "PDF WITHOUT DIMENSION"
        self.lbl_ws_subtitle.configure(text=f"FILE WORKSPACE — {mode_label}")

        if self.selected_files:
            display_title = os.path.basename(self.selected_files[0]) if len(self.selected_files) == 1 else f"{len(self.selected_files)} PDF Files"
            self.lbl_ws_filename.configure(text=display_title)
            
        self.render_tree_data(self.filtered_data)

        # Switch view to Workspace (Soft & Simple View)
        self.show_workspace_view()

        # Display soft non-intrusive notification banner instead of native popup
        self.show_toast(f"✓ Extraction complete ({mode_label})! {len(self.extracted_data)} BBS rows extracted. Double-click any cell to edit before export.")

    def on_search_change(self, event=None):
        query = self.search_box.get().strip().lower()
        if not query:
            self.filtered_data = list(self.extracted_data)
        else:
            self.filtered_data = [
                r for r in self.extracted_data 
                if any(query in str(v).lower() for v in r.values())
            ]
        self.render_tree_data(self.filtered_data)

    def render_tree_data(self, data_list):
        self.destroy_cell_editor()
        for item in self.tree.get_children():
            self.tree.delete(item)

        self.tree_item_map = {}
        for i, r in enumerate(data_list[:500]): # render top 500 matching rows
            vals = [r.get(c, "") for c in self.cols]
            tag = 'evenrow' if i % 2 == 0 else 'oddrow'
            item_id = self.tree.insert("", "end", values=vals, tags=(tag,))
            self.tree_item_map[item_id] = r

    def on_tree_double_click(self, event):
        """Activates inline cell editing on double click or Enter key"""
        focus_item = self.tree.focus()
        if hasattr(event, 'y') and event.y > 0:
            row_id = self.tree.identify_row(event.y)
        else:
            row_id = focus_item

        if not row_id:
            row_id = focus_item
        if not row_id:
            return

        self.tree.selection_set(row_id)
        
        if hasattr(event, 'x') and event.x > 0:
            region = self.tree.identify("region", event.x, event.y)
            if region != "cell":
                return
            column = self.tree.identify_column(event.x)
        else:
            column = "#1"

        if not column or column == "#0":
            return

        try:
            col_idx = int(column.replace("#", "")) - 1
        except ValueError:
            return

        if col_idx < 0 or col_idx >= len(self.cols):
            return

        col_name = self.cols[col_idx]
        bbox = self.tree.bbox(row_id, column)
        if not bbox:
            return

        x, y, w, h = bbox

        current_values = list(self.tree.item(row_id, "values"))
        current_val = str(current_values[col_idx]) if col_idx < len(current_values) else ""

        self.destroy_cell_editor()

        self.cell_editor = tk.Entry(
            self.tree,
            font=('Segoe UI', 11),
            bg="#FFFFFF",
            fg="#0F172A",
            relief="solid",
            bd=1,
            highlightbackground="#2563EB",
            highlightcolor="#2563EB",
            highlightthickness=1,
            insertbackground="#0F172A"
        )
        self.cell_editor.insert(0, current_val)
        self.cell_editor.select_range(0, "end")
        self.cell_editor.place(x=x, y=y, width=w, height=h)
        self.cell_editor.focus_set()

        self.editing_context = {
            'row_id': row_id,
            'col_idx': col_idx,
            'col_name': col_name,
            'original_val': current_val
        }

        self.cell_editor.bind("<Return>", lambda e: self.save_cell_edit())
        self.cell_editor.bind("<KP_Enter>", lambda e: self.save_cell_edit())
        self.cell_editor.bind("<Escape>", lambda e: self.destroy_cell_editor())
        self.cell_editor.bind("<FocusOut>", lambda e: self.save_cell_edit())

    def save_cell_edit(self):
        if not hasattr(self, 'editing_context') or not self.editing_context:
            return

        ctx = self.editing_context
        self.editing_context = None

        if not hasattr(self, 'cell_editor') or not self.cell_editor:
            return

        new_val = self.cell_editor.get().strip()
        row_id = ctx['row_id']
        col_idx = ctx['col_idx']
        col_name = ctx['col_name']
        orig_val = ctx['original_val']

        self.destroy_cell_editor()

        if not self.tree.exists(row_id):
            return

        values = list(self.tree.item(row_id, "values"))
        if col_idx < len(values):
            values[col_idx] = new_val
            
            # Update associated dictionary in self.extracted_data / self.filtered_data
            dict_ref = getattr(self, 'tree_item_map', {}).get(row_id)
            if dict_ref is not None:
                dict_ref[col_name] = new_val

                # Auto-update TOTAL if MBRS or EACH modified
                if col_name in ['No. of MBRS', 'No. of EACH']:
                    try:
                        nbrs = float(dict_ref.get('No. of MBRS', 0))
                        each = float(dict_ref.get('No. of EACH', 0))
                        tot = int(nbrs * each) if (nbrs * each).is_integer() else round(nbrs * each, 2)
                        dict_ref['TOTAL'] = str(tot)
                        tot_idx = self.cols.index('TOTAL')
                        values[tot_idx] = str(tot)
                    except (ValueError, TypeError):
                        pass

            self.tree.item(row_id, values=values)

        self.update_summary_metrics()
        self.show_toast(f"✓ Updated '{col_name}' to '{new_val}'. Changes ready for export.")

    def destroy_cell_editor(self):
        self.editing_context = None
        if hasattr(self, 'cell_editor') and self.cell_editor:
            try:
                self.cell_editor.destroy()
            except Exception:
                pass
            self.cell_editor = None

    def update_summary_metrics(self):
        row_count = len(self.extracted_data)
        unique_members = len(set(r.get('Member', '') for r in self.extracted_data if r.get('Member')))
        unique_shapes = len(set(r.get('SHAPE CODE', '') for r in self.extracted_data if r.get('SHAPE CODE')))
        
        self.badge_rows.lbl_val.configure(text=f"{row_count} rows")
        self.badge_members.lbl_val.configure(text=f"{unique_members} unique")
        self.badge_shapes.lbl_val.configure(text=f"{unique_shapes} codes")

    def save_excel(self):
        if not self.extracted_data:
            return

        # Default directory to User's Downloads folder
        downloads_dir = os.path.join(os.path.expanduser('~'), 'Downloads')
        if not os.path.exists(downloads_dir):
            os.makedirs(downloads_dir, exist_ok=True)

        # Derive default Excel filename directly from the input PDF filename
        if self.selected_files:
            if len(self.selected_files) == 1:
                base_name = os.path.splitext(os.path.basename(self.selected_files[0]))[0]
                default_filename = f"{base_name}.xlsx"
            else:
                first_base = os.path.splitext(os.path.basename(self.selected_files[0]))[0]
                default_filename = f"{first_base}_batch_converted.xlsx"
        else:
            default_filename = "BBS_Converted.xlsx"

        save_path = filedialog.asksaveasfilename(
            title="Save Excel File", 
            initialdir=downloads_dir,
            initialfile=default_filename,
            defaultextension=".xlsx", 
            filetypes=[("Excel Workbook", "*.xlsx")]
        )

        if save_path:
            export_to_excel(self.extracted_data, save_path)
            self.show_toast(f"✓ Saved to Downloads: {os.path.basename(save_path)}")

    def reset_all(self):
        self.selected_files = []
        self.extracted_data = []
        self.filtered_data = []
        self.search_box.delete(0, "end")
        self.hide_toast()
        self.update_upload_card()
        
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        self.badge_rows.lbl_val.configure(text="0 rows")
        self.badge_members.lbl_val.configure(text="0 unique")
        self.badge_shapes.lbl_val.configure(text="0 codes")
        
        self.progressbar.set(0)
        self.show_upload_view()

    def reset_buttons(self):
        self.update_upload_card()

if __name__ == "__main__":
    app = BBSConverterApp()
    app.mainloop()
