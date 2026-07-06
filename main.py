import os
import sys
import json
import webbrowser
import tkinter as tk
from tkinter import filedialog, messagebox, font as tkfont
from PIL import Image, ImageTk


def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller."""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

# =============================================================================
# CONFIGURATION
# =============================================================================

# Pre-defined subfolder types operators commonly use
SUBFOLDER_TYPES = [
    "Input", "Output", "Docs", "Raw", "Processed", "Backup",
    "Logs", "Temp", "Config", "Data", "Reports", "Images",
    "Scripts", "Databases",
]

DEFAULT_TOP_FOLDERS = ["WIP", "SHARED", "ARCHIVE"]

# Preset templates: name -> (top_folders, subfolders)
PRESETS = {
    "Standard Project": (
        ["WIP", "SHARED", "ARCHIVE"],
        ["Input", "Output", "Docs", "Backup"],
    ),
    "Data Processing": (
        ["WIP", "SHARED"],
        ["Input", "Output", "Raw", "Processed", "Logs"],
    ),
    "Report Package": (
        ["SHARED", "ARCHIVE"],
        ["Reports", "Docs", "Images", "Data"],
    ),
}

# =============================================================================
# COLOR THEME
# =============================================================================
COLORS = {
    "bg_dark": "#1e1e2e",
    "bg_card": "#2a2a3d",
    "bg_input": "#363650",
    "bg_preview": "#1a1a2e",
    "fg_primary": "#e0e0f0",
    "fg_secondary": "#9090b0",
    "fg_dim": "#606080",
    "accent_blue": "#5b8def",
    "accent_green": "#4ecb71",
    "accent_red": "#ef5b5b",
    "accent_orange": "#f0a050",
    "accent_purple": "#9b7bef",
    "border": "#3a3a55",
    "tag_bg": "#3d3d5c",
    "preset_bg": "#2d3a50",
    "preset_hover": "#3a4a60",
}


# =============================================================================
# APPLICATION CLASS
# =============================================================================
class FolderCreatorApp:
    COMPANY_URL = "https://www.ruknbim.com/"
    COMPANY_EMAIL = "support@ruknbim.com"

    def __init__(self, root):
        self.root = root
        self.root.title("Folder Creator — RuknBIM")
        self.root.geometry("960x750")
        self.root.minsize(800, 620)
        self.root.configure(bg=COLORS["bg_dark"])

        # Data
        self.base_path_var = tk.StringVar()
        self.top_folder_vars = {}  # name -> BooleanVar
        self.subfolder_vars = {}   # name -> BooleanVar
        self.custom_top_folders = []  # list of strings
        self.custom_subfolders = []   # list of strings
        self.status_var = tk.StringVar(value="✅  Ready — select folders and click Create")

        # Fonts
        self.font_title = ("Segoe UI", 15, "bold")
        self.font_section = ("Segoe UI", 11, "bold")
        self.font_normal = ("Segoe UI", 10)
        self.font_small = ("Segoe UI", 9)
        self.font_button = ("Segoe UI", 10, "bold")
        self.font_tree = ("Consolas", 10)
        self.font_status = ("Segoe UI", 10)

        self._build_ui()
        self._update_preview()

    # -------------------------------------------------------------------------
    # UI BUILDING
    # -------------------------------------------------------------------------
    def _build_ui(self):
        # Title bar with logo
        title_bar = tk.Frame(self.root, bg=COLORS["accent_blue"], height=52)
        title_bar.pack(fill="x")
        title_bar.pack_propagate(False)

        # Load and display logo
        try:
            logo_path = resource_path("logo.png")
            logo_img = Image.open(logo_path)
            logo_img = logo_img.resize((120, 38), Image.LANCZOS)
            self._logo_photo = ImageTk.PhotoImage(logo_img)
            logo_label = tk.Label(title_bar, image=self._logo_photo,
                                  bg=COLORS["accent_blue"], cursor="hand2")
            logo_label.pack(side="left", padx=(12, 6), pady=6)
            logo_label.bind("<Button-1>", lambda e: webbrowser.open(self.COMPANY_URL))
        except Exception:
            pass  # Graceful fallback if logo not found

        tk.Label(
            title_bar, text="📁  Folder Creator", font=self.font_title,
            bg=COLORS["accent_blue"], fg="white", padx=8
        ).pack(side="left", fill="y")

        # Main container
        container = tk.Frame(self.root, bg=COLORS["bg_dark"], padx=12, pady=8)
        container.pack(fill="both", expand=True)

        # ---------- Row 1: Base path ----------
        self._build_path_section(container)

        # ---------- Row 2: Presets ----------
        self._build_preset_section(container)

        # ---------- Row 3: Two-column layout ----------
        columns = tk.Frame(container, bg=COLORS["bg_dark"])
        columns.pack(fill="both", expand=True, pady=(6, 0))
        columns.columnconfigure(0, weight=3, minsize=380)
        columns.columnconfigure(1, weight=2, minsize=300)

        # Left column: selections
        left = tk.Frame(columns, bg=COLORS["bg_dark"])
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 6))

        self._build_top_folders_section(left)
        self._build_subfolder_section(left)

        # Right column: preview
        right = tk.Frame(columns, bg=COLORS["bg_dark"])
        right.grid(row=0, column=1, sticky="nsew", padx=(6, 0))

        self._build_preview_section(right)

        # ---------- Row 4: Buttons ----------
        self._build_buttons(container)

        # ---------- Row 5: Company footer ----------
        self._build_company_footer()

        # ---------- Row 6: Status bar ----------
        self._build_status_bar()

    def _make_card(self, parent, **pack_kwargs):
        """Create a rounded-look card frame."""
        card = tk.Frame(parent, bg=COLORS["bg_card"], highlightbackground=COLORS["border"],
                        highlightthickness=1, padx=12, pady=10)
        card.pack(fill="x", **pack_kwargs)
        return card

    # --- Path section ---
    def _build_path_section(self, parent):
        card = self._make_card(parent, pady=(0, 6))
        tk.Label(card, text="Base Path", font=self.font_section,
                 bg=COLORS["bg_card"], fg=COLORS["accent_blue"]).pack(anchor="w")

        row = tk.Frame(card, bg=COLORS["bg_card"])
        row.pack(fill="x", pady=(6, 0))

        entry = tk.Entry(row, textvariable=self.base_path_var, font=self.font_normal,
                         bg=COLORS["bg_input"], fg=COLORS["fg_primary"],
                         insertbackground=COLORS["fg_primary"], relief="flat",
                         highlightthickness=1, highlightbackground=COLORS["border"])
        entry.pack(side="left", fill="x", expand=True, ipady=4)

        browse_btn = tk.Button(row, text="📂 Browse", font=self.font_button,
                               bg=COLORS["accent_green"], fg="white", relief="flat",
                               padx=14, pady=2, cursor="hand2",
                               activebackground="#3aaa55",
                               command=self._browse_path)
        browse_btn.pack(side="left", padx=(8, 0))

    # --- Presets section ---
    def _build_preset_section(self, parent):
        card = self._make_card(parent, pady=(0, 6))
        tk.Label(card, text="⚡ Quick Presets", font=self.font_section,
                 bg=COLORS["bg_card"], fg=COLORS["accent_orange"]).pack(anchor="w")

        row = tk.Frame(card, bg=COLORS["bg_card"])
        row.pack(fill="x", pady=(6, 0))

        for name in PRESETS:
            btn = tk.Button(row, text=name, font=self.font_normal,
                            bg=COLORS["preset_bg"], fg=COLORS["fg_primary"],
                            relief="flat", padx=16, pady=5, cursor="hand2",
                            activebackground=COLORS["preset_hover"],
                            highlightthickness=1,
                            highlightbackground=COLORS["border"],
                            command=lambda n=name: self._apply_preset(n))
            btn.pack(side="left", padx=(0, 8))

    # --- Top folders section ---
    def _build_top_folders_section(self, parent):
        card = self._make_card(parent, pady=(0, 6))
        tk.Label(card, text="Top-Level Folders", font=self.font_section,
                 bg=COLORS["bg_card"], fg=COLORS["accent_blue"]).pack(anchor="w")

        # Default checkboxes
        cb_row = tk.Frame(card, bg=COLORS["bg_card"])
        cb_row.pack(fill="x", pady=(6, 0))

        for folder in DEFAULT_TOP_FOLDERS:
            var = tk.BooleanVar(value=True)
            var.trace_add("write", lambda *_: self._update_preview())
            cb = tk.Checkbutton(cb_row, text=folder, variable=var,
                                font=self.font_normal,
                                bg=COLORS["bg_card"], fg=COLORS["fg_primary"],
                                selectcolor=COLORS["bg_input"],
                                activebackground=COLORS["bg_card"],
                                activeforeground=COLORS["fg_primary"],
                                cursor="hand2")
            cb.pack(side="left", padx=(0, 18))
            self.top_folder_vars[folder] = var

        # Custom top-level folders
        add_row = tk.Frame(card, bg=COLORS["bg_card"])
        add_row.pack(fill="x", pady=(8, 0))

        tk.Label(add_row, text="+ Add custom:", font=self.font_small,
                 bg=COLORS["bg_card"], fg=COLORS["fg_secondary"]).pack(side="left")

        self.custom_top_entry = tk.Entry(add_row, font=self.font_normal, width=16,
                                         bg=COLORS["bg_input"], fg=COLORS["fg_primary"],
                                         insertbackground=COLORS["fg_primary"],
                                         relief="flat", highlightthickness=1,
                                         highlightbackground=COLORS["border"])
        self.custom_top_entry.pack(side="left", padx=(6, 4), ipady=2)
        self.custom_top_entry.bind("<Return>", lambda e: self._add_custom_top())

        tk.Button(add_row, text="Add", font=self.font_small,
                  bg=COLORS["accent_blue"], fg="white", relief="flat",
                  padx=10, cursor="hand2", activebackground="#4a7adf",
                  command=self._add_custom_top).pack(side="left")

        # Tags container for custom top folders
        self.custom_top_tags_frame = tk.Frame(card, bg=COLORS["bg_card"])
        self.custom_top_tags_frame.pack(fill="x", pady=(4, 0))

    # --- Subfolder section ---
    def _build_subfolder_section(self, parent):
        card = self._make_card(parent, pady=(0, 6))
        card.pack(fill="both", expand=True)
        tk.Label(card, text="Subfolders", font=self.font_section,
                 bg=COLORS["bg_card"], fg=COLORS["accent_purple"]).pack(anchor="w")

        # Checkbox grid
        grid_frame = tk.Frame(card, bg=COLORS["bg_card"])
        grid_frame.pack(fill="x", pady=(6, 0))

        cols = 4
        for i, name in enumerate(SUBFOLDER_TYPES):
            var = tk.BooleanVar(value=False)
            var.trace_add("write", lambda *_: self._update_preview())
            cb = tk.Checkbutton(grid_frame, text=name, variable=var,
                                font=self.font_normal,
                                bg=COLORS["bg_card"], fg=COLORS["fg_primary"],
                                selectcolor=COLORS["bg_input"],
                                activebackground=COLORS["bg_card"],
                                activeforeground=COLORS["fg_primary"],
                                width=11, anchor="w", cursor="hand2")
            cb.grid(row=i // cols, column=i % cols, sticky="w", padx=2, pady=1)
            self.subfolder_vars[name] = var

        # Custom subfolder entry
        add_row = tk.Frame(card, bg=COLORS["bg_card"])
        add_row.pack(fill="x", pady=(8, 0))

        tk.Label(add_row, text="+ Custom:", font=self.font_small,
                 bg=COLORS["bg_card"], fg=COLORS["fg_secondary"]).pack(side="left")

        self.custom_sub_entry = tk.Entry(add_row, font=self.font_normal, width=16,
                                         bg=COLORS["bg_input"], fg=COLORS["fg_primary"],
                                         insertbackground=COLORS["fg_primary"],
                                         relief="flat", highlightthickness=1,
                                         highlightbackground=COLORS["border"])
        self.custom_sub_entry.pack(side="left", padx=(6, 4), ipady=2)
        self.custom_sub_entry.bind("<Return>", lambda e: self._add_custom_sub())

        tk.Button(add_row, text="Add", font=self.font_small,
                  bg=COLORS["accent_purple"], fg="white", relief="flat",
                  padx=10, cursor="hand2", activebackground="#8a6adf",
                  command=self._add_custom_sub).pack(side="left")

        # Tags container for custom subfolders
        self.custom_sub_tags_frame = tk.Frame(card, bg=COLORS["bg_card"])
        self.custom_sub_tags_frame.pack(fill="x", pady=(4, 0))

    # --- Preview section ---
    def _build_preview_section(self, parent):
        card = tk.Frame(parent, bg=COLORS["bg_card"], highlightbackground=COLORS["border"],
                        highlightthickness=1, padx=12, pady=10)
        card.pack(fill="both", expand=True)

        tk.Label(card, text="📂 Live Preview", font=self.font_section,
                 bg=COLORS["bg_card"], fg=COLORS["accent_green"]).pack(anchor="w")

        self.preview_text = tk.Text(
            card, font=self.font_tree, bg=COLORS["bg_preview"],
            fg=COLORS["fg_primary"], relief="flat", wrap="none",
            highlightthickness=1, highlightbackground=COLORS["border"],
            padx=10, pady=8, cursor="arrow"
        )
        self.preview_text.pack(fill="both", expand=True, pady=(6, 0))
        self.preview_text.config(state="disabled")

        # Tag colors
        self.preview_text.tag_configure("folder_icon", foreground=COLORS["accent_orange"])
        self.preview_text.tag_configure("top_folder", foreground=COLORS["accent_blue"],
                                         font=("Consolas", 10, "bold"))
        self.preview_text.tag_configure("sub_folder", foreground=COLORS["fg_primary"])
        self.preview_text.tag_configure("tree_line", foreground=COLORS["fg_dim"])
        self.preview_text.tag_configure("base_path", foreground=COLORS["accent_green"],
                                         font=("Consolas", 10, "bold"))

    # --- Buttons ---
    def _build_buttons(self, parent):
        btn_frame = tk.Frame(parent, bg=COLORS["bg_dark"])
        btn_frame.pack(fill="x", pady=(8, 0))

        create_btn = tk.Button(
            btn_frame, text="🚀  Create Folders", font=("Segoe UI", 12, "bold"),
            bg=COLORS["accent_blue"], fg="white", relief="flat",
            padx=32, pady=8, cursor="hand2", activebackground="#4a7adf",
            command=self._create_folders
        )
        create_btn.pack(side="left", expand=True)

        reset_btn = tk.Button(
            btn_frame, text="🔄  Reset", font=self.font_button,
            bg=COLORS["bg_card"], fg=COLORS["fg_secondary"], relief="flat",
            padx=20, pady=8, cursor="hand2",
            highlightthickness=1, highlightbackground=COLORS["border"],
            activebackground=COLORS["bg_input"],
            command=self._reset_all
        )
        reset_btn.pack(side="left", padx=(8, 0))

        export_btn = tk.Button(
            btn_frame, text="💾  Export", font=self.font_button,
            bg=COLORS["accent_green"], fg="white", relief="flat",
            padx=16, pady=8, cursor="hand2", activebackground="#3aaa55",
            command=self._export_config
        )
        export_btn.pack(side="left", padx=(8, 0))

        import_btn = tk.Button(
            btn_frame, text="📂  Import", font=self.font_button,
            bg=COLORS["accent_orange"], fg="white", relief="flat",
            padx=16, pady=8, cursor="hand2", activebackground="#d08a40",
            command=self._import_config
        )
        import_btn.pack(side="left", padx=(8, 0))

    # --- Company footer ---
    def _build_company_footer(self):
        footer = tk.Frame(self.root, bg=COLORS["bg_dark"], height=36)
        footer.pack(fill="x", side="bottom", pady=(4, 0))
        footer.pack_propagate(False)

        inner = tk.Frame(footer, bg=COLORS["bg_dark"])
        inner.pack(expand=True)

        # Website link
        web_label = tk.Label(
            inner, text="🌐 www.ruknbim.com",
            font=("Segoe UI", 9), bg=COLORS["bg_dark"],
            fg=COLORS["accent_blue"], cursor="hand2"
        )
        web_label.pack(side="left", padx=(0, 20))
        web_label.bind("<Button-1>", lambda e: webbrowser.open(self.COMPANY_URL))
        web_label.bind("<Enter>", lambda e: web_label.config(fg="#8ab4ff"))
        web_label.bind("<Leave>", lambda e: web_label.config(fg=COLORS["accent_blue"]))

        # Separator
        tk.Label(inner, text="|", font=("Segoe UI", 9),
                 bg=COLORS["bg_dark"], fg=COLORS["fg_dim"]).pack(side="left", padx=(0, 20))

        # Email
        email_label = tk.Label(
            inner, text=f"✉ {self.COMPANY_EMAIL}",
            font=("Segoe UI", 9), bg=COLORS["bg_dark"],
            fg=COLORS["accent_green"], cursor="hand2"
        )
        email_label.pack(side="left", padx=(0, 20))
        email_label.bind("<Button-1>", lambda e: webbrowser.open(f"mailto:{self.COMPANY_EMAIL}"))
        email_label.bind("<Enter>", lambda e: email_label.config(fg="#7aeaa0"))
        email_label.bind("<Leave>", lambda e: email_label.config(fg=COLORS["accent_green"]))

        # Separator
        tk.Label(inner, text="|", font=("Segoe UI", 9),
                 bg=COLORS["bg_dark"], fg=COLORS["fg_dim"]).pack(side="left", padx=(0, 20))

        # Company name
        tk.Label(
            inner, text="Powered by RuknBIM",
            font=("Segoe UI", 9, "italic"), bg=COLORS["bg_dark"],
            fg=COLORS["fg_dim"]
        ).pack(side="left")

    # --- Status bar ---
    def _build_status_bar(self):
        bar = tk.Frame(self.root, bg=COLORS["bg_card"], height=32)
        bar.pack(fill="x", side="bottom")
        bar.pack_propagate(False)
        tk.Label(bar, textvariable=self.status_var, font=self.font_status,
                 bg=COLORS["bg_card"], fg=COLORS["fg_secondary"],
                 padx=12).pack(side="left", fill="y")

    # -------------------------------------------------------------------------
    # ACTIONS
    # -------------------------------------------------------------------------
    def _browse_path(self):
        path = filedialog.askdirectory()
        if path:
            self.base_path_var.set(path)
            self._update_preview()

    def _export_config(self):
        """Export the current folder configuration to a JSON file."""
        config = {
            "base_path": self.base_path_var.get().strip(),
            "top_folders": {name: var.get() for name, var in self.top_folder_vars.items()},
            "subfolders": {name: var.get() for name, var in self.subfolder_vars.items()},
            "custom_top_folders": list(self.custom_top_folders),
            "custom_subfolders": list(self.custom_subfolders),
        }
        file_path = filedialog.asksaveasfilename(
            title="Export Configuration",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            initialfile="folder_config.json"
        )
        if not file_path:
            return
        try:
            with open(file_path, "w", encoding="utf-8") as f:
                json.dump(config, f, indent=2, ensure_ascii=False)
            self.status_var.set(f"💾  Configuration exported to {os.path.basename(file_path)}")
        except Exception as e:
            messagebox.showerror("Export Error", str(e))
            self.status_var.set("❌  Export failed")

    def _import_config(self):
        """Import a folder configuration from a JSON file."""
        file_path = filedialog.askopenfilename(
            title="Import Configuration",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if not file_path:
            return
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                config = json.load(f)

            # Apply base path
            self.base_path_var.set(config.get("base_path", ""))

            # Apply top folder selections
            saved_top = config.get("top_folders", {})
            for name, var in self.top_folder_vars.items():
                var.set(saved_top.get(name, False))

            # Apply subfolder selections
            saved_sub = config.get("subfolders", {})
            for name, var in self.subfolder_vars.items():
                var.set(saved_sub.get(name, False))

            # Apply custom top folders
            self.custom_top_folders.clear()
            for name in config.get("custom_top_folders", []):
                if name not in self.custom_top_folders and name not in DEFAULT_TOP_FOLDERS:
                    self.custom_top_folders.append(name)
            self._refresh_custom_top_tags()

            # Apply custom subfolders
            self.custom_subfolders.clear()
            for name in config.get("custom_subfolders", []):
                if name not in self.custom_subfolders and name not in SUBFOLDER_TYPES:
                    self.custom_subfolders.append(name)
            self._refresh_custom_sub_tags()

            self._update_preview()
            self.status_var.set(f"📂  Configuration imported from {os.path.basename(file_path)}")
        except json.JSONDecodeError:
            messagebox.showerror("Import Error", "Invalid JSON file.")
            self.status_var.set("❌  Import failed — invalid file")
        except Exception as e:
            messagebox.showerror("Import Error", str(e))
            self.status_var.set("❌  Import failed")

    def _apply_preset(self, name):
        top_folders, subfolders = PRESETS[name]

        # Set top folder checkboxes
        for folder, var in self.top_folder_vars.items():
            var.set(folder in top_folders)

        # Set subfolder checkboxes
        for folder, var in self.subfolder_vars.items():
            var.set(folder in subfolders)

        self.status_var.set(f"⚡  Preset '{name}' applied")
        self._update_preview()

    def _add_custom_top(self):
        name = self.custom_top_entry.get().strip()
        if not name:
            return
        if name in self.custom_top_folders or name in DEFAULT_TOP_FOLDERS:
            self.status_var.set(f"⚠️  '{name}' already exists")
            return
        self.custom_top_folders.append(name)
        self.custom_top_entry.delete(0, tk.END)
        self._refresh_custom_top_tags()
        self._update_preview()

    def _remove_custom_top(self, name):
        if name in self.custom_top_folders:
            self.custom_top_folders.remove(name)
        self._refresh_custom_top_tags()
        self._update_preview()

    def _refresh_custom_top_tags(self):
        for widget in self.custom_top_tags_frame.winfo_children():
            widget.destroy()
        for name in self.custom_top_folders:
            tag = tk.Frame(self.custom_top_tags_frame, bg=COLORS["tag_bg"],
                           padx=6, pady=2, highlightbackground=COLORS["border"],
                           highlightthickness=1)
            tag.pack(side="left", padx=(0, 4), pady=2)
            tk.Label(tag, text=name, font=self.font_small,
                     bg=COLORS["tag_bg"], fg=COLORS["fg_primary"]).pack(side="left")
            tk.Button(tag, text="✕", font=("Segoe UI", 8), bg=COLORS["tag_bg"],
                      fg=COLORS["accent_red"], relief="flat", padx=2, cursor="hand2",
                      activebackground=COLORS["tag_bg"],
                      command=lambda n=name: self._remove_custom_top(n)).pack(side="left", padx=(4, 0))

    def _add_custom_sub(self):
        name = self.custom_sub_entry.get().strip()
        if not name:
            return
        if name in self.custom_subfolders or name in SUBFOLDER_TYPES:
            self.status_var.set(f"⚠️  '{name}' already exists")
            return
        self.custom_subfolders.append(name)
        self.custom_sub_entry.delete(0, tk.END)
        self._refresh_custom_sub_tags()
        self._update_preview()

    def _remove_custom_sub(self, name):
        if name in self.custom_subfolders:
            self.custom_subfolders.remove(name)
        self._refresh_custom_sub_tags()
        self._update_preview()

    def _refresh_custom_sub_tags(self):
        for widget in self.custom_sub_tags_frame.winfo_children():
            widget.destroy()
        for name in self.custom_subfolders:
            tag = tk.Frame(self.custom_sub_tags_frame, bg=COLORS["tag_bg"],
                           padx=6, pady=2, highlightbackground=COLORS["border"],
                           highlightthickness=1)
            tag.pack(side="left", padx=(0, 4), pady=2)
            tk.Label(tag, text=name, font=self.font_small,
                     bg=COLORS["tag_bg"], fg=COLORS["fg_primary"]).pack(side="left")
            tk.Button(tag, text="✕", font=("Segoe UI", 8), bg=COLORS["tag_bg"],
                      fg=COLORS["accent_red"], relief="flat", padx=2, cursor="hand2",
                      activebackground=COLORS["tag_bg"],
                      command=lambda n=name: self._remove_custom_sub(n)).pack(side="left", padx=(4, 0))

    def _get_selected_top_folders(self):
        selected = [f for f, var in self.top_folder_vars.items() if var.get()]
        selected.extend(self.custom_top_folders)
        return selected

    def _get_selected_subfolders(self):
        selected = [f for f, var in self.subfolder_vars.items() if var.get()]
        selected.extend(self.custom_subfolders)
        return selected

    def _update_preview(self):
        top_folders = self._get_selected_top_folders()
        subfolders = self._get_selected_subfolders()
        base = self.base_path_var.get().strip() or "(select a base path)"

        self.preview_text.config(state="normal")
        self.preview_text.delete("1.0", tk.END)

        # Base path
        self.preview_text.insert(tk.END, "📁 ", "folder_icon")
        self.preview_text.insert(tk.END, os.path.basename(base) or base, "base_path")
        self.preview_text.insert(tk.END, "\n")

        if not top_folders:
            self.preview_text.insert(tk.END, "    (no top folders selected)\n", "tree_line")
            self.preview_text.config(state="disabled")
            return

        for ti, top in enumerate(top_folders):
            is_last_top = (ti == len(top_folders) - 1)
            connector = "└── " if is_last_top else "├── "
            prefix = "    " if is_last_top else "│   "

            self.preview_text.insert(tk.END, connector, "tree_line")
            self.preview_text.insert(tk.END, "📁 ", "folder_icon")
            self.preview_text.insert(tk.END, top, "top_folder")
            self.preview_text.insert(tk.END, "\n")

            if subfolders:
                for si, sub in enumerate(subfolders):
                    is_last_sub = (si == len(subfolders) - 1)
                    sub_connector = "└── " if is_last_sub else "├── "
                    self.preview_text.insert(tk.END, prefix, "tree_line")
                    self.preview_text.insert(tk.END, sub_connector, "tree_line")
                    self.preview_text.insert(tk.END, "📄 ", "folder_icon")
                    self.preview_text.insert(tk.END, sub, "sub_folder")
                    self.preview_text.insert(tk.END, "\n")
            else:
                self.preview_text.insert(tk.END, prefix, "tree_line")
                self.preview_text.insert(tk.END, "(no subfolders)\n", "tree_line")

        self.preview_text.config(state="disabled")

    def _reset_all(self):
        self.base_path_var.set("")
        for var in self.top_folder_vars.values():
            var.set(True)
        for var in self.subfolder_vars.values():
            var.set(False)
        self.custom_top_folders.clear()
        self.custom_subfolders.clear()
        self._refresh_custom_top_tags()
        self._refresh_custom_sub_tags()
        self.custom_top_entry.delete(0, tk.END)
        self.custom_sub_entry.delete(0, tk.END)
        self.status_var.set("🔄  Reset — all selections cleared")
        self._update_preview()

    def _create_folders(self):
        base_path = self.base_path_var.get().strip()
        top_folders = self._get_selected_top_folders()
        subfolders = self._get_selected_subfolders()

        # Validation
        if not base_path:
            messagebox.showerror("Error", "Please select a base path.")
            self.status_var.set("❌  Error — no base path selected")
            return
        if not top_folders:
            messagebox.showerror("Error", "Please select at least one top-level folder.")
            self.status_var.set("❌  Error — no top folders selected")
            return
        if not subfolders:
            messagebox.showerror("Error", "Please select at least one subfolder.")
            self.status_var.set("❌  Error — no subfolders selected")
            return

        created = []
        try:
            os.makedirs(base_path, exist_ok=True)

            for top in top_folders:
                top_path = os.path.join(base_path, top)
                os.makedirs(top_path, exist_ok=True)
                created.append(top_path)

                for sub in subfolders:
                    sub_path = os.path.join(top_path, sub)
                    os.makedirs(sub_path, exist_ok=True)
                    created.append(sub_path)

            total = len(created)
            self.status_var.set(
                f"✅  Success — created {total} folders in {os.path.basename(base_path)}"
            )
            messagebox.showinfo(
                "Success",
                f"Created {total} folders successfully!\n\n"
                f"Base: {base_path}\n"
                f"Top folders: {', '.join(top_folders)}\n"
                f"Subfolders: {', '.join(subfolders)}"
            )

        except PermissionError as e:
            messagebox.showerror("Error", f"Permission denied:\n{e}")
            self.status_var.set("❌  Error — permission denied")
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.status_var.set(f"❌  Error — {e}")


# =============================================================================
# MAIN
# =============================================================================
if __name__ == "__main__":
    root = tk.Tk()
    app = FolderCreatorApp(root)
    root.mainloop()