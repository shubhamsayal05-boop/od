"""Tool Tutorial Dashboard — Tkinter desktop application."""

from __future__ import annotations

import json
import os
import platform
import subprocess
import sys
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk


APP_DIR = Path(__file__).resolve().parent
TOOLS_FILE = APP_DIR / "tools.json"

COLORS = {
    "bg": "#0f1419",
    "card": "#1e2836",
    "card_hover": "#243044",
    "border": "#2d3a4d",
    "text": "#e8edf4",
    "muted": "#8b9cb3",
    "accent": "#3b82f6",
    "accent_soft": "#1a2d4a",
    "accent_dark": "#2563eb",
    "warning_bg": "#2a2318",
    "warning_border": "#6b4f1a",
    "warning_text": "#f59e0b",
    "success": "#22c55e",
}


def load_tools() -> list[dict]:
    with open(TOOLS_FILE, encoding="utf-8") as f:
        return json.load(f)


def open_file(path: Path) -> None:
    if not path.exists():
        messagebox.showwarning("File not found", f"Presentation not found:\n{path}")
        return

    system = platform.system()
    try:
        if system == "Windows":
            os.startfile(path)  # noqa: S606
        elif system == "Darwin":
            subprocess.run(["open", str(path)], check=True)
        else:
            subprocess.run(["xdg-open", str(path)], check=True)
    except Exception as exc:  # noqa: BLE001
        messagebox.showerror("Cannot open file", f"Could not open presentation:\n{exc}")


class ToolTutorialApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("DriveScope Tutorial Hub")
        self.geometry("1050x720")
        self.minsize(900, 600)
        self.configure(bg=COLORS["bg"])

        self.tools = load_tools()
        self.search_var = tk.StringVar()
        self._ui_ready = False
        self._search_placeholder = "Search tools..."

        self._setup_styles()
        self._build_ui()
        self._ui_ready = True
        self.search_var.trace_add("write", lambda *_: self._filter_tools())
        self._show_home()

    def _setup_styles(self) -> None:
        style = ttk.Style(self)
        style.theme_use("clam")

        style.configure("TFrame", background=COLORS["bg"])
        style.configure("Card.TFrame", background=COLORS["card"])
        style.configure(
            "Title.TLabel",
            background=COLORS["bg"],
            foreground=COLORS["text"],
            font=("Segoe UI", 22, "bold"),
        )
        style.configure(
            "Subtitle.TLabel",
            background=COLORS["bg"],
            foreground=COLORS["muted"],
            font=("Segoe UI", 11),
        )
        style.configure(
            "CardTitle.TLabel",
            background=COLORS["card"],
            foreground=COLORS["text"],
            font=("Segoe UI", 13, "bold"),
        )
        style.configure(
            "CardText.TLabel",
            background=COLORS["card"],
            foreground=COLORS["muted"],
            font=("Segoe UI", 10),
        )
        style.configure(
            "Meta.TLabel",
            background=COLORS["card"],
            foreground=COLORS["accent"],
            font=("Segoe UI", 9, "bold"),
        )
        style.configure(
            "StepTitle.TLabel",
            background=COLORS["card"],
            foreground=COLORS["text"],
            font=("Segoe UI", 16, "bold"),
        )
        style.configure(
            "StepBody.TLabel",
            background=COLORS["card"],
            foreground=COLORS["muted"],
            font=("Segoe UI", 11),
            wraplength=520,
        )
        style.configure(
            "TipsTitle.TLabel",
            background=COLORS["warning_bg"],
            foreground=COLORS["warning_text"],
            font=("Segoe UI", 9, "bold"),
        )
        style.configure(
            "TipsBody.TLabel",
            background=COLORS["warning_bg"],
            foreground=COLORS["muted"],
            font=("Segoe UI", 10),
            wraplength=480,
        )
        style.configure(
            "Accent.TButton",
            font=("Segoe UI", 10, "bold"),
            padding=(14, 8),
        )
        style.map(
            "Accent.TButton",
            background=[("active", COLORS["accent_dark"]), ("!disabled", COLORS["accent"])],
            foreground=[("!disabled", "white")],
        )
        style.configure(
            "Ghost.TButton",
            font=("Segoe UI", 10),
            padding=(12, 8),
        )
        style.map(
            "Ghost.TButton",
            background=[("active", COLORS["card_hover"]), ("!disabled", COLORS["card"])],
            foreground=[("!disabled", COLORS["text"])],
        )

    def _build_ui(self) -> None:
        self.main = ttk.Frame(self, padding=24)
        self.main.pack(fill=tk.BOTH, expand=True)

        self.home_frame = ttk.Frame(self.main)
        self.tutorial_frame = ttk.Frame(self.main)

        header = ttk.Frame(self.home_frame)
        header.pack(fill=tk.X, pady=(0, 18))
        ttk.Label(header, text="DriveScope Tutorial Hub", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            header,
            text="Step-by-step guides for DriveScope and related driveability tools.",
            style="Subtitle.TLabel",
        ).pack(anchor="w", pady=(4, 0))

        search_row = ttk.Frame(self.home_frame)
        search_row.pack(fill=tk.X, pady=(0, 16))
        self.search_entry = tk.Entry(
            search_row,
            textvariable=self.search_var,
            font=("Segoe UI", 11),
            bg=COLORS["card"],
            fg=COLORS["text"],
            insertbackground=COLORS["text"],
            relief=tk.FLAT,
            highlightthickness=1,
            highlightbackground=COLORS["border"],
            highlightcolor=COLORS["accent"],
        )
        self.search_entry.pack(fill=tk.X, ipady=8)
        self.search_var.set(self._search_placeholder)
        self.search_entry.config(fg=COLORS["muted"])

        def on_focus_in(_event: tk.Event) -> None:
            if self.search_var.get() == self._search_placeholder:
                self.search_var.set("")
                self.search_entry.config(fg=COLORS["text"])

        def on_focus_out(_event: tk.Event) -> None:
            if not self.search_var.get().strip():
                self.search_var.set(self._search_placeholder)
                self.search_entry.config(fg=COLORS["muted"])

        self.search_entry.bind("<FocusIn>", on_focus_in)
        self.search_entry.bind("<FocusOut>", on_focus_out)

        canvas_wrap = ttk.Frame(self.home_frame)
        canvas_wrap.pack(fill=tk.BOTH, expand=True)

        self.canvas = tk.Canvas(canvas_wrap, bg=COLORS["bg"], highlightthickness=0)
        scrollbar = ttk.Scrollbar(canvas_wrap, orient=tk.VERTICAL, command=self.canvas.yview)
        self.tool_grid = ttk.Frame(self.canvas)

        self.tool_grid.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")),
        )
        self.canvas.create_window((0, 0), window=self.tool_grid, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind_all("<Button-4>", self._on_mousewheel_linux)
        self.canvas.bind_all("<Button-5>", self._on_mousewheel_linux)

        self._render_tool_cards(self.tools)

        self._build_tutorial_frame()

    def _build_tutorial_frame(self) -> None:
        top = ttk.Frame(self.tutorial_frame)
        top.pack(fill=tk.X, pady=(0, 16))

        ttk.Button(top, text="← All tools", style="Ghost.TButton", command=self._show_home).pack(
            side=tk.LEFT
        )
        self.open_ppt_btn = ttk.Button(
            top,
            text="Open Presentation",
            style="Accent.TButton",
            command=self._open_current_ppt,
        )
        self.open_ppt_btn.pack(side=tk.RIGHT)

        body = ttk.Frame(self.tutorial_frame)
        body.pack(fill=tk.BOTH, expand=True)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1)

        left = ttk.Frame(body, style="Card.TFrame", padding=16)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 12))

        ttk.Label(left, text="Steps", style="CardTitle.TLabel").pack(anchor="w")
        self.progress_label = ttk.Label(left, text="0/0", style="CardText.TLabel")
        self.progress_label.pack(anchor="w", pady=(2, 10))

        self.step_listbox = tk.Listbox(
            left,
            width=34,
            height=18,
            font=("Segoe UI", 10),
            bg=COLORS["bg"],
            fg=COLORS["text"],
            selectbackground=COLORS["accent"],
            selectforeground="white",
            activestyle="none",
            highlightthickness=1,
            highlightbackground=COLORS["border"],
            relief=tk.FLAT,
            borderwidth=0,
        )
        self.step_listbox.pack(fill=tk.BOTH, expand=True)
        self.step_listbox.bind("<<ListboxSelect>>", self._on_step_selected)

        right = ttk.Frame(body, style="Card.TFrame", padding=24)
        right.grid(row=0, column=1, sticky="nsew")

        self.tool_title = ttk.Label(right, text="", style="StepTitle.TLabel")
        self.tool_title.pack(anchor="w")
        self.tool_desc = ttk.Label(right, text="", style="StepBody.TLabel")
        self.tool_desc.pack(anchor="w", pady=(4, 18))

        self.step_heading = ttk.Label(right, text="", style="StepTitle.TLabel")
        self.step_heading.pack(anchor="w")
        self.step_body = ttk.Label(right, text="", style="StepBody.TLabel")
        self.step_body.pack(anchor="w", pady=(8, 16))

        self.tips_frame = tk.Frame(right, bg=COLORS["warning_bg"], padx=12, pady=12)
        self.tips_body = tk.Label(
            self.tips_frame,
            text="",
            bg=COLORS["warning_bg"],
            fg=COLORS["muted"],
            font=("Segoe UI", 10),
            justify=tk.LEFT,
            wraplength=520,
        )
        self.tips_body.pack(anchor="w")

        nav = ttk.Frame(right)
        nav.pack(fill=tk.X, pady=(24, 0))
        self.prev_btn = ttk.Button(nav, text="← Previous", style="Ghost.TButton", command=self._prev_step)
        self.prev_btn.pack(side=tk.LEFT)
        self.next_btn = ttk.Button(nav, text="Next →", style="Accent.TButton", command=self._next_step)
        self.next_btn.pack(side=tk.RIGHT)

        self.current_tool: dict | None = None
        self.current_step = 0

    def _on_mousewheel(self, event: tk.Event) -> None:
        if self.home_frame.winfo_ismapped():
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _on_mousewheel_linux(self, event: tk.Event) -> None:
        if self.home_frame.winfo_ismapped():
            direction = -1 if event.num == 4 else 1
            self.canvas.yview_scroll(direction, "units")

    def _active_search_query(self) -> str:
        query = self.search_var.get().strip().lower()
        if query == self._search_placeholder.lower():
            return ""
        return query

    def _filter_tools(self) -> None:
        if not self._ui_ready:
            return
        query = self._active_search_query()
        if not query:
            self._render_tool_cards(self.tools)
            return
        filtered = [
            t
            for t in self.tools
            if query in t["name"].lower()
            or query in t["description"].lower()
            or query in t.get("category", "").lower()
        ]
        self._render_tool_cards(filtered)

    def _render_tool_cards(self, tools: list[dict]) -> None:
        for child in self.tool_grid.winfo_children():
            child.destroy()

        if not tools:
            ttk.Label(self.tool_grid, text="No tools match your search.", style="Subtitle.TLabel").grid(
                row=0, column=0, padx=8, pady=20
            )
            return

        columns = 2
        for index, tool in enumerate(tools):
            row, col = divmod(index, columns)
            card = self._make_tool_card(self.tool_grid, tool)
            card.grid(row=row, column=col, sticky="nsew", padx=8, pady=8)
            self.tool_grid.columnconfigure(col, weight=1)

    def _make_tool_card(self, parent: ttk.Frame, tool: dict) -> ttk.Frame:
        outer = tk.Frame(parent, bg=COLORS["border"], padx=1, pady=1)
        card = tk.Frame(outer, bg=COLORS["card"], padx=18, pady=16)
        card.pack(fill=tk.BOTH, expand=True)

        title = tk.Label(card, text=tool["name"], bg=COLORS["card"], fg=COLORS["text"], font=("Segoe UI", 13, "bold"))
        title.pack(anchor="w")

        desc = tk.Label(
            card,
            text=tool["description"],
            bg=COLORS["card"],
            fg=COLORS["muted"],
            font=("Segoe UI", 10),
            wraplength=360,
            justify=tk.LEFT,
        )
        desc.pack(anchor="w", pady=(6, 10))

        meta = tk.Frame(card, bg=COLORS["card"])
        meta.pack(anchor="w", pady=(0, 12))
        for label in (
            tool.get("category", "General"),
            f"{len(tool['steps'])} steps",
            f"~{tool.get('estimatedMinutes', '?')} min",
        ):
            tk.Label(
                meta,
                text=label,
                bg=COLORS["accent_soft"],
                fg=COLORS["accent"],
                font=("Segoe UI", 8, "bold"),
                padx=8,
                pady=3,
            ).pack(side=tk.LEFT, padx=(0, 6))

        btn = tk.Button(
            card,
            text="Open tutorial →",
            bg=COLORS["accent"],
            fg="white",
            activebackground=COLORS["accent_dark"],
            activeforeground="white",
            relief=tk.FLAT,
            font=("Segoe UI", 10, "bold"),
            padx=12,
            pady=8,
            cursor="hand2",
            command=lambda t=tool: self._open_tool(t),
        )
        btn.pack(anchor="w")

        for widget in (card, title, desc, meta, btn):
            widget.bind("<Enter>", lambda _e, c=card: c.configure(bg=COLORS["card_hover"]))
            widget.bind("<Leave>", lambda _e, c=card: c.configure(bg=COLORS["card"]))

        return outer

    def _show_home(self) -> None:
        self.tutorial_frame.pack_forget()
        self.home_frame.pack(fill=tk.BOTH, expand=True)

    def _open_tool(self, tool: dict) -> None:
        self.current_tool = tool
        self.current_step = 0
        self.home_frame.pack_forget()
        self.tutorial_frame.pack(fill=tk.BOTH, expand=True)

        self.tool_title.configure(text=tool["name"])
        self.tool_desc.configure(text=tool["description"])

        self.step_listbox.delete(0, tk.END)
        for i, step in enumerate(tool["steps"], start=1):
            self.step_listbox.insert(tk.END, f"  {i}. {step['title']}")

        self._show_step(0)

    def _show_step(self, index: int) -> None:
        if not self.current_tool:
            return

        steps = self.current_tool["steps"]
        index = max(0, min(index, len(steps) - 1))
        self.current_step = index

        self.step_listbox.selection_clear(0, tk.END)
        self.step_listbox.selection_set(index)
        self.step_listbox.see(index)

        step = steps[index]
        self.step_heading.configure(text=f"Step {index + 1}: {step['title']}")
        self.step_body.configure(text=step["description"])
        self.progress_label.configure(text=f"{index + 1}/{len(steps)}")

        tips = step.get("tips") or []
        if tips:
            self.tips_frame.pack(anchor="w", fill=tk.X, pady=(0, 8))
            self.tips_body.configure(text="TIPS\n" + "\n".join(f"• {tip}" for tip in tips))
        else:
            self.tips_frame.pack_forget()
            self.tips_body.configure(text="")

        self.prev_btn.state(["!disabled"] if index > 0 else ["disabled"])
        self.next_btn.configure(text="Finish ✓" if index == len(steps) - 1 else "Next →")

    def _on_step_selected(self, _event: tk.Event | None = None) -> None:
        selection = self.step_listbox.curselection()
        if selection:
            self._show_step(selection[0])

    def _prev_step(self) -> None:
        self._show_step(self.current_step - 1)

    def _next_step(self) -> None:
        if not self.current_tool:
            return
        if self.current_step >= len(self.current_tool["steps"]) - 1:
            messagebox.showinfo("Tutorial complete", f"You finished the {self.current_tool['name']} tutorial.")
            return
        self._show_step(self.current_step + 1)

    def _open_current_ppt(self) -> None:
        if not self.current_tool:
            return
        rel = self.current_tool.get("pptFile", "")
        if not rel:
            messagebox.showinfo("No presentation", "No PPT file is linked for this tool.")
            return
        open_file(APP_DIR / rel)

    def destroy(self) -> None:
        if hasattr(self, "canvas"):
            self.canvas.unbind_all("<MouseWheel>")
            self.canvas.unbind_all("<Button-4>")
            self.canvas.unbind_all("<Button-5>")
        super().destroy()


def main() -> None:
    if not TOOLS_FILE.exists():
        messagebox.showerror("Missing data", f"tools.json not found at:\n{TOOLS_FILE}")
        sys.exit(1)
    app = ToolTutorialApp()
    app.mainloop()


if __name__ == "__main__":
    main()
