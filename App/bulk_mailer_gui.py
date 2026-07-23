"""
Bulk Mailer — Simple GUI (Outlook COM Automation)
===================================================
A point-and-click tool for sending personalised bulk emails through the
locally installed Outlook desktop application. Designed for non-technical
users — no code editing required.

Requirements:
    pip install pywin32 pandas openpyxl
    (see setup.bat)

Usage:
    Double-click "Send Bulk Mail.bat", or run:  python bulk_mailer_gui.py
"""

import os
import re
import sys
import glob
import html
import queue
import threading
import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk, filedialog, messagebox, scrolledtext, colorchooser
from datetime import datetime

try:
    import pandas as pd
except ImportError:
    messagebox.showerror(
        "Missing component",
        "The 'pandas' package is not installed.\n\nPlease run setup.bat in this folder, then try again.",
    )
    sys.exit(1)

try:
    import win32com.client
    import pythoncom
except ImportError:
    messagebox.showerror(
        "Missing component",
        "The 'pywin32' package is not installed.\n\nPlease run setup.bat in this folder, then try again.",
    )
    sys.exit(1)


LOG_FILE_GLOB = "sent_log_*.txt"
CUSTOM_FIELDS_SHEET_NAME = "Custom Fields"


def guess_email_column(columns: list) -> str:
    return next(
        (c for c in columns if c.strip().lower() in ("email", "email id", "email address")),
        columns[0] if columns else "",
    )

# Merge fields use <ColumnName> rather than {ColumnName} — easier for
# non-technical users to recognise and type, and avoids confusion with
# curly braces which have no special meaning elsewhere in this tool.
PLACEHOLDER_RE = re.compile(r"<([^<>]+)>")


def render_template(template: str, fields: dict) -> str:
    def replace(match):
        key = match.group(1).strip()
        return fields.get(key, f"<MISSING: {key}>")
    return PLACEHOLDER_RE.sub(replace, template)


def format_cc_list(raw: str) -> str:
    """Turn a CC cell containing one or more comma-separated addresses into
    the semicolon-separated form Outlook expects (Outlook does not reliably
    split on commas)."""
    if not raw:
        return ""
    addresses = [addr.strip() for addr in raw.split(",")]
    return "; ".join(addr for addr in addresses if addr)


# Inline images/tables are marked in the plain-text body with a square-bracket
# token (not angle brackets, so it survives html.escape() untouched and can't
# be confused with a <MergeField> placeholder).
PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001E"

# Rich-text style tags recognised when converting the editor's Text widget
# content to HTML. "underline" carries no font info, so it never conflicts
# with the "bold"/"italic" tags' font objects.
STYLE_TAGS = ("bold", "italic", "underline")

# Alignment is a per-paragraph (line) property, not a per-character run, so
# it is tracked separately via dump_line_alignments rather than via runs.
ALIGN_TAGS = ("align_center", "align_right")


def dump_runs(text_widget) -> list:
    """Serialise a Tk Text widget's content + tags into a plain, thread-safe
    list of (frozenset_of_tags, text_chunk) — must be called on the main/UI
    thread, since Tk widgets cannot be touched from a background thread."""
    active_tags = set()
    runs = []
    for key, value, _index in text_widget.dump("1.0", "end-1c", tag=True, text=True):
        if key == "tagon":
            active_tags.add(value)
        elif key == "tagoff":
            active_tags.discard(value)
        elif key == "text":
            runs.append((frozenset(active_tags), value))
    return runs


def dump_line_alignments(text_widget) -> list:
    """Return a list where index i holds the alignment ('left'/'center'/
    'right') of line i+1 — must be called on the main/UI thread, same
    reasoning as dump_runs."""
    last_line = int(text_widget.index("end-1c").split(".")[0])
    alignments = []
    for line_num in range(1, last_line + 1):
        tags_here = text_widget.tag_names(f"{line_num}.0")
        if "align_center" in tags_here:
            alignments.append("center")
        elif "align_right" in tags_here:
            alignments.append("right")
        else:
            alignments.append("left")
    return alignments


def runs_to_plain_text(runs: list, fields: dict) -> str:
    """Merge-field-rendered plain text, ignoring all formatting — used for
    the empty-body check and the plain-text draft preview."""
    return "".join(render_template(chunk, fields) for _, chunk in runs)


def render_rich_email(runs: list, fields: dict, inline_images: list, inline_tables: list, line_alignments: list = None):
    """Return (body_format, body_content) for a MailItem, built from the
    editor's formatted runs plus any inline images/tables. body_format is
    1 (plain text) if nothing but plain text was used, or 2 (HTML) once any
    formatting, alignment, image, or table is present."""
    line_alignments = line_alignments or []
    has_alignment = any(a != "left" for a in line_alignments)
    has_formatting = any(
        tags & set(STYLE_TAGS) or any(t.startswith(("color_", "bgcolor_")) for t in tags)
        for tags, _ in runs
    )
    if not has_formatting and not has_alignment and not inline_images and not inline_tables:
        return 1, runs_to_plain_text(runs, fields)

    parts = []
    for tags, chunk in runs:
        rendered = render_template(chunk, fields)
        escaped = html.escape(rendered)
        if "bold" in tags:
            escaped = f"<b>{escaped}</b>"
        if "italic" in tags:
            escaped = f"<i>{escaped}</i>"
        if "underline" in tags:
            escaped = f"<u>{escaped}</u>"
        color_tag = next((t for t in tags if t.startswith("color_")), None)
        bgcolor_tag = next((t for t in tags if t.startswith("bgcolor_")), None)
        if color_tag or bgcolor_tag:
            style = ""
            if color_tag:
                style += f"color:#{color_tag[len('color_'):]};"
            if bgcolor_tag:
                style += f"background-color:#{bgcolor_tag[len('bgcolor_'):]};"
            escaped = f'<span style="{style}">{escaped}</span>'
        parts.append(escaped)

    lines = "".join(parts).split("\n")
    aligned_lines = []
    for i, line in enumerate(lines):
        align = line_alignments[i] if i < len(line_alignments) else "left"
        if align != "left":
            aligned_lines.append(f'<div style="text-align:{align};">{line}</div>')
        else:
            aligned_lines.append(line)
    body = "<br>\n".join(aligned_lines)

    for image in inline_images:
        img_tag = f'<img src="cid:{image["cid"]}" style="max-width:100%;">'
        body = body.replace(image["token"], img_tag)

    for table in inline_tables:
        table_html = build_table_html(table["rows"], fields)
        body = body.replace(table["token"], table_html)

    return 2, f'<html><body style="font-family:Calibri,Arial,sans-serif;font-size:11pt;">{body}</body></html>'


def build_table_html(rows: list, fields: dict) -> str:
    row_html = []
    for row in rows:
        cells = "".join(
            f'<td style="border:1px solid #999;padding:4px 8px;">{html.escape(render_template(cell, fields))}</td>'
            for cell in row
        )
        row_html.append(f"<tr>{cells}</tr>")
    return f'<table style="border-collapse:collapse;margin:6px 0;">{"".join(row_html)}</table>'


def load_already_sent() -> set:
    sent = set()
    for log_file in glob.glob(LOG_FILE_GLOB):
        with open(log_file, encoding="utf-8") as f:
            for line in f:
                match = re.match(r"\[SENT\]\s+(\S+@\S+)", line)
                if match:
                    sent.add(match.group(1).lower())
    return sent


class BulkMailerApp:
    def __init__(self, root):
        self.root = root
        root.title("Bulk Mailer")
        root.geometry("920x820")
        root.minsize(800, 620)

        self.df = None
        self.columns = []
        self.outlook = None
        self.accounts = []
        self.attachments = []
        self.inline_images = []
        self.inline_tables = []
        self._image_counter = 0
        self._table_counter = 0
        self._color_tag_counter = 0
        self.last_focus_widget = None
        self.stop_event = threading.Event()
        self.send_thread = None
        self.msg_queue = queue.Queue()

        self._setup_styles()
        self._build_scroll_container()
        self._build_ui()
        self._set_recipients_dependent_state(False)
        self._update_highlight()
        self.root.after(100, self._poll_queue)

    def _setup_styles(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("Highlight.TButton", background="#2e7d32", foreground="white", font=("Segoe UI", 10, "bold"))
        style.map("Highlight.TButton", background=[("active", "#1b5e20")], foreground=[("disabled", "#cccccc")])

    # ── Scrollable page (so instructions + all steps stay reachable
    #    even on smaller screens) ────────────────────────────────────

    def _build_scroll_container(self):
        outer = ttk.Frame(self.root)
        outer.pack(fill="both", expand=True)

        canvas = tk.Canvas(outer, highlightthickness=0)
        scrollbar = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        self.page = ttk.Frame(canvas)

        self.page.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.page, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

    # ── UI construction ─────────────────────────────────────────────

    def _build_ui(self):
        pad = {"padx": 8, "pady": 6}
        parent = self.page

        # --- Always-visible instructions ---
        instructions_frame = ttk.LabelFrame(parent, text="How to use this tool")
        instructions_frame.pack(fill="x", **pad)
        instructions_text = (
            "Follow Steps 1 to 6 below in order. A step's controls stay greyed out until the step "
            "before it is done, and the green button is always the one to click next.\n\n"
            "STEP 1 — Choose your Excel file and click 'Load Recipients'. It must contain a "
            "'Recipients' sheet (Email, and CC if needed) and, for merge fields, a "
            "'Custom Fields' sheet (Email plus one column per field — see recipients_template.xlsx).\n"
            "STEP 2 — Select which column holds the email address, and the CC column if you have one. "
            "A CC cell can be left blank, hold one address, or hold several separated by commas "
            "(e.g. a@x.com, b@x.com).\n"
            "STEP 3 — Click a field button (e.g. <Name>) to insert it into the Subject or Message body, "
            "at the spot you last clicked.\n"
            "STEP 4 — Write your Subject and Message. Choose 'Simple text email' for a plain message, "
            "or 'Rich email' to reveal a toolbar for Bold/Italic/Underline, text color, highlight "
            "color, left/center/right alignment, bullet lists, tables, and inline images. Switching "
            "back to Simple removes any formatting/images/tables already added, so a simple send can "
            "never accidentally carry stray extras.\n"
            "STEP 5 — Attach any files you want sent with every email (optional).\n"
            "STEP 6 — In Outlook, set the mailbox you want to send from as your DEFAULT account "
            "(File > Account Settings > Account Settings > select the account > Set as Default) "
            "before sending — some Outlook/Exchange setups only send correctly from the default "
            "account, regardless of which one you pick here. Then click 'Connect to Outlook' and "
            "choose the matching account.\n"
            "FINALLY — Click 'Send Bulk Mail', review the draft preview for the first recipient, "
            "then confirm. You can click 'Stop' at any time — if you restart later, recipients "
            "already sent to will automatically be skipped.\n"
            "Note: fields like <Name> or <Company> come from the 'Custom Fields' sheet's column headings, "
            "matched to each recipient by their Email address."
        )
        ttk.Label(instructions_frame, text=instructions_text, justify="left", wraplength=860).pack(
            anchor="w", padx=8, pady=8
        )

        # --- Always-visible "what to do next" banner ---
        self.next_step_var = tk.StringVar()
        next_step_banner = ttk.Label(
            parent, textvariable=self.next_step_var, font=("Segoe UI", 11, "bold"),
            foreground="#1b5e20", wraplength=880, justify="left",
        )
        next_step_banner.pack(anchor="w", padx=16, pady=(0, 10))

        # --- Recipients file section ---
        file_frame = ttk.LabelFrame(parent, text="Step 1 — Recipient list (Excel file)")
        file_frame.pack(fill="x", **pad)

        self.file_path_var = tk.StringVar(value=os.path.join(os.getcwd(), "recipients.xlsx"))
        ttk.Entry(file_frame, textvariable=self.file_path_var).pack(side="left", fill="x", expand=True, padx=(8, 4), pady=6)
        ttk.Button(file_frame, text="Browse...", command=self._browse_file).pack(side="left", padx=4, pady=6)

        ttk.Label(file_frame, text="Recipients sheet:").pack(side="left", padx=(12, 4))
        self.sheet_name_var = tk.StringVar(value="Recipients")
        ttk.Entry(file_frame, textvariable=self.sheet_name_var, width=14).pack(side="left", padx=4)

        self.load_button = ttk.Button(file_frame, text="Load Recipients", command=self._load_recipients)
        self.load_button.pack(side="left", padx=8)

        self.recipients_status_var = tk.StringVar(value="No file loaded yet.")
        ttk.Label(parent, textvariable=self.recipients_status_var, foreground="#555").pack(anchor="w", padx=16)

        # --- Column mapping ---
        self.map_frame = ttk.LabelFrame(parent, text="Step 2 — Which columns are the email addresses?")
        self.map_frame.pack(fill="x", **pad)

        ttk.Label(self.map_frame, text="Send To column:").grid(row=0, column=0, padx=8, pady=6, sticky="w")
        self.to_column_var = tk.StringVar()
        self.to_column_combo = ttk.Combobox(self.map_frame, textvariable=self.to_column_var, state="disabled", width=25)
        self.to_column_combo.grid(row=0, column=1, padx=8, pady=6, sticky="w")

        ttk.Label(self.map_frame, text="CC column (optional, comma-separate for multiple):").grid(row=0, column=2, padx=8, pady=6, sticky="w")
        self.cc_column_var = tk.StringVar()
        self.cc_column_combo = ttk.Combobox(self.map_frame, textvariable=self.cc_column_var, state="disabled", width=25)
        self.cc_column_combo.grid(row=0, column=3, padx=8, pady=6, sticky="w")

        # --- Merge fields ---
        merge_frame = ttk.LabelFrame(parent, text="Step 3 — Available merge fields (click to insert at cursor)")
        merge_frame.pack(fill="x", **pad)
        self.merge_buttons_frame = ttk.Frame(merge_frame)
        self.merge_buttons_frame.pack(fill="x", padx=8, pady=6)
        self.merge_hint_label = ttk.Label(merge_frame, text="Load recipients in Step 1 to see available fields.", foreground="#555")
        self.merge_hint_label.pack(anchor="w", padx=8, pady=(0, 6))

        # --- Subject / Body ---
        self.content_frame = ttk.LabelFrame(parent, text="Step 4 — Email content")
        self.content_frame.pack(fill="both", expand=True, **pad)

        ttk.Label(self.content_frame, text="Subject:").pack(anchor="w", padx=8, pady=(6, 0))
        self.subject_var = tk.StringVar(value="Subject line goes here, e.g. Dear <Name>")
        self.subject_entry = ttk.Entry(self.content_frame, textvariable=self.subject_var, state="disabled")
        self.subject_entry.pack(fill="x", padx=8, pady=(0, 6))
        self.subject_entry.bind("<FocusIn>", lambda e: self._set_last_focus(self.subject_entry))

        # --- Email type switch: keeps the two workflows visually and
        # functionally separate, so a "simple" send can never accidentally
        # carry stray formatting, an orphaned image, or a table. ---
        mode_frame = ttk.Frame(self.content_frame)
        mode_frame.pack(fill="x", padx=8, pady=(4, 4))
        ttk.Label(mode_frame, text="Email type:").pack(side="left", padx=(0, 8))
        self.email_mode_var = tk.StringVar(value="simple")
        self.simple_mode_radio = ttk.Radiobutton(
            mode_frame, text="Simple text email", value="simple",
            variable=self.email_mode_var, command=self._on_email_mode_changed,
        )
        self.simple_mode_radio.pack(side="left", padx=(0, 12))
        self.rich_mode_radio = ttk.Radiobutton(
            mode_frame, text="Rich email (formatting, images, tables)", value="rich",
            variable=self.email_mode_var, command=self._on_email_mode_changed,
        )
        self.rich_mode_radio.pack(side="left")

        # --- Formatting toolbar (Rich mode only) ---
        self.toolbar_frame = ttk.Frame(self.content_frame)
        toolbar_row1 = ttk.Frame(self.toolbar_frame)
        toolbar_row1.pack(fill="x")
        ttk.Button(toolbar_row1, text="B", width=3, command=lambda: self._toggle_style("bold")).pack(side="left", padx=2)
        ttk.Button(toolbar_row1, text="I", width=3, command=lambda: self._toggle_style("italic")).pack(side="left", padx=2)
        ttk.Button(toolbar_row1, text="U", width=3, command=lambda: self._toggle_style("underline")).pack(side="left", padx=2)
        ttk.Button(toolbar_row1, text="Text Color...", command=self._choose_color).pack(side="left", padx=(8, 2))
        ttk.Button(toolbar_row1, text="Highlight Color...", command=self._choose_bg_color).pack(side="left", padx=2)
        ttk.Button(toolbar_row1, text="• Bullet List", command=self._toggle_bullet).pack(side="left", padx=(8, 2))
        ttk.Button(toolbar_row1, text="Insert Table...", command=self._insert_table).pack(side="left", padx=(8, 2))
        self.insert_image_button = ttk.Button(toolbar_row1, text="Insert Image...", command=self._insert_image)
        self.insert_image_button.pack(side="left", padx=(8, 2))

        toolbar_row2 = ttk.Frame(self.toolbar_frame)
        toolbar_row2.pack(fill="x", pady=(4, 0))
        ttk.Label(toolbar_row2, text="Align:").pack(side="left", padx=(2, 4))
        ttk.Button(toolbar_row2, text="Left", command=lambda: self._set_alignment("left")).pack(side="left", padx=2)
        ttk.Button(toolbar_row2, text="Center", command=lambda: self._set_alignment("center")).pack(side="left", padx=2)
        ttk.Button(toolbar_row2, text="Right", command=lambda: self._set_alignment("right")).pack(side="left", padx=2)

        ttk.Label(self.content_frame, text="Message body:").pack(anchor="w", padx=8)
        self.body_text = scrolledtext.ScrolledText(self.content_frame, wrap="word", height=12)
        self.body_text.pack(fill="both", expand=True, padx=8, pady=(0, 4))
        self.body_text.insert("1.0", "Dear <Name>,\n\nType your message here.\n\nBest regards,\nYour Company Name")
        self.body_text.bind("<FocusIn>", lambda e: self._set_last_focus(self.body_text))
        self.body_text.config(state="disabled")

        base_font = tkfont.nametofont(self.body_text.cget("font"))
        bold_font = base_font.copy(); bold_font.configure(weight="bold")
        italic_font = base_font.copy(); italic_font.configure(slant="italic")
        self.body_text.tag_configure("bold", font=bold_font)
        self.body_text.tag_configure("italic", font=italic_font)
        self.body_text.tag_configure("underline", underline=True)
        self.body_text.tag_configure("align_center", justify="center")
        self.body_text.tag_configure("align_right", justify="right")

        self.rich_hint_label = ttk.Label(
            self.content_frame,
            text="Images/tables are inserted as [[IMAGE:n]] / [[TABLE:n]] markers at your cursor and "
                 "rendered in place in the email. Do not edit or delete marker text by hand. Combined "
                 "Bold+Italic on the exact same text may not preview perfectly on screen but still "
                 "sends correctly.",
            foreground="#555", wraplength=860, justify="left",
        )

        self._apply_email_mode_visibility()

        # --- Attachments ---
        attach_frame = ttk.LabelFrame(parent, text="Step 5 — Attachments (optional, sent with every email)")
        attach_frame.pack(fill="x", **pad)

        attach_list_frame = ttk.Frame(attach_frame)
        attach_list_frame.pack(fill="x", padx=8, pady=6)
        self.attachment_listbox = tk.Listbox(attach_list_frame, height=4)
        self.attachment_listbox.pack(side="left", fill="x", expand=True)
        attach_scrollbar = ttk.Scrollbar(attach_list_frame, orient="vertical", command=self.attachment_listbox.yview)
        attach_scrollbar.pack(side="left", fill="y")
        self.attachment_listbox.config(yscrollcommand=attach_scrollbar.set)

        attach_button_frame = ttk.Frame(attach_frame)
        attach_button_frame.pack(fill="x", padx=8, pady=(0, 8))
        ttk.Button(attach_button_frame, text="Add Attachment...", command=self._add_attachments).pack(side="left", padx=(0, 8))
        ttk.Button(attach_button_frame, text="Remove Selected", command=self._remove_attachment).pack(side="left")

        # --- Outlook account + delay ---
        send_setup_frame = ttk.LabelFrame(parent, text="Step 6 — Sending options")
        send_setup_frame.pack(fill="x", **pad)

        warning_text = (
            "⚠ IMPORTANT: even if you select a different account below, on some Outlook/Exchange "
            "setups emails will still actually be sent FROM whichever account is set as your "
            "DEFAULT account in Outlook — the dropdown selection is not always honoured. Before "
            "connecting, go to Outlook > File > Account Settings > Account Settings > select the "
            "mailbox you want to send from > Set as Default.\n"
            "⚠ IMPORTANT: emails will be sent using ONLY the Excel file currently loaded in Step 1. "
            "If you change or update that file, click 'Load Recipients' again in Step 1 before "
            "sending, or the tool will keep using the earlier data."
        )
        ttk.Label(
            send_setup_frame, text=warning_text, justify="left", wraplength=860,
            foreground="#b45309", font=("Segoe UI", 9, "bold"),
        ).grid(row=0, column=0, columnspan=5, padx=8, pady=(8, 4), sticky="w")

        ttk.Label(send_setup_frame, text="Send from account:").grid(row=1, column=0, padx=8, pady=6, sticky="w")
        self.account_var = tk.StringVar()
        self.account_combo = ttk.Combobox(send_setup_frame, textvariable=self.account_var, state="readonly", width=35)
        self.account_combo.grid(row=1, column=1, padx=8, pady=6, sticky="w")
        self.account_combo.bind("<<ComboboxSelected>>", lambda e: self._update_highlight())
        self.connect_button = ttk.Button(send_setup_frame, text="Connect to Outlook", command=self._connect_outlook)
        self.connect_button.grid(row=1, column=2, padx=8, pady=6)

        ttk.Label(send_setup_frame, text="Delay between emails (seconds):").grid(row=1, column=3, padx=(20, 4), pady=6, sticky="w")
        self.delay_var = tk.DoubleVar(value=2.0)
        ttk.Spinbox(send_setup_frame, from_=0.5, to=10.0, increment=0.5, textvariable=self.delay_var, width=6).grid(row=1, column=4, padx=4, pady=6)

        self.confirm_var = tk.BooleanVar(value=False)
        confirm_check = ttk.Checkbutton(
            send_setup_frame,
            text="I confirm the data file loaded in Step 1 is up to date, and my DEFAULT account "
                 "in Outlook is the one I want to send from.",
            variable=self.confirm_var, command=self._update_highlight,
        )
        confirm_check.grid(row=2, column=0, columnspan=5, padx=8, pady=(4, 8), sticky="w")

        # --- Send / Stop / Progress ---
        action_frame = ttk.Frame(parent)
        action_frame.pack(fill="x", **pad)

        self.send_button = ttk.Button(action_frame, text="Send Bulk Mail", command=self._on_send_clicked)
        self.send_button.pack(side="left", padx=8)
        self.stop_button = ttk.Button(action_frame, text="Stop", command=self._on_stop_clicked, state="disabled")
        self.stop_button.pack(side="left", padx=8)

        self.progress_var = tk.StringVar(value="")
        ttk.Label(action_frame, textvariable=self.progress_var).pack(side="left", padx=16)

        self.progress_bar = ttk.Progressbar(action_frame, mode="determinate")
        self.progress_bar.pack(side="left", fill="x", expand=True, padx=8)

        # --- Log ---
        log_frame = ttk.LabelFrame(parent, text="Send log")
        log_frame.pack(fill="both", expand=True, **pad)
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap="word", height=8, state="disabled")
        self.log_text.pack(fill="both", expand=True, padx=8, pady=8)

    def _set_last_focus(self, widget):
        self.last_focus_widget = widget

    # ── Step gating (grey out steps that can't be done yet) ─────────

    def _set_recipients_dependent_state(self, enabled: bool):
        state = "readonly" if enabled else "disabled"
        self.to_column_combo.config(state=state)
        self.cc_column_combo.config(state=state)
        self.subject_entry.config(state=("normal" if enabled else "disabled"))
        self.body_text.config(state=("normal" if enabled else "disabled"))
        widget_state = "normal" if enabled else "disabled"
        for row in self.toolbar_frame.winfo_children():
            for child in row.winfo_children():
                child.config(state=widget_state)
        self.simple_mode_radio.config(state=widget_state)
        self.rich_mode_radio.config(state=widget_state)

    def _update_highlight(self):
        """Highlight the single button the user should click next, and spell
        out that same instruction in the always-visible banner above Step 1."""
        if self.df is None:
            target = self.load_button
            message = "➡ Next: choose your Excel file in Step 1 and click 'Load Recipients'."
        elif not self.outlook:
            target = self.connect_button
            message = "➡ Next: fill in Steps 2-5 as needed, then click 'Connect to Outlook' in Step 6."
        elif not self.account_var.get():
            target = self.connect_button
            message = "➡ Next: choose which account to send from in Step 6."
        elif not self.confirm_var.get():
            target = None
            message = "➡ Next: in Step 6, tick the checkbox confirming your data file and default Outlook account are correct."
        else:
            target = self.send_button
            message = "➡ Next: check your Subject and Message, then click 'Send Bulk Mail' below."
        for button in (self.load_button, self.connect_button, self.send_button):
            button.configure(style="Highlight.TButton" if button is target else "TButton")
        self.send_button.config(state=("normal" if target is self.send_button else "disabled"))
        self.next_step_var.set(message)

    # ── Recipients loading ──────────────────────────────────────────

    def _browse_file(self):
        path = filedialog.askopenfilename(
            title="Choose recipient list",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if path:
            self.file_path_var.set(path)

    def _load_recipients(self):
        path = self.file_path_var.get().strip()
        sheet = self.sheet_name_var.get().strip() or 0
        if not os.path.exists(path):
            messagebox.showerror("File not found", f"Could not find:\n{path}")
            return
        try:
            df = pd.read_excel(path, sheet_name=sheet, dtype=str)
        except Exception as exc:
            messagebox.showerror("Could not read file", str(exc))
            return

        df.columns = df.columns.str.strip()
        recipient_columns = list(df.columns)
        join_col = guess_email_column(recipient_columns)

        merge_field_columns = []
        custom_status = "No 'Custom Fields' sheet found — merge fields are not available."
        try:
            custom_df = pd.read_excel(path, sheet_name=CUSTOM_FIELDS_SHEET_NAME, dtype=str)
            custom_df.columns = custom_df.columns.str.strip()
        except Exception:
            custom_df = None

        if custom_df is not None:
            custom_join_col = guess_email_column(list(custom_df.columns))
            if join_col and custom_join_col and join_col in df.columns:
                custom_df[custom_join_col] = custom_df[custom_join_col].str.strip().str.lower()
                df["_join_key"] = df[join_col].str.strip().str.lower()
                df = df.merge(
                    custom_df.rename(columns={custom_join_col: "_join_key"}),
                    on="_join_key", how="left",
                ).drop(columns=["_join_key"])
                merge_field_columns = [c for c in custom_df.columns if c != custom_join_col]
                custom_status = f"Merge fields loaded from 'Custom Fields' sheet: {len(merge_field_columns)} field(s)."
            else:
                custom_status = "Found 'Custom Fields' sheet, but could not find a matching Email column in both sheets."

        self.df = df
        self.columns = recipient_columns  # To/CC choices come from the Recipients sheet only

        self.to_column_combo["values"] = self.columns
        self.cc_column_combo["values"] = [""] + self.columns

        cc_guess = next((c for c in self.columns if c.strip().lower() == "cc"), "")
        self.to_column_var.set(join_col)
        self.cc_column_var.set(cc_guess)

        self.recipients_status_var.set(
            f"Loaded {len(df)} recipient(s) from '{os.path.basename(path)}', sheet '{sheet}'. {custom_status}"
        )

        for widget in self.merge_buttons_frame.winfo_children():
            widget.destroy()
        for col in merge_field_columns:
            ttk.Button(
                self.merge_buttons_frame, text=f"<{col}>",
                command=lambda c=col: self._insert_merge_field(c),
            ).pack(side="left", padx=3, pady=2)
        if merge_field_columns:
            self.merge_hint_label.config(text="Click any field above to insert it into the Subject or Message body (wherever you clicked last).")
        else:
            self.merge_hint_label.config(text=custom_status)

        self._set_recipients_dependent_state(True)
        self.confirm_var.set(False)  # any (re)load requires re-confirming the data file is correct
        self._update_highlight()

    def _insert_merge_field(self, column_name):
        token = f"<{column_name}>"
        widget = self.last_focus_widget or self.subject_entry
        widget.insert(tk.INSERT, token)

    # ── Email type switch (Simple vs Rich) ──────────────────────────

    def _mode_has_rich_content(self) -> bool:
        has_style_tags = any(self.body_text.tag_ranges(t) for t in STYLE_TAGS + ALIGN_TAGS)
        has_color_tags = any(
            t.startswith(("color_", "bgcolor_")) and self.body_text.tag_ranges(t)
            for t in self.body_text.tag_names()
        )
        return has_style_tags or has_color_tags or bool(self.inline_images) or bool(self.inline_tables)

    def _clear_all_rich_content(self):
        for tag in list(self.body_text.tag_names()):
            if tag in STYLE_TAGS or tag in ALIGN_TAGS or tag.startswith(("color_", "bgcolor_")):
                self.body_text.tag_remove(tag, "1.0", "end")
        self.inline_images.clear()
        self.inline_tables.clear()

    def _apply_email_mode_visibility(self):
        if self.email_mode_var.get() == "rich":
            self.toolbar_frame.pack(fill="x", padx=8, pady=(2, 4), before=self.body_text)
            self.rich_hint_label.pack(anchor="w", padx=8, pady=(0, 8), after=self.body_text)
        else:
            self.toolbar_frame.pack_forget()
            self.rich_hint_label.pack_forget()

    def _on_email_mode_changed(self):
        if self.email_mode_var.get() == "simple" and self._mode_has_rich_content():
            proceed = messagebox.askyesno(
                "Switch to Simple email?",
                "Switching to a Simple text email will remove all formatting (bold/italic/underline/"
                "colors), inserted images, and tables from your message. Continue?",
            )
            if not proceed:
                self.email_mode_var.set("rich")
                return
            self._clear_all_rich_content()
        self._apply_email_mode_visibility()

    # ── Rich-text formatting ─────────────────────────────────────────

    def _toggle_style(self, tag_name: str):
        try:
            start, end = self.body_text.index("sel.first"), self.body_text.index("sel.last")
        except tk.TclError:
            messagebox.showinfo("Select text first", "Select some text in the message body, then click this button.")
            return
        if tag_name in self.body_text.tag_names(start):
            self.body_text.tag_remove(tag_name, start, end)
        else:
            self.body_text.tag_add(tag_name, start, end)

    def _choose_color(self):
        self._apply_color_choice(prefix="color_", title="Choose text color", tk_option="foreground")

    def _choose_bg_color(self):
        self._apply_color_choice(prefix="bgcolor_", title="Choose highlight color", tk_option="background")

    def _apply_color_choice(self, prefix: str, title: str, tk_option: str):
        try:
            start, end = self.body_text.index("sel.first"), self.body_text.index("sel.last")
        except tk.TclError:
            messagebox.showinfo("Select text first", "Select some text in the message body, then click this button.")
            return
        _, hex_color = colorchooser.askcolor(title=title)
        if not hex_color:
            return
        for tag in self.body_text.tag_names(start):
            if tag.startswith(prefix):
                self.body_text.tag_remove(tag, start, end)
        self._color_tag_counter += 1
        tag_name = f"{prefix}{hex_color.lstrip('#')}"
        self.body_text.tag_configure(tag_name, **{tk_option: hex_color})
        self.body_text.tag_add(tag_name, start, end)

    def _toggle_bullet(self):
        try:
            start_line = int(self.body_text.index("sel.first").split(".")[0])
            end_line = int(self.body_text.index("sel.last").split(".")[0])
        except tk.TclError:
            start_line = end_line = int(self.body_text.index("insert").split(".")[0])

        for line_num in range(start_line, end_line + 1):
            line_start = f"{line_num}.0"
            line_text = self.body_text.get(line_start, f"{line_num}.end")
            if line_text.startswith("• "):
                self.body_text.delete(line_start, f"{line_num}.2")
            else:
                self.body_text.insert(line_start, "• ")

    def _set_alignment(self, align_name: str):
        try:
            start_line = int(self.body_text.index("sel.first").split(".")[0])
            end_line = int(self.body_text.index("sel.last").split(".")[0])
        except tk.TclError:
            start_line = end_line = int(self.body_text.index("insert").split(".")[0])

        for line_num in range(start_line, end_line + 1):
            line_start, line_end = f"{line_num}.0", f"{line_num}.end"
            for tag in ALIGN_TAGS:
                self.body_text.tag_remove(tag, line_start, line_end)
            if align_name != "left":
                self.body_text.tag_add(f"align_{align_name}", line_start, line_end)

    # ── Inline tables ────────────────────────────────────────────────

    def _insert_table(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Insert Table")
        dialog.transient(self.root)
        dialog.grab_set()

        size_frame = ttk.Frame(dialog)
        size_frame.pack(fill="x", padx=12, pady=12)
        ttk.Label(size_frame, text="Rows:").pack(side="left")
        rows_var = tk.IntVar(value=2)
        ttk.Spinbox(size_frame, from_=1, to=10, textvariable=rows_var, width=5).pack(side="left", padx=(4, 16))
        ttk.Label(size_frame, text="Columns:").pack(side="left")
        cols_var = tk.IntVar(value=2)
        ttk.Spinbox(size_frame, from_=1, to=10, textvariable=cols_var, width=5).pack(side="left", padx=(4, 0))

        grid_frame = ttk.Frame(dialog)
        grid_frame.pack(padx=12, pady=(0, 12))
        cell_entries = []

        def rebuild_grid():
            for widget in grid_frame.winfo_children():
                widget.destroy()
            cell_entries.clear()
            for r in range(rows_var.get()):
                row_entries = []
                for c in range(cols_var.get()):
                    entry = ttk.Entry(grid_frame, width=16)
                    entry.grid(row=r, column=c, padx=2, pady=2)
                    row_entries.append(entry)
                cell_entries.append(row_entries)

        ttk.Button(size_frame, text="Set Size", command=rebuild_grid).pack(side="left", padx=(16, 0))
        rebuild_grid()

        ttk.Label(
            dialog, text="Tip: cells can use <MergeField> placeholders too.",
            foreground="#555",
        ).pack(anchor="w", padx=12, pady=(0, 4))

        button_frame = ttk.Frame(dialog)
        button_frame.pack(fill="x", padx=12, pady=(0, 12))

        def on_insert():
            rows = [[entry.get() for entry in row] for row in cell_entries]
            self._table_counter += 1
            index = self._table_counter
            token = f"[[TABLE:{index}]]"
            self.inline_tables.append({"token": token, "rows": rows})
            self.body_text.insert(tk.INSERT, token)
            dialog.destroy()

        ttk.Button(button_frame, text="Insert", command=on_insert).pack(side="right", padx=4)
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side="right", padx=4)

    # ── Inline images ────────────────────────────────────────────────

    def _insert_image(self):
        path = filedialog.askopenfilename(
            title="Choose an image to insert",
            filetypes=[("Image files", "*.png *.jpg *.jpeg *.gif *.bmp"), ("All files", "*.*")],
        )
        if not path:
            return
        self._image_counter += 1
        index = self._image_counter
        token = f"[[IMAGE:{index}]]"
        self.inline_images.append({"token": token, "path": path, "cid": f"inlineimage{index}"})
        # Images only make sense in the message body, not the one-line Subject.
        self.body_text.insert(tk.INSERT, token)

    # ── Attachments ──────────────────────────────────────────────────

    def _add_attachments(self):
        paths = filedialog.askopenfilenames(title="Choose file(s) to attach")
        for path in paths:
            if path not in self.attachments:
                self.attachments.append(path)
                self.attachment_listbox.insert("end", os.path.basename(path))

    def _remove_attachment(self):
        selected = list(self.attachment_listbox.curselection())
        for index in reversed(selected):
            self.attachment_listbox.delete(index)
            del self.attachments[index]

    # ── Outlook connection ───────────────────────────────────────────

    def _connect_outlook(self):
        try:
            self._log("Opening Outlook — please log in if prompted...")
            self.outlook = win32com.client.gencache.EnsureDispatch("Outlook.Application")
            namespace = self.outlook.GetNamespace("MAPI")
            namespace.Logon()
            accounts = self.outlook.Session.Accounts
            self.accounts = [accounts.Item(i) for i in range(1, accounts.Count + 1)]
            display = [a.SmtpAddress for a in self.accounts]
            self.account_combo["values"] = display
            if display:
                self.account_combo.current(0)
            self._log(f"Connected. Found {len(display)} account(s).")
            self._update_highlight()
        except Exception as exc:
            messagebox.showerror("Could not connect to Outlook", str(exc))

    # ── Sending ──────────────────────────────────────────────────────

    def _on_send_clicked(self):
        if self.df is None or self.df.empty:
            messagebox.showwarning("No recipients", "Please load a recipient list first.")
            return
        to_col = self.to_column_var.get()
        if not to_col:
            messagebox.showwarning("Missing column", "Please choose the 'Send To' column.")
            return
        if not self.outlook or not self.accounts:
            messagebox.showwarning("Not connected", "Please click 'Connect to Outlook' first.")
            return
        if not self.account_var.get():
            messagebox.showwarning("No account selected", "Please choose which account to send from.")
            return
        if not self.confirm_var.get():
            messagebox.showwarning(
                "Confirmation required",
                "Please tick the confirmation checkbox in Step 6 first — it confirms your data "
                "file and Outlook default account are correct.",
            )
            return
        # Runs and line alignments must be captured here, on the main/UI
        # thread — Tk widgets cannot be touched from the background send thread.
        runs = dump_runs(self.body_text)
        line_alignments = dump_line_alignments(self.body_text)
        plain_body = "".join(chunk for _, chunk in runs).strip()
        if not plain_body:
            messagebox.showwarning("Empty message", "Please enter a message body.")
            return

        df = self.df.dropna(subset=[to_col]).reset_index(drop=True)
        already_sent = load_already_sent()
        pending_mask = ~df[to_col].str.strip().str.lower().isin(already_sent)
        pending = df[pending_mask]

        if pending.empty:
            messagebox.showinfo("Nothing to send", "All recipients in this file have already been sent to (see existing sent_log files).")
            return

        # Only images/tables whose marker is still present in the body are
        # actually sent — if the user deleted a marker by hand, that item is
        # dropped rather than silently attached with nothing referencing it.
        active_images = [img for img in self.inline_images if img["token"] in plain_body]
        active_tables = [tbl for tbl in self.inline_tables if tbl["token"] in plain_body]

        missing_paths = [p for p in self.attachments if not os.path.exists(p)]
        missing_paths += [img["path"] for img in active_images if not os.path.exists(img["path"])]
        if missing_paths:
            messagebox.showerror(
                "File not found",
                "These attached or inserted files can no longer be found:\n" + "\n".join(missing_paths),
            )
            return

        subject_template = self.subject_var.get()
        cc_col = self.cc_column_var.get() or None
        attachments = list(self.attachments)

        if not self._show_preview_and_confirm(pending, to_col, cc_col, subject_template, runs, attachments, active_images, active_tables, len(df), len(pending)):
            return

        account_index = self.account_combo.current()
        account_smtp = self.accounts[account_index].SmtpAddress
        delay = float(self.delay_var.get())

        self.stop_event.clear()
        self.send_button.config(state="disabled")
        self.stop_button.config(state="normal")
        self.progress_bar.config(maximum=len(pending), value=0)

        self.send_thread = threading.Thread(
            target=self._send_worker,
            args=(pending, to_col, cc_col, subject_template, runs, line_alignments, attachments, active_images, active_tables, account_smtp, delay, len(already_sent), len(df)),
            daemon=True,
        )
        self.send_thread.start()

    def _show_preview_and_confirm(self, pending, to_col, cc_col, subject_template, runs, attachments, inline_images, inline_tables, total, pending_count) -> bool:
        """Show the rendered draft for the first pending recipient and let the
        user choose to proceed with the full send or cancel."""
        first_row = pending.iloc[0]
        fields = {k: ("" if pd.isna(v) else str(v).strip()) for k, v in first_row.items()}
        preview_to = fields.get(to_col, "")
        preview_cc = format_cc_list(fields.get(cc_col, "")) if cc_col else ""
        preview_subject = render_template(subject_template, fields)
        preview_body = runs_to_plain_text(runs, fields)
        for image in inline_images:
            preview_body = preview_body.replace(image["token"], f'[Image: {os.path.basename(image["path"])}]')
        for table_item in inline_tables:
            rows, cols = len(table_item["rows"]), (len(table_item["rows"][0]) if table_item["rows"] else 0)
            preview_body = preview_body.replace(table_item["token"], f'[Table: {rows} rows x {cols} columns]')

        result = {"proceed": False}

        dialog = tk.Toplevel(self.root)
        dialog.title("Preview draft email — first recipient")
        dialog.geometry("700x600")
        dialog.transient(self.root)
        dialog.grab_set()

        ttk.Label(
            dialog,
            text=f"This is exactly what will be sent to the first recipient "
                 f"({pending_count} email(s) total will be sent this way, each personalised).",
            justify="left", wraplength=660,
        ).pack(anchor="w", padx=12, pady=(12, 6))

        info_frame = ttk.Frame(dialog)
        info_frame.pack(fill="x", padx=12)
        ttk.Label(info_frame, text="To:", width=8).grid(row=0, column=0, sticky="w")
        ttk.Label(info_frame, text=preview_to).grid(row=0, column=1, sticky="w")
        ttk.Label(info_frame, text="CC:", width=8).grid(row=1, column=0, sticky="w")
        ttk.Label(info_frame, text=preview_cc or "(none)").grid(row=1, column=1, sticky="w")
        ttk.Label(info_frame, text="Subject:", width=8).grid(row=2, column=0, sticky="w")
        ttk.Label(info_frame, text=preview_subject, wraplength=580).grid(row=2, column=1, sticky="w")
        attachment_names = ", ".join(os.path.basename(p) for p in attachments) if attachments else "(none)"
        ttk.Label(info_frame, text="Attachments:", width=10).grid(row=3, column=0, sticky="w")
        ttk.Label(info_frame, text=attachment_names, wraplength=580).grid(row=3, column=1, sticky="w")
        inline_names = ", ".join(os.path.basename(img["path"]) for img in inline_images) if inline_images else "(none)"
        ttk.Label(info_frame, text="Inline images:", width=10).grid(row=4, column=0, sticky="w")
        ttk.Label(info_frame, text=inline_names, wraplength=580).grid(row=4, column=1, sticky="w")
        table_count = len(inline_tables)
        ttk.Label(info_frame, text="Tables:", width=10).grid(row=5, column=0, sticky="w")
        ttk.Label(info_frame, text=(f"{table_count} table(s)" if table_count else "(none)")).grid(row=5, column=1, sticky="w")

        ttk.Label(
            dialog,
            text="Message body — [Image: ...] and [Table: ...] show where each item will appear. "
                 "Bold/italic/underline/color formatting is applied in the real email but is not "
                 "shown in this plain preview.",
            wraplength=660, justify="left",
        ).pack(anchor="w", padx=12, pady=(10, 0))
        body_preview = scrolledtext.ScrolledText(dialog, wrap="word", height=18)
        body_preview.pack(fill="both", expand=True, padx=12, pady=(0, 8))
        body_preview.insert("1.0", preview_body)
        body_preview.config(state="disabled")

        button_frame = ttk.Frame(dialog)
        button_frame.pack(fill="x", padx=12, pady=(0, 12))

        def on_send_all():
            result["proceed"] = True
            dialog.destroy()

        def on_cancel():
            result["proceed"] = False
            dialog.destroy()

        ttk.Button(button_frame, text=f"Looks good — Send all {pending_count}", command=on_send_all).pack(side="right", padx=4)
        ttk.Button(button_frame, text="Cancel", command=on_cancel).pack(side="right", padx=4)

        dialog.protocol("WM_DELETE_WINDOW", on_cancel)
        self.root.wait_window(dialog)
        return result["proceed"]

    def _on_stop_clicked(self):
        self.stop_event.set()
        self._log("Stop requested — will halt after the current email.")

    def _send_worker(self, pending, to_col, cc_col, subject_template, runs, line_alignments, attachments, inline_images, inline_tables, account_smtp, delay, already_sent_count, total):
        import time

        # COM objects are bound to the thread that created them. self.outlook
        # was created on the UI thread, so it cannot be used here — this
        # thread needs its own CoInitialize call and its own Outlook/account
        # objects, or every call fails with "marshalled for a different thread".
        pythoncom.CoInitialize()
        try:
            outlook = win32com.client.gencache.EnsureDispatch("Outlook.Application")
            accounts = outlook.Session.Accounts
            account = next(
                (accounts.Item(i) for i in range(1, accounts.Count + 1)
                 if accounts.Item(i).SmtpAddress == account_smtp),
                None,
            )
            if account is None:
                self.msg_queue.put(("done", f"Could not find account '{account_smtp}' on this thread. No emails were sent."))
                return

            log_file = f"sent_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
            entries = []
            sent_count = already_sent_count
            fail_count = 0
            skip_count = already_sent_count
            pending_total = len(pending)

            def write_log():
                with open(log_file, "w", encoding="utf-8") as f:
                    f.write(f"Bulk Email Log — {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n" + "=" * 60 + "\n")
                    for e in entries:
                        f.write(e + "\n")

            for i, (_, row) in enumerate(pending.iterrows(), start=1):
                if self.stop_event.is_set():
                    self.msg_queue.put(("log", "Stopped by user."))
                    break

                recipient = str(row[to_col]).strip()
                fields = {k: ("" if pd.isna(v) else str(v).strip()) for k, v in row.items()}

                try:
                    mail = outlook.CreateItem(0)
                    mail.To = recipient
                    if cc_col and fields.get(cc_col):
                        mail.CC = format_cc_list(fields[cc_col])
                    mail.Subject = render_template(subject_template, fields)

                    body_format, body_content = render_rich_email(runs, fields, inline_images, inline_tables, line_alignments)
                    mail.BodyFormat = body_format
                    if body_format == 2:
                        mail.HTMLBody = body_content
                    else:
                        mail.Body = body_content

                    for attachment_path in attachments:
                        mail.Attachments.Add(attachment_path)
                    for image in inline_images:
                        image_attachment = mail.Attachments.Add(image["path"])
                        image_attachment.PropertyAccessor.SetProperty(PR_ATTACH_CONTENT_ID, image["cid"])

                    mail.SendUsingAccount = account
                    mail.Send()

                    sent_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    status = f"[SENT]   {recipient} — {sent_time}"
                    sent_count += 1
                except Exception as exc:
                    status = f"[FAILED] {recipient} | Error: {exc}"
                    fail_count += 1

                entries.append(status)
                write_log()
                self.msg_queue.put(("progress", (skip_count + i, total, status)))

                if i < pending_total and not self.stop_event.is_set():
                    time.sleep(delay)

            summary = f"Done. Total: {total} | Sent: {sent_count} | Failed: {fail_count} | Log: {log_file}"
            self.msg_queue.put(("done", summary))
        finally:
            pythoncom.CoUninitialize()

    # ── Thread-safe UI updates ──────────────────────────────────────

    def _log(self, message):
        self.log_text.config(state="normal")
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")
        self.log_text.config(state="disabled")

    def _poll_queue(self):
        try:
            while True:
                kind, payload = self.msg_queue.get_nowait()
                if kind == "log":
                    self._log(payload)
                elif kind == "progress":
                    done, total, status = payload
                    self._log(status)
                    self.progress_bar.config(value=done)
                    self.progress_var.set(f"{done}/{total}")
                elif kind == "done":
                    self._log(payload)
                    self.progress_var.set(payload)
                    self.send_button.config(state="normal")
                    self.stop_button.config(state="disabled")
                    self._update_highlight()
                    messagebox.showinfo("Finished", payload)
        except queue.Empty:
            pass
        self.root.after(150, self._poll_queue)


def main():
    root = tk.Tk()
    app = BulkMailerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
