"""
Jira Key Review – Maninder Edition (Enhanced UX – v2)
-----------------------------------------------------
This build applies the requested UX fixes:
1) No startup flicker/sluggishness (hidden build + debounced resize)
2) Buttons: grey by default; Light SkyBlue on hover/press; black text
3) Username, API Token, and Jira Key entries shortened (half width)
4) Button text black by default
5) Themes: White and Black only
6) Responsive on maximize/minimize; elements re-arrange nicely
7) Overall smoother UX (scroll, resize, layout)

Dependencies:
    pip install jira sv-ttk
Optional:
    pip install pywinstyles
"""

import json
import os
import sys
import threading
import queue
import tempfile
import webbrowser
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkinter.scrolledtext import ScrolledText
import tkinter.font as tkfont

# ---- Optional Theming (Windows 11 Sun Valley) ----
try:
    import sv_ttk  # pip install sv-ttk
except Exception:
    sv_ttk = None
try:
    import pywinstyles  # pip install pywinstyles (optional)
except Exception:
    pywinstyles = None

from jira import JIRA  # pip install jira

# -----------------------
# Configuration
# -----------------------
JIRA_SERVER = "https://hp-jira.external.hp.com"  # <-- SET YOUR JIRA BASE URL
CREDENTIALS_FILE = "credentials.json"

# Custom field IDs (adjust to your Jira if different)
CF_ENCOUNTERED_BY = "customfield_13073"   # Encountered By
CF_FOUND_IN_FW_VER = "customfield_11405"  # Found in FW Version

# Username & Token constraints
MAX_USER_LEN = 254
MAX_TOKEN_LEN = 512

# -----------------------
# UX Colors & Fonts (New)
# -----------------------
BTN_BASE_BG = "#d0d5db"       # Grey default
BTN_HOVER_BG = "#87cefa"      # Light SkyBlue hover
BTN_PRESSED_BG = "#6bbcf6"    # Slightly darker SkyBlue for pressed
BTN_FG = "#000000"            # Black text

# -----------------------
# Helpers
# -----------------------
def load_credentials():
    creds = {"jira_user": "", "jira_api_token": ""}
    try:
        with open(CREDENTIALS_FILE, "r") as f:
            stored = json.load(f)
        creds["jira_user"] = stored.get("jira_user", "")
        creds["jira_api_token"] = stored.get("jira_api_token", "")
    except FileNotFoundError:
        pass
    return creds


def save_credentials(user, token):
    try:
        with open(CREDENTIALS_FILE, "w") as f:
            json.dump(
                {"jira_server": JIRA_SERVER, "jira_user": user, "jira_api_token": token},
                f,
            )
    except Exception:
        pass


def get_client(user, token):
    if not user or not token:
        raise ValueError("Username and API token are required.")
    return JIRA(server=JIRA_SERVER, basic_auth=(user, token))


def iso_to_local(dt_str: str) -> str:
    """Convert Jira ISO date to 'YYYY-MM-DD HH:MM'; fallback to raw if parsing fails."""
    if not dt_str:
        return ""
    fmts = ["%Y-%m-%dT%H:%M:%S.%f%z", "%Y-%m-%dT%H:%M:%S%z"]
    for fmt in fmts:
        try:
            dt = datetime.strptime(dt_str, fmt)
            return dt.strftime("%Y-%m-%d %H:%M")
        except Exception:
            continue
    return dt_str


def as_text(v):
    """Render Jira field values robustly (supports option objects, lists, strings)."""
    if v is None:
        return ""
    if isinstance(v, list):
        return ", ".join(as_text(x) for x in v if x is not None)
    return getattr(v, "value", v)


def pick_last_comment(issue):
    comments_container = getattr(issue.fields, "comment", None)
    if not comments_container or not getattr(comments_container, "comments", None):
        return None
    return max(comments_container.comments, key=lambda c: getattr(c, "created", ""))


def get_description_text(issue):
    desc = getattr(issue.fields, "description", None)
    if desc:
        return str(desc)
    try:
        return issue.raw.get("renderedFields", {}).get("description", "") or ""
    except Exception:
        return ""


# -----------------------
# Tk thread-safe bridge
# -----------------------
ui_queue = queue.Queue()


def post_ui(fn, *args, **kwargs):
    ui_queue.put((fn, args, kwargs))


def process_ui_queue():
    while True:
        try:
            fn, args, kwargs = ui_queue.get_nowait()
            fn(*args, **kwargs)
        except queue.Empty:
            break
    root.after(50, process_ui_queue)


# -----------------------
# Globals
# -----------------------
current_issue = None
selected_files = []  # files chosen to attach when adding a comment
token_is_visible = False

style = None  # ttk.Style will be set after root creation

# Debounce IDs
_resize_debounce_id = None
_attcol_debounce_id = None

# -----------------------
# Workers
# -----------------------
def do_fetch_by_key(key):
    user = username_var.get().strip()
    token = token_var.get().strip()
    if not key:
        post_ui(messagebox.showerror, "Missing Key", "Please enter a Jira Key (e.g., ABC-123).")
        post_ui(set_status, "Idle")
        post_ui(enable_controls, True)
        return
    try:
        jira = get_client(user, token)
        fields = ",".join(
            [
                "attachment",
                "summary",
                "description",
                "comment",
                "issuetype",
                "status",
                "created",
                "updated",
                CF_ENCOUNTERED_BY,
                CF_FOUND_IN_FW_VER,
            ]
        )
        issue = jira.issue(
            key,
            fields=fields,
            expand="renderedFields",
        )
        save_credentials(user, token)
        post_ui(show_issue, issue)
        post_ui(set_status, f"Loaded {issue.key}")
    except Exception as e:
        post_ui(messagebox.showerror, "Error", str(e))
        post_ui(set_status, "Idle")
    finally:
        post_ui(enable_controls, True)


def do_add_comment_and_attachments(body_text: str, files: list):
    global current_issue
    if not current_issue:
        post_ui(messagebox.showerror, "No Issue", "Load an issue first.")
        post_ui(enable_controls, True)
        post_ui(set_status, "Idle")
        return
    user = username_var.get().strip()
    token = token_var.get().strip()
    try:
        jira = get_client(user, token)
        # 1) Add comment (supports pasted URLs / Jira markup)
        if body_text:
            jira.add_comment(current_issue.key, body_text)
        # 2) Upload attachments (if any)
        for fp in files or []:
            if os.path.isfile(fp):
                jira.add_attachment(issue=current_issue.key, attachment=fp)
        # 3) Reload to refresh last comment and attachments
        fields = ",".join(
            [
                "attachment",
                "summary",
                "description",
                "comment",
                "issuetype",
                "status",
                "created",
                "updated",
                CF_ENCOUNTERED_BY,
                CF_FOUND_IN_FW_VER,
            ]
        )
        refreshed = jira.issue(
            current_issue.key,
            fields=fields,
            expand="renderedFields",
        )
        post_ui(show_issue, refreshed)
        post_ui(clear_selected_files)
        post_ui(messagebox.showinfo, "Success", "Comment and attachments have been added successfully.")
        post_ui(set_status, f"Comment/attachments added to {current_issue.key}")
    except Exception as e:
        post_ui(messagebox.showerror, "Error", str(e))
        post_ui(set_status, "Idle")
    finally:
        post_ui(enable_controls, True)


def do_download_attachment(attachment, save_path, open_after=False):
    """Download an attachment via the authenticated Jira session."""
    user = username_var.get().strip()
    token = token_var.get().strip()
    try:
        jira = get_client(user, token)
        resp = jira._session.get(attachment.content, stream=True)  # authenticated content URL
        resp.raise_for_status()
        with open(save_path, "wb") as f:
            for chunk in resp.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)
        if open_after:
            webbrowser.open(f"file://{os.path.abspath(save_path)}")
        post_ui(set_status, f"Saved: {os.path.basename(save_path)}")
    except Exception as e:
        post_ui(
            messagebox.showerror,
            "Download Error",
            f"{getattr(attachment, 'filename', 'attachment')}:\n\n{e}",
        )
        post_ui(set_status, "Idle")
    finally:
        post_ui(enable_controls, True)


# -----------------------
# UI utilities
# -----------------------
def set_status(text):
    status_var.set(text)


def enable_controls(enable: bool):
    state = "normal" if enable else "disabled"
    for w in (
        fetch_btn, reset_btn, key_entry, open_btn, copy_desc_btn, copy_comment_btn,
        add_comment_btn, comment_input, add_files_btn, clear_files_btn,
        refresh_att_btn, open_att_btn, download_att_btn, theme_select, show_hide_btn
    ):
        try:
            w.configure(state=("readonly" if w is theme_select and enable else state))
        except Exception:
            pass


def on_fetch():
    enable_controls(False)
    set_status("Loading...")
    key = key_entry_var.get().strip()
    threading.Thread(target=do_fetch_by_key, args=(key,), daemon=True).start()


def on_reset():
    global current_issue
    current_issue = None
    key_entry_var.set("")
    # Clear meta
    for var in (
        key_value_var,
        summary_value_var,
        status_value_var,
        created_value_var,
        updated_value_var,
        encountered_by_var,
        fw_version_var,
    ):
        var.set("")
    # Clear text areas
    for w in (desc_text, comment_text):
        w.configure(state="normal")
        w.delete("1.0", tk.END)
        w.configure(state="disabled")
    comment_input.configure(state="normal")
    comment_input.delete("1.0", tk.END)
    # Clear attachments and selected files
    clear_attachments()
    clear_selected_files()
    set_status("Reset")


def copy_description():
    text = desc_text.get("1.0", tk.END).strip()
    root.clipboard_clear()
    root.clipboard_append(text)
    messagebox.showinfo("Copied", "Description copied to clipboard.")


def copy_last_comment():
    text = comment_text.get("1.0", tk.END).strip()
    root.clipboard_clear()
    root.clipboard_append(text)
    messagebox.showinfo("Copied", "Last comment copied to clipboard.")


def on_add_comment():
    body = comment_input.get("1.0", tk.END).strip()
    if not body and not selected_files:
        messagebox.showwarning("Nothing to add", "Type a comment or choose files to attach.")
        return
    enable_controls(False)
    set_status("Submitting comment and attachments...")
    threading.Thread(
        target=do_add_comment_and_attachments,
        args=(body, selected_files.copy()),
        daemon=True,
    ).start()


def add_files():
    files = filedialog.askopenfilenames(title="Select files to attach")
    if not files:
        return
    for f in files:
        if f not in selected_files:
            selected_files.append(f)
            files_list.insert(tk.END, os.path.basename(f))


def clear_selected_files():
    selected_files.clear()
    files_list.delete(0, tk.END)


def _safe_attr(obj, name, default=None):
    if hasattr(obj, name):
        return getattr(obj, name)
    if isinstance(obj, dict):
        return obj.get(name, default)
    return default


def populate_attachments(issue):
    clear_attachments()
    attachments = getattr(issue.fields, "attachment", []) or []
    for idx, a in enumerate(attachments):
        filename = _safe_attr(a, "filename", "")
        size_kb = round((_safe_attr(a, "size", 0) or 0) / 1024, 1)
        who = _safe_attr(_safe_attr(a, "author", {}) or {}, "displayName", "Unknown")
        when = iso_to_local(_safe_attr(a, "created", ""))
        att_tree.insert("", "end", iid=str(idx), values=(filename, size_kb, who, when))
    att_tree.attachments = attachments
    set_status(f"{len(attachments)} attachment(s)")


def clear_attachments():
    for row in att_tree.get_children():
        att_tree.delete(row)
    att_tree.attachments = []


def on_refresh_attachments():
    if not current_issue:
        messagebox.showinfo("No Issue", "Load an issue first.")
        return
    enable_controls(False)
    set_status("Refreshing attachments...")
    threading.Thread(target=do_fetch_by_key, args=(current_issue.key,), daemon=True).start()


def get_selected_attachment():
    sel = att_tree.selection()
    if not sel:
        messagebox.showinfo("Select Attachment", "Please select an attachment first.")
        return None
    idx = int(sel[0])
    try:
        return att_tree.attachments[idx]
    except Exception:
        messagebox.showerror("Error", "Could not resolve selected attachment.")
        return None


def on_open_attachment():
    a = get_selected_attachment()
    if not a:
        return
    enable_controls(False)
    set_status(f"Opening {getattr(a, 'filename', 'attachment')}...")
    fd, tmp_path = tempfile.mkstemp(prefix="jira_att_", suffix=f"_{getattr(a, 'filename', 'file')}")
    os.close(fd)
    threading.Thread(target=do_download_attachment, args=(a, tmp_path, True), daemon=True).start()


def on_download_attachment():
    a = get_selected_attachment()
    if not a:
        return
    out_dir = filedialog.askdirectory(title="Select folder to save attachment")
    if not out_dir:
        return
    save_path = os.path.join(out_dir, getattr(a, "filename", "file"))
    enable_controls(False)
    set_status(f"Downloading {getattr(a, 'filename', 'attachment')}...")
    threading.Thread(target=do_download_attachment, args=(a, save_path, False), daemon=True).start()


def show_issue(issue):
    """Populate UI with a full issue (details, comment, extra fields, attachments)."""
    global current_issue
    current_issue = issue
    key_value_var.set(issue.key)
    summary_value_var.set(getattr(issue.fields, "summary", "") or "")
    status_value_var.set(getattr(getattr(issue.fields, "status", None), "name", "") if issue.fields.status else "")
    created_value_var.set(iso_to_local(getattr(issue.fields, "created", "")))
    updated_value_var.set(iso_to_local(getattr(issue.fields, "updated", "")))

    # Extra fields
    encountered_by = as_text(getattr(issue.fields, CF_ENCOUNTERED_BY, None))
    fw_version = as_text(getattr(issue.fields, CF_FOUND_IN_FW_VER, None))
    encountered_by_var.set(encountered_by)
    fw_version_var.set(fw_version)

    # Description
    desc = get_description_text(issue)
    desc_text.configure(state="normal")
    desc_text.delete("1.0", tk.END)
    desc_text.insert(tk.END, desc if desc else "(No description)")
    desc_text.configure(state="disabled")

    # Last comment
    lc = pick_last_comment(issue)
    comment_text.configure(state="normal")
    comment_text.delete("1.0", tk.END)
    if lc:
        who = getattr(getattr(lc, "author", None), "displayName", "Unknown")
        when = iso_to_local(getattr(lc, "created", ""))
        body = getattr(lc, "body", "") or ""
        comment_text.insert(tk.END, f"By: {who}\nOn: {when}\n\n{body}")
    else:
        comment_text.insert(tk.END, "(No comments)")
    comment_text.configure(state="disabled")

    # Attachments
    populate_attachments(issue)

    # Browser link
    open_btn.configure(command=lambda: webbrowser.open(f"{JIRA_SERVER}/browse/{issue.key}"))


# -----------------------
# Theming (White/Black only)
# -----------------------
def set_titlebar(color_hex: str):
    if pywinstyles:
        try:
            ver = sys.getwindowsversion()
            if ver.major == 10 and ver.build >= 22000:  # Windows 11
                pywinstyles.change_header_color(root, color_hex)
        except Exception:
            pass


def apply_theme(choice: str):
    """Two themes only: White (light) / Black (dark)"""
    if sv_ttk is None:
        return  # use default ttk
    key = (choice or "White").lower()
    if key == "white":
        sv_ttk.set_theme("light")
        root.after(50, lambda: set_titlebar("#6d6868f7"))
    else:  # black
        sv_ttk.set_theme("dark")
        root.after(50, lambda: set_titlebar("#1c1c1c"))


def on_theme_change(event=None):
    apply_theme(theme_var.get())


# -----------------------
# Validation (username/token)
# -----------------------
def limit_len_factory(limit, var, counter_label):
    def _validate(new_text):
        if len(new_text) <= limit:
            counter_label.config(text=f"{len(new_text)}/{limit}")
            return True
        root.bell()
        return False

    vcmd = (root.register(_validate), "%P")
    return vcmd


def toggle_token_visibility():
    global token_is_visible
    token_is_visible = not token_is_visible
    token_entry.config(show="" if token_is_visible else "•")
    show_hide_btn.config(text="Hide" if token_is_visible else "Show")


# -----------------------
# Scroll handling (Enhanced)
# -----------------------
def _on_frame_configure(event):
    main_canvas.configure(scrollregion=main_canvas.bbox("all"))


# -----------------------
# Build GUI (hidden to avoid flicker)
# -----------------------
root = tk.Tk()
root.withdraw()  # hide while constructing to avoid flicker
root.title("Jira Issue Review/Update – By Maninder Singh")

# Sane min/max and adaptive default
screen_w, screen_h = root.winfo_screenwidth(), root.winfo_screenheight()
root.minsize(980, 720)
try:
    if sys.platform.startswith("win"):
        root.state("zoomed")
    else:
        w, h = int(screen_w * 0.82), int(screen_h * 0.82)
        root.geometry(f"{w}x{h}+{(screen_w-w)//2}+{(screen_h-h)//2}")
except Exception:
    root.geometry("1220x900")

# ttk Style & Sun Valley theme (base)
style = ttk.Style(root)
if sv_ttk:
    try:
        sv_ttk.set_theme("light")  # initial base
    except Exception:
        pass

# Fonts
try:
    default_font = tkfont.nametofont("TkDefaultFont"); default_font.configure(family="Segoe UI", size=10)
    heading_font = tkfont.nametofont("TkHeadingFont"); heading_font.configure(family="Segoe UI", size=11, weight="bold")
    TITLE_FONT = tkfont.Font(family="Segoe UI", size=14, weight="bold")
    SECTION_FONT = tkfont.Font(family="Segoe UI Semibold", size=12)
except Exception:
    TITLE_FONT = tkfont.nametofont("TkHeadingFont"); TITLE_FONT.configure(size=14, weight="bold")
    SECTION_FONT = tkfont.nametofont("TkDefaultFont"); SECTION_FONT.configure(size=12, weight="bold")

# Button style: Grey default, SkyBlue hover/pressed, black text
def init_button_style():
    style.configure(
        "Modern.TButton",
        padding=(12, 8),
        relief="flat",
        foreground=BTN_FG,
        background=BTN_BASE_BG,
        borderwidth=1,
        focusthickness=1,
        focuscolor=BTN_PRESSED_BG,
    )
    style.map(
        "Modern.TButton",
        background=[("pressed", BTN_PRESSED_BG), ("active", BTN_HOVER_BG), ("!active", BTN_BASE_BG)],
        foreground=[("disabled", "#666666"), ("!disabled", BTN_FG)],
        relief=[("pressed", "flat"), ("!pressed", "flat")],
    )

# Section title visuals
def init_section_style():
    style.configure("Section.TLabelframe.Label", font=SECTION_FONT, foreground="#4c4f57")
    style.configure("Section.TLabelframe", padding=(6, 6, 6, 6))

init_button_style()
init_section_style()

# --- Scrollable container (canvas + frame) ---
outer = ttk.Frame(root)
outer.pack(fill="both", expand=True)

main_canvas = tk.Canvas(outer, highlightthickness=0)
vbar = ttk.Scrollbar(outer, orient="vertical", command=main_canvas.yview)
main_canvas.configure(yscrollcommand=vbar.set)
vbar.pack(side="right", fill="y")
main_canvas.pack(side="left", fill="both", expand=True)

scroll_container = ttk.Frame(main_canvas)
container_window = main_canvas.create_window((0, 0), window=scroll_container, anchor="nw")
scroll_container.bind("<Configure>", _on_frame_configure)

# keep inner frame width synced to canvas width
def _sync_canvas_width(event):
    try:
        main_canvas.itemconfig(container_window, width=event.width)
    except tk.TclError:
        pass
main_canvas.bind("<Configure>", _sync_canvas_width)

# Scrolling (only when cursor is over the canvas)
def _bind_wheel_to_canvas(_=None):
    try:
        def _mw(e):
            step = -int(e.delta / 120)  # responsive & not sluggish
            main_canvas.yview_scroll(step, "units")
        root.bind_all("<MouseWheel>", _mw)                             # Windows/Mac
        root.bind_all("<Button-4>", lambda e: main_canvas.yview_scroll(-3, "units"))  # Linux up
        root.bind_all("<Button-5>", lambda e: main_canvas.yview_scroll( 3, "units"))  # Linux down
    except Exception:
        pass

def _unbind_wheel_from_canvas(_=None):
    try:
        root.unbind_all("<MouseWheel>")
        root.unbind_all("<Button-4>")
        root.unbind_all("<Button-5>")
    except Exception:
        pass

main_canvas.bind("<Enter>", _bind_wheel_to_canvas)
main_canvas.bind("<Leave>", _unbind_wheel_from_canvas)

# ---- Title row + Theme selector (White/Black only) ----
title_row = ttk.Frame(scroll_container)
title_row.grid(row=0, column=0, padx=10, pady=(10, 5), sticky="ew")
title_lbl = ttk.Label(title_row, text="Jira Issue Review/Update", font=TITLE_FONT)
title_lbl.pack(side="left")

ttk.Label(title_row, text="Theme:").pack(side="right", padx=(8, 4))
theme_var = tk.StringVar(value="White")
theme_select = ttk.Combobox(title_row, textvariable=theme_var, values=["White", "Black"], state="readonly", width=8)
theme_select.pack(side="right")
theme_select.bind("<<ComboboxSelected>>", on_theme_change)

# ---- Attribution ----
about_row = ttk.Frame(scroll_container)
about_row.grid(row=1, column=0, padx=10, pady=(0, 10), sticky="ew")
ttk.Label(about_row, text="By Maninder Singh", foreground="#0a6efd").pack(side="left")

# ---- Credentials (shorter entries) ----
creds_frame = ttk.LabelFrame(scroll_container, text="Credentials (Jira URL is hard-coded)", style="Section.TLabelframe")
creds_frame.grid(row=2, column=0, padx=10, pady=10, sticky="ew")

# No expanding entry columns (keep them visually shorter)
creds_frame.columnconfigure(0, weight=0)
creds_frame.columnconfigure(1, weight=0)
creds_frame.columnconfigure(2, weight=0)
creds_frame.columnconfigure(3, weight=0)

ttk.Label(creds_frame, text="Jira URL:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
ttk.Label(creds_frame, text=JIRA_SERVER).grid(row=0, column=1, sticky="w", padx=5, pady=5, columnspan=3)

# Username
ttk.Label(creds_frame, text="Username (email):").grid(row=1, column=0, sticky="w", padx=5, pady=(8, 2))
username_var = tk.StringVar()
username_entry = ttk.Entry(creds_frame, textvariable=username_var, width=32)  # shortened
username_entry.grid(row=1, column=1, sticky="w", padx=5, pady=(8, 2))
user_count = ttk.Label(creds_frame, text=f"0/{MAX_USER_LEN}")
user_count.grid(row=1, column=2, sticky="w", padx=5)

# Token
ttk.Label(creds_frame, text="API Token / Password:").grid(row=2, column=0, sticky="w", padx=5, pady=(2, 8))
token_var = tk.StringVar()
token_entry = ttk.Entry(creds_frame, textvariable=token_var, show="•", width=32)  # shortened
token_entry.grid(row=2, column=1, sticky="w", padx=5, pady=(2, 8))
token_count = ttk.Label(creds_frame, text=f"0/{MAX_TOKEN_LEN}")
token_count.grid(row=2, column=2, sticky="w", padx=5)
show_hide_btn = ttk.Button(creds_frame, text="Show", width=8, command=toggle_token_visibility, style="Modern.TButton")
show_hide_btn.grid(row=2, column=3, sticky="w", padx=5)

# Length validation
username_entry.configure(validate="key", validatecommand=limit_len_factory(MAX_USER_LEN, username_var, user_count))
token_entry.configure(validate="key", validatecommand=limit_len_factory(MAX_TOKEN_LEN, token_var, token_count))

# Load saved creds
_creds = load_credentials()
if _creds.get("jira_user"):
    username_var.set(_creds["jira_user"]); user_count.config(text=f"{len(_creds['jira_user'])}/{MAX_USER_LEN}")
if _creds.get("jira_api_token"):
    token_var.set(_creds["jira_api_token"]); token_count.config(text=f"{len(_creds['jira_api_token'])}/{MAX_TOKEN_LEN}")

# ---- Fetch controls (shorter key entry) ----
controls = ttk.LabelFrame(scroll_container, text="Fetch by Key", style="Section.TLabelframe")
controls.grid(row=3, column=0, padx=10, pady=10, sticky="ew")
controls.columnconfigure(0, weight=0)
controls.columnconfigure(1, weight=0)
controls.columnconfigure(2, weight=0)
controls.columnconfigure(3, weight=0)

ttk.Label(controls, text="Jira Key:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
key_entry_var = tk.StringVar()
key_entry = ttk.Entry(controls, textvariable=key_entry_var, width=24)  # shortened
key_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")

fetch_btn = ttk.Button(controls, text="Fetch", command=on_fetch, style="Modern.TButton")
fetch_btn.grid(row=0, column=2, padx=(5, 8), pady=5, sticky="w")
reset_btn = ttk.Button(controls, text="Reset", command=on_reset, style="Modern.TButton")
reset_btn.grid(row=0, column=3, padx=5, pady=5, sticky="w")

# ---- Issue details ----
details_frame = ttk.LabelFrame(scroll_container, text="Issue Details", style="Section.TLabelframe")
details_frame.grid(row=4, column=0, padx=10, pady=10, sticky="nsew")
details_frame.columnconfigure(0, weight=1)

meta_grid = ttk.Frame(details_frame)
meta_grid.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
for i in range(4):
    meta_grid.columnconfigure(i, weight=1)

key_value_var = tk.StringVar()
summary_value_var = tk.StringVar()
status_value_var = tk.StringVar()
created_value_var = tk.StringVar()
updated_value_var = tk.StringVar()
encountered_by_var = tk.StringVar()
fw_version_var = tk.StringVar()

ttk.Label(meta_grid, text="Key:").grid(row=0, column=0, sticky="w"); ttk.Label(meta_grid, textvariable=key_value_var).grid(row=0, column=1, sticky="w")
ttk.Label(meta_grid, text="Status:").grid(row=0, column=2, sticky="w"); ttk.Label(meta_grid, textvariable=status_value_var).grid(row=0, column=3, sticky="w")
ttk.Label(meta_grid, text="Summary:").grid(row=1, column=0, sticky="w"); ttk.Label(meta_grid, textvariable=summary_value_var).grid(row=1, column=1, columnspan=3, sticky="w")
ttk.Label(meta_grid, text="Created:").grid(row=2, column=0, sticky="w"); ttk.Label(meta_grid, textvariable=created_value_var).grid(row=2, column=1, sticky="w")
ttk.Label(meta_grid, text="Updated:").grid(row=2, column=2, sticky="w"); ttk.Label(meta_grid, textvariable=updated_value_var).grid(row=2, column=3, sticky="w")
ttk.Label(meta_grid, text="Encountered By:").grid(row=3, column=0, sticky="w"); ttk.Label(meta_grid, textvariable=encountered_by_var).grid(row=3, column=1, sticky="w")
ttk.Label(meta_grid, text="Found in FW Version:").grid(row=3, column=2, sticky="w"); ttk.Label(meta_grid, textvariable=fw_version_var).grid(row=3, column=3, sticky="w")

# Description & Last Comment (expandable + scrollable)
text_frame = ttk.Frame(details_frame)
text_frame.grid(row=1, column=0, padx=5, pady=5, sticky="nsew")
details_frame.rowconfigure(1, weight=1)
details_frame.columnconfigure(0, weight=1)

desc_box = ttk.LabelFrame(text_frame, text="Description", style="Section.TLabelframe")
desc_box.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
comment_box = ttk.LabelFrame(text_frame, text="Last Comment", style="Section.TLabelframe")
comment_box.grid(row=0, column=1, padx=5, pady=5, sticky="nsew")

text_frame.columnconfigure(0, weight=1)
text_frame.columnconfigure(1, weight=1)
text_frame.rowconfigure(0, weight=1)

desc_text = ScrolledText(desc_box, wrap="word", height=16)
desc_text.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
desc_text.configure(state="disabled")

comment_text = ScrolledText(comment_box, wrap="word", height=16)
comment_text.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
comment_text.configure(state="disabled")

# Action buttons under the text areas
buttons_frame = ttk.Frame(details_frame)
buttons_frame.grid(row=2, column=0, padx=5, pady=5, sticky="w")

open_btn = ttk.Button(buttons_frame, text="Open in Browser", command=lambda: None, style="Modern.TButton")
open_btn.grid(row=0, column=0, padx=5)
copy_desc_btn = ttk.Button(buttons_frame, text="Copy Description", command=copy_description, style="Modern.TButton")
copy_desc_btn.grid(row=0, column=1, padx=5)
copy_comment_btn = ttk.Button(buttons_frame, text="Copy Last Comment", command=copy_last_comment, style="Modern.TButton")
copy_comment_btn.grid(row=0, column=2, padx=5)

# ---- Add Comment (supports links + file attachments) ----
add_comment_frame = ttk.LabelFrame(scroll_container, text="Add Comment & Attach Files", style="Section.TLabelframe")
add_comment_frame.grid(row=5, column=0, padx=10, pady=10, sticky="ew")
add_comment_frame.columnconfigure(0, weight=1)

comment_input = ScrolledText(add_comment_frame, height=5, wrap="word")
comment_input.grid(row=0, column=0, padx=5, pady=5, sticky="ew", columnspan=3)

files_list = tk.Listbox(add_comment_frame, height=4)
files_list.grid(row=1, column=0, padx=5, pady=5, sticky="ew")

add_files_btn = ttk.Button(add_comment_frame, text="Attach files…", command=add_files, style="Modern.TButton")
add_files_btn.grid(row=1, column=1, padx=5, pady=5, sticky="w")
clear_files_btn = ttk.Button(add_comment_frame, text="Clear files", command=clear_selected_files, style="Modern.TButton")
clear_files_btn.grid(row=1, column=2, padx=5, pady=5, sticky="w")
add_comment_btn = ttk.Button(add_comment_frame, text="Add Comment", command=on_add_comment, style="Modern.TButton")
add_comment_btn.grid(row=2, column=0, padx=5, pady=5, sticky="w")

# ---- Attachments section ----
attachments_frame = ttk.LabelFrame(scroll_container, text="Attachments", style="Section.TLabelframe")
attachments_frame.grid(row=6, column=0, padx=10, pady=10, sticky="nsew")

att_cols = ("Name", "Size (KB)", "Author", "Created")
att_tree = ttk.Treeview(attachments_frame, columns=att_cols, show="headings", height=8)
for c in att_cols:
    att_tree.heading(c, text=c)
    if c == "Name":
        att_tree.column(c, width=520, anchor="w")
    elif c == "Size (KB)":
        att_tree.column(c, width=90, anchor="e")
    else:
        att_tree.column(c, width=180, anchor="w")
att_tree.grid(row=0, column=0, columnspan=4, padx=5, pady=5, sticky="nsew")

# Scrollbar for attachments
att_scroll = ttk.Scrollbar(attachments_frame, orient="vertical", command=att_tree.yview)
att_tree.configure(yscroll=att_scroll.set)
att_scroll.grid(row=0, column=4, sticky="ns", padx=(0, 5), pady=5)

attachments_frame.columnconfigure(0, weight=1)
attachments_frame.rowconfigure(0, weight=1)

def _autosize_attachment_columns():
    total = max(att_tree.winfo_width() - 20, 640)
    widths = {
        "Name": int(total * 0.50),
        "Size (KB)": int(total * 0.10),
        "Author": int(total * 0.20),
        "Created": int(total * 0.20),
    }
    for col, w in widths.items():
        att_tree.column(col, width=w)

def _debounced_attcols(event=None):
    global _attcol_debounce_id
    if _attcol_debounce_id:
        root.after_cancel(_attcol_debounce_id)
    _attcol_debounce_id = root.after(60, _autosize_attachment_columns)

att_tree.bind("<Configure>", _debounced_attcols)

refresh_att_btn = ttk.Button(attachments_frame, text="Refresh", command=on_refresh_attachments, style="Modern.TButton")
open_att_btn = ttk.Button(attachments_frame, text="Open", command=on_open_attachment, style="Modern.TButton")
download_att_btn = ttk.Button(attachments_frame, text="Download…", command=on_download_attachment, style="Modern.TButton")
refresh_att_btn.grid(row=1, column=1, padx=5, pady=5, sticky="e")
open_att_btn.grid(row=1, column=2, padx=5, pady=5, sticky="e")
download_att_btn.grid(row=1, column=3, padx=5, pady=5, sticky="e")

# ---- Status bar ----
status_var = tk.StringVar(value="Idle")
status_bar = ttk.Label(scroll_container, textvariable=status_var, relief="sunken", anchor="w")
status_bar.grid(row=7, column=0, sticky="ew", padx=10, pady=(0, 10))

# ---- Resize text areas with debounce ----
_row_height_px = tkfont.Font(font=desc_text["font"]).metrics("linespace")
def _resize_text_areas():
    try:
        h = details_frame.winfo_height()
        rows = max(8, min(22, (h - 220) // max(16, _row_height_px)))
        for w in (desc_text, comment_text):
            w.configure(height=int(rows))
    except Exception:
        pass

def _debounced_resize(event=None):
    global _resize_debounce_id
    if _resize_debounce_id:
        root.after_cancel(_resize_debounce_id)
    _resize_debounce_id = root.after(60, _resize_text_areas)

root.bind("<Configure>", _debounced_resize)

# Start UI queue pump
process_ui_queue()

# Apply initial theme AFTER we show the window (reduces flicker)
root.update_idletasks()
root.deiconify()
apply_theme("White")

# Enter main loop
root.mainloop()
