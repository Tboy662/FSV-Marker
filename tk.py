import datetime
import http.client
import json
import os
import re
import smtplib
import socket
import ssl
import sys
import threading
import tkinter as tk
import urllib.request
import uuid
from email.message import EmailMessage
from tkinter import messagebox, ttk

from PIL import Image, ImageDraw, ImageFilter, ImageOps, ImageTk  # pip install pillow


# ---------- Paths (work both as a script and inside a PyInstaller .app) ----------
def resource_path(name):
    """Files bundled with the app (like icon.png)."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, name)


# Settings and logs go in the user's Library folder, since the app bundle is read-only.
DATA_DIR = os.path.join(os.path.expanduser("~/Library/Application Support"), "MarkerTracker")
SAVE_PATH = os.path.join(DATA_DIR, "save.json")
LOG_PATH = os.path.join(DATA_DIR, "markerlog.txt")
ICON_PATH = resource_path("icon.png")

SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
BUTTON_COLORS = ["Blue", "Red", "Green", "Black", "Orange"]
# The dropdown is long enough to fill the screen. Only some of these are real EXPO colors.
OTHER_COLORS = ["Amber", "Amethyst", "Apricot", "Aqua", "Aquamarine", "Beige", "Berry", "Brick Red",
                "Bronze", "Brown", "Burgundy", "Butter", "Canary", "Caramel", "Cherry",
                "Chestnut", "Cobalt", "Copper", "Coral", "Cream", "Crimson", "Cyan",
                "Denim", "Emerald", "Forest", "Fuchsia", "Garnet", "Gold", "Gray",
                "Honey", "Indigo", "Ivory", "Jade", "Khaki", "Lavender", "Lemon",
                "Lilac", "Lime", "Magenta", "Mahogany", "Mango", "Maroon", "Mauve",
                "Mint", "Mocha", "Moss", "Mustard", "Navy Blue", "Olive", "Peach",
                "Periwinkle", "Pine", "Pink", "Plum", "Pumpkin", "Purple", "Raspberry",
                "Rose", "Ruby", "Rust", "Sage", "Salmon", "Sand", "Sapphire", "Scarlet",
                "Seafoam", "Sienna", "Silver", "Sky Blue", "Slate", "Tan", "Tangerine",
                "Taupe", "Teal", "Turquoise", "Violet", "White", "Wine", "Yellow"]
COLORS = BUTTON_COLORS + OTHER_COLORS
ROOMS = [str(r) for r in (*range(200, 221), *range(300, 321))]
SUPPLY_BUTTONS = ["Staples", "Erase Spray"]
OTHER_SUPPLIES = ["Batteries", "Binder Clips", "Calculators", "Chalk", "Disinfecting Wipes",
                  "Erasers", "Folders", "Glue Sticks", "Hand Sanitizer", "Highlighters",
                  "Index Cards", "Notebooks", "Paper", "Paper Clips", "Pencils", "Pens",
                  "Printer Toner", "Rulers", "Scissors", "Sticky Notes", "Tape", "Tissues",
                  "Whiteboard Erasers"]

# ntfy notifications: must match receiver.py
# The topic works like a password: anyone who knows it can read and post.
NTFY_TOPIC = "marker-d2VuZHkncyBiYWNvbmF0b3I"
NTFY_URL = f"https://ntfy.sh/{NTFY_TOPIC}"
NTFY_RETRY = 10  # seconds to wait before trying again
NET_ERRORS = (OSError, http.client.HTTPException, ValueError, KeyError)

# Use certifi's certificates when available, so HTTPS works in the built app on other Macs
try:
    import certifi
    SSL_CONTEXT = ssl.create_default_context(cafile=certifi.where())
except ImportError:
    SSL_CONTEXT = ssl.create_default_context()
RAW_MARKER = "|DATA|"  # everything after this in a message is raw data for the receiver

FONT = ("Helvetica Neue", 13)
FONT_BOLD = ("Helvetica Neue", 13, "bold")
FONT_TITLE = ("Helvetica Neue", 24, "bold")

DEFAULTS = {
    "Subject": "Automated Marker Tracker Email",
    "From": "",
    "To": "",
    "Pass": "",
    # Emails can only be turned on or off by editing this value in save.json (SAVE_PATH),
    # with the app closed. There is no button for it in the app.
    "EmailEnabled": False,
    "Room": ROOMS[0],
    "Body": "The {color} marker lasted {years} years, {months} months, {weeks} weeks, {days} days and {hours} hours.",
    "markers": {c: None for c in COLORS + SUPPLY_BUTTONS + OTHER_SUPPLIES},
}


def item_label(name):
    return f"{name} marker" if name in COLORS else name


# ---------- Persistence ----------
def load_save():
    try:
        with open(SAVE_PATH, "r") as f:
            data = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        data = {}
    merged = {**DEFAULTS, **data}
    merged["markers"] = {**DEFAULTS["markers"], **merged.get("markers", {})}
    return merged


def write_save():
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(SAVE_PATH, "w") as f:
        json.dump(settings, f, indent=4)
    try:
        os.chmod(SAVE_PATH, 0o600)  # file holds the app password
    except OSError:
        pass


settings = load_save()
if not os.path.exists(SAVE_PATH):
    write_save()


# ---------- Helpers ----------
def cfstring(content, context):
    """Replace {name} placeholders; unknown names are left as typed."""
    def replacer(match):
        name = match.group(1).strip()
        return str(context[name]) if name in context else match.group(0)
    return re.sub(r"\{([^}]+)\}", replacer, content)


def send_email(body):
    msg = EmailMessage()
    msg["Subject"] = settings["Subject"]
    msg["From"] = settings["From"]
    msg["To"] = settings["To"]
    msg.set_content(body)
    context = SSL_CONTEXT
    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=15) as server:
        server.ehlo()
        server.starttls(context=context)
        server.ehlo()
        server.login(settings["From"], settings["Pass"])
        server.send_message(msg)


# ---------- ntfy notifications ----------
def ntfy_post(payload):
    """Publish a JSON payload to the topic; returns ntfy's id for the message."""
    data = json.dumps(payload).encode()
    with urllib.request.urlopen(NTFY_URL, data=data, timeout=15, context=SSL_CONTEXT) as resp:
        return json.load(resp)["id"]


def ntfy_read(line):
    """Parse one line of ntfy's stream into (ntfy id, our payload), or (None, None)."""
    try:
        event = json.loads(line)
        if event.get("event") == "message":
            payload = json.loads(event["message"])
            if isinstance(payload, dict):
                return event["id"], payload
    except (ValueError, KeyError, AttributeError):
        pass  # keepalives, or something on the topic that isn't ours
    return None, None


class Broadcaster:
    """Posts a notification to ntfy, then watches the topic until someone sends a stop."""

    def __init__(self):
        self.stop_event = None
        self.msg_id = None
        self.status = ""

    def start(self, msg):
        if self.stop_event is not None:
            self.stop_event.set()  # stop watching the previous notification
        ev = threading.Event()
        self.stop_event = ev
        self.msg_id = uuid.uuid4().hex[:8]
        self.status = "Sending notification…"
        payload = {"type": "msg", "id": self.msg_id, "from": socket.gethostname(), "text": msg}
        threading.Thread(target=self._run, args=(payload, ev), daemon=True).start()

    def send_silent(self, kind, msg, sent_status):
        """Post data for the receivers to save, with no popup for anyone to stop."""
        payload = {"type": kind, "id": uuid.uuid4().hex[:8],
                   "from": socket.gethostname(), "text": msg}
        threading.Thread(target=self._post_silent, args=(payload, sent_status), daemon=True).start()

    def _post_silent(self, payload, sent_status):
        while True:
            try:
                ntfy_post(payload)
                self.status = sent_status
                return
            except NET_ERRORS:
                self.status = "Can't reach ntfy, retrying…"
                threading.Event().wait(NTFY_RETRY)

    def stop(self):
        """Tell the receivers the current notification has been handled."""
        ev = self.stop_event
        if ev is not None and not ev.is_set():
            ev.set()
            self.status = "Broadcast cancelled"
            threading.Thread(target=self._post_stop, args=(self.msg_id, ev), daemon=True).start()

    def _post_stop(self, msg_id, ev):
        try:
            ntfy_post({"type": "stop", "id": msg_id, "from": socket.gethostname()})
        except NET_ERRORS:
            self._set_status(ev, "Couldn't reach ntfy to cancel")

    def _set_status(self, ev, text):
        if self.stop_event is ev:  # ignore threads that were replaced
            self.status = text

    def _run(self, payload, ev):
        since = None
        while since is None and not ev.is_set():
            try:
                since = ntfy_post(payload)
            except NET_ERRORS:
                self._set_status(ev, "Can't reach ntfy, retrying…")
                ev.wait(NTFY_RETRY)
        if ev.is_set():
            return
        self._set_status(ev, "Notification sent")

        while not ev.is_set():
            try:
                url = f"{NTFY_URL}/json?since={since}"
                with urllib.request.urlopen(url, timeout=90, context=SSL_CONTEXT) as stream:
                    for line in stream:
                        if ev.is_set():
                            return
                        event_id, data = ntfy_read(line)
                        if event_id is None:
                            continue
                        since = event_id
                        if data.get("type") == "stop" and data.get("id") == payload["id"]:
                            self._set_status(ev, f"Stopped by {data.get('from', 'someone')}")
                            ev.set()
                            return
            except NET_ERRORS:
                pass
            ev.wait(NTFY_RETRY)


def load_icon(path, scale=0.8, size=1024):
    img = Image.open(path).convert("RGBA")
    bbox = img.getbbox()
    if bbox:
        img = img.crop(bbox)
    inner = int(size * scale)
    img.thumbnail((inner, inner), Image.LANCZOS)
    canvas = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    canvas.paste(img, ((size - img.width) // 2, (size - img.height) // 2), img)
    return ImageTk.PhotoImage(canvas)


# ---------- Glass rendering ----------
def make_background(w, h):
    orange, blue = (214, 90, 0), (20, 90, 255)

    # diagonal gradient: orange at the top left, blue at the bottom right
    v = Image.linear_gradient("L").resize((w, h))
    hz = Image.linear_gradient("L").rotate(90).resize((w, h))
    diag = Image.blend(v, hz, 0.5)
    img = Image.composite(Image.new("RGB", (w, h), blue),
                          Image.new("RGB", (w, h), orange), diag)

    d = ImageDraw.Draw(img)
    d.ellipse((w * 0.5, -h * 0.2, w * 1.2, h * 0.5), fill=(235, 120, 20))
    d.ellipse((-w * 0.2, h * 0.5, w * 0.5, h * 1.2), fill=(0, 60, 220))
    img = img.filter(ImageFilter.GaussianBlur(min(w, h) * 0.15))

    return Image.blend(img, Image.new("RGB", (w, h), (15, 20, 45)), 0.2)


def glass_image(blurred, x, y, w, h, r, tint_alpha=38, highlight=0.35, border=130):
    """Frosted glass: blurred backdrop + white tint + top sheen + bright rim."""
    S = 3
    region = blurred.crop((x, y, x + w, y + h)).convert("RGBA")
    region = Image.alpha_composite(region, Image.new("RGBA", (w, h), (255, 255, 255, tint_alpha)))

    gh = max(1, h // 2)
    grad = ImageOps.invert(Image.linear_gradient("L").resize((w, gh)))
    grad = grad.point(lambda v: int(v * highlight))
    sheen = Image.new("RGBA", (w, h), (255, 255, 255, 0))
    sheen.paste(Image.new("RGBA", (w, gh), (255, 255, 255, 255)), (0, 0), grad)
    region = Image.alpha_composite(region, sheen)

    rim = Image.new("RGBA", (w * S, h * S), (0, 0, 0, 0))
    ImageDraw.Draw(rim).rounded_rectangle(
        (0, 0, w * S - 1, h * S - 1), radius=r * S,
        outline=(255, 255, 255, border), width=int(S * 1.5))
    region = Image.alpha_composite(region, rim.resize((w, h), Image.LANCZOS))

    mask = Image.new("L", (w * S, h * S), 0)
    ImageDraw.Draw(mask).rounded_rectangle((0, 0, w * S - 1, h * S - 1), radius=r * S, fill=255)
    region.putalpha(mask.resize((w, h), Image.LANCZOS))
    return ImageTk.PhotoImage(region)


BUTTON_STYLES = {
    "normal": dict(tint_alpha=38, highlight=0.35, border=120),
    "hover": dict(tint_alpha=70, highlight=0.45, border=170),
    "pressed": dict(tint_alpha=15, highlight=0.10, border=90),
    "selected": dict(tint_alpha=100, highlight=0.50, border=210),
}


class GlassCanvas(tk.Canvas):
    def __init__(self, parent, w, h):
        super().__init__(parent, width=w, height=h, highlightthickness=0, bd=0)
        self.bg_img = make_background(w, h)
        self.blurred = self.bg_img.filter(ImageFilter.GaussianBlur(22))
        self.bg_photo = ImageTk.PhotoImage(self.bg_img)
        self.create_image(0, 0, anchor="nw", image=self.bg_photo)
        self._refs = []

    def panel(self, x, y, w, h, r=26):
        img = glass_image(self.blurred, x, y, w, h, r, tint_alpha=30, highlight=0.25, border=110)
        self._refs.append(img)
        self.create_image(x, y, anchor="nw", image=img)

    def title(self, x, y, text, font=FONT_TITLE, anchor="center"):
        self.create_text(x, y, text=text, fill="white", font=font, anchor=anchor)


class GlassButton:
    def __init__(self, cv, x, y, w, h, text, command, font=FONT_BOLD):
        self.cv, self.command = cv, command
        self.selected = False
        self.inside = False
        self.imgs = {k: glass_image(cv.blurred, x, y, w, h, h // 2, **s)
                     for k, s in BUTTON_STYLES.items()}
        self.item = cv.create_image(x, y, anchor="nw", image=self.imgs["normal"])
        self.label = cv.create_text(x + w / 2, y + h / 2, text=text, fill="white", font=font)
        for it in (self.item, self.label):
            cv.tag_bind(it, "<Enter>", self._enter)
            cv.tag_bind(it, "<Leave>", self._leave)
            cv.tag_bind(it, "<ButtonPress-1>", self._press)
            cv.tag_bind(it, "<ButtonRelease-1>", self._release)

    def _show(self, state):
        self.cv.itemconfigure(self.item, image=self.imgs[state])

    def _base(self):
        return "selected" if self.selected else "normal"

    def _enter(self, _):
        self.inside = True
        self.cv.configure(cursor="hand2")
        self._show("selected" if self.selected else "hover")

    def _leave(self, _):
        self.inside = False
        self.cv.configure(cursor="")
        self._show(self._base())

    def _press(self, _):
        self._show("pressed")

    def _release(self, _):
        self._show("selected" if self.selected else "hover")
        if self.inside and self.command:
            self.command()

    def set_selected(self, flag):
        self.selected = flag
        self._show(self._base())


# ---------- App ----------
class Application:
    def __init__(self, root):
        self.root = root
        self.settings_win = None
        self.selected_color = "Blue"
        self.broadcaster = Broadcaster()
        root.title("Marker Tracker")
        root.resizable(False, False)

        menubar = tk.Menu(root)
        tools = tk.Menu(menubar, tearoff=0)
        tools.add_command(label="Preferences...", command=self.open_settings_window)
        tools.add_command(label="Stop Broadcast", command=self.broadcaster.stop)
        menubar.add_cascade(label="Tools", menu=tools)
        root.config(menu=menubar)
        try:
            root.createcommand("tk::mac::ShowPreferences", self.open_settings_window)
        except tk.TclError:
            pass

        W, H = 360, 335
        cv = GlassCanvas(root, W, H)
        cv.pack()
        self.cv = cv

        cv.title(W / 2, 38, "Marker Tracker")
        cv.panel(20, 70, 320, 245)

        cv.create_text(36, 88, text="Room", fill="white", font=FONT, anchor="w")
        self.room_var = tk.StringVar(value=settings["Room"])
        cv.create_window(82, 88, anchor="w", window=ttk.Entry(root, textvariable=self.room_var), width=60)

        # Color chips: 3 per row, 3 x 90 + 2 x 10 = 290px fills the panel
        CHIPS_PER_ROW = 3
        chip_w, chip_h, gap = 90, 32, 10
        self.chips = {}
        for i, name in enumerate(BUTTON_COLORS):
            row, col = divmod(i, CHIPS_PER_ROW)
            x = 35 + col * (chip_w + gap)
            y = 105 + row * (chip_h + 6)
            self.chips[name] = GlassButton(cv, x, y, chip_w, chip_h, name,
                                           lambda n=name: self.select_color(n), font=FONT)

        # The last slot in the grid is a dropdown for the other colors
        self.other_var = tk.StringVar(value="Other")
        other_box = ttk.Combobox(root, textvariable=self.other_var, values=OTHER_COLORS,
                                 height=len(OTHER_COLORS), state="readonly", font=FONT)
        other_box.bind("<<ComboboxSelected>>", lambda _: self.select_color(self.other_var.get()))
        cv.create_window(35 + 2 * (chip_w + gap), 105 + chip_h + 6, anchor="nw",
                         window=other_box, width=chip_w, height=chip_h)

        # Supplies work like colors, on a third row of the same grid
        supply_y = 105 + 2 * (chip_h + 6)
        for col, item in enumerate(SUPPLY_BUTTONS):
            self.chips[item] = GlassButton(cv, 35 + col * (chip_w + gap), supply_y, chip_w, chip_h, item,
                                           lambda i=item: self.select_color(i), font=FONT)
        self.supply_var = tk.StringVar(value="Supplies")
        supply_box = ttk.Combobox(root, textvariable=self.supply_var, values=OTHER_SUPPLIES,
                                  height=len(OTHER_SUPPLIES), state="readonly", font=FONT)
        supply_box.bind("<<ComboboxSelected>>", lambda _: self.select_color(self.supply_var.get()))
        cv.create_window(35 + 2 * (chip_w + gap), supply_y, anchor="nw",
                         window=supply_box, width=chip_w, height=chip_h)
        self.select_color("Blue")

        GlassButton(cv, 35, 225, 165, 44, "New Marker/Supply", self.add_color, font=FONT)
        GlassButton(cv, 210, 225, 115, 44, "Mark As Dead", self.del_color, font=FONT)

        self.status_text = cv.create_text(W / 2, 295, text="", fill="white",
                                          font=("Helvetica Neue", 11))
        self._poll_broadcast()

        try:
            self.icon = load_icon(ICON_PATH)  # keep the reference
            root.iconphoto(True, self.icon)
        except Exception as e:
            print(f"Icon not loaded: {e}")

    def select_color(self, name):
        self.selected_color = name
        for n, btn in self.chips.items():
            btn.set_selected(n == name)
        self.other_var.set(name if name in OTHER_COLORS else "Other")
        self.supply_var.set(name if name in OTHER_SUPPLIES else "Supplies")

    def read_room(self):
        """The typed room, saved for next time; warns and returns None if it isn't valid."""
        room = self.room_var.get().strip()
        if room not in ROOMS:
            messagebox.showwarning("Error", "Enter a room from 200-220 or 300-320.")
            return None
        settings["Room"] = room
        write_save()
        return room

    def _poll_broadcast(self):
        # The broadcast thread can't touch Tk, so the UI reads its status here
        self.cv.itemconfigure(self.status_text, text=self.broadcaster.status)
        self.root.after(500, self._poll_broadcast)

    # ----- markers -----
    def add_color(self):
        room = self.read_room()
        if room is None:
            return
        color = self.selected_color
        now = datetime.datetime.now()
        settings["markers"][color] = now.isoformat()
        write_save()
        raw = f"start:{room}:{color}:{now.date().toordinal()}"
        self.broadcaster.send_silent("start", f"New {item_label(color)} in room {room}.{RAW_MARKER}{raw}",
                                     "Marker start sent")

    def del_color(self):
        room = self.read_room()
        if room is None:
            return
        color = self.selected_color
        started = settings["markers"].get(color)
        if not started:
            # Started on another computer, e.g. by a teacher who moved rooms. The receiver
            # works out how long it lasted from its own start day for this room, if it has one.
            os.makedirs(DATA_DIR, exist_ok=True)
            with open(LOG_PATH, "a", encoding="utf-8") as f:
                f.write(f"\nColor:{color},Lasted unknown")
            raw = f"stop:{room}:{color}:?"
            self.broadcaster.start(f"The {item_label(color)} in room {room} ran out.{RAW_MARKER}{raw}")
            messagebox.showinfo("INFO", f"The {item_label(color)} in room {room} is marked as out.")
            return  # no email, since there's no time to report

        delta = datetime.datetime.now() - datetime.datetime.fromisoformat(started)
        days = delta.days
        values = {
            "color": color,
            "years": days // 365,
            "months": (days % 365) // 30,
            "weeks": (days % 365) % 30 // 7,
            "days": (days % 365) % 30 % 7,
            "hours": delta.seconds // 3600,
            "total_days": days,
        }
        summary = (f"{values['years']} years, {values['months']} months, "
                   f"{values['weeks']} weeks, {values['days']} days, "
                   f"and {values['hours']} hours")

        os.makedirs(DATA_DIR, exist_ok=True)
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(f"\nColor:{color},Lasted {values['years']}y,{values['months']}m,"
                    f"{values['weeks']}w,{values['days']}d,{values['hours']}h")

        settings["markers"][color] = None
        write_save()
        raw = f"stop:{room}:{color}:{delta.total_seconds() / 86400:.2f}"
        self.broadcaster.start(f"The {item_label(color)} in room {room} ran out after {summary}.{RAW_MARKER}{raw}")
        messagebox.showinfo("INFO", f"{item_label(color)} lasted {summary}.")

        if not settings["EmailEnabled"]:
            return
        if not (settings["From"] and settings["Pass"] and settings["To"]):
            messagebox.showinfo("Email not set up",
                                "Open Preferences (Tools menu) to set up email.")
            return
        try:
            send_email(cfstring(settings["Body"], values))
            messagebox.showinfo("Success", "Email sent successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to send email:\n\n{e}")

    # ----- settings window -----
    def open_settings_window(self):
        if self.settings_win is not None and self.settings_win.winfo_exists():
            self.settings_win.lift()
            return

        win = tk.Toplevel(self.root)
        self.settings_win = win
        win.title("Settings")
        win.resizable(False, False)

        W, H = 400, 640
        cv = GlassCanvas(win, W, H)
        cv.pack()
        cv.title(W / 2, 34, "Settings")
        cv.panel(20, 62, 360, 560)

        self.selfemail_var = tk.StringVar(value=settings["From"])
        self.password_var = tk.StringVar(value=settings["Pass"])
        self.targetemail_var = tk.StringVar(value=settings["To"])

        def field(label, y, var, cmd, btn_text, show=None):
            cv.create_text(40, y, text=label, fill="white", font=FONT, anchor="w")
            e = ttk.Entry(win, textvariable=var, show=show)
            cv.create_window(40, y + 12, anchor="nw", window=e, width=320)
            GlassButton(cv, 40, y + 44, 320, 32, btn_text, cmd, font=FONT)
            return e

        self.selfemail = field("Your email", 86, self.selfemail_var, self.savemail, "Save Email (Self)")
        self.password = field("App password", 176, self.password_var, self.savepass, "Save Password", show="*")
        self.targetemail = field("Target email", 266, self.targetemail_var, self.savetmail, "Save Target Email")

        cv.create_text(40, 356, text="Email content  ({color} {years} {months} {weeks} {days} {hours} {total_days})",
                       fill="white", font=("Helvetica Neue", 11), anchor="w")
        self.text1 = tk.Text(win, wrap="word")
        cv.create_window(40, 368, anchor="nw", window=self.text1, width=320, height=90)
        self.text1.insert("1.0", settings["Body"])
        GlassButton(cv, 40, 470, 320, 32, "Save Email Content", self.on_button4, font=FONT)

        # Emails have no switch in the app on purpose: see "EmailEnabled" in DEFAULTS
        state = "On" if settings["EmailEnabled"] else "Off"
        self.email_note = cv.create_text(
            40, 516, anchor="nw", width=320, fill="white", font=("Helvetica Neue", 11),
            text=f"Emails: {state}. To change this, quit the app, set \"EmailEnabled\" to true or "
                 f"false in the save file, then reopen the app:\nPlease run it by IT for help")

    def on_button4(self):
        # Save the raw template; placeholders are filled in when the email is sent
        settings["Body"] = self.text1.get("1.0", tk.END).strip()
        write_save()

    def savemail(self):
        settings["From"] = self.selfemail.get().strip()
        write_save()

    def savepass(self):
        settings["Pass"] = self.password.get().replace(" ", "")
        write_save()

    def savetmail(self):
        settings["To"] = self.targetemail.get().strip()
        write_save()


if __name__ == "__main__":
    root = tk.Tk()
    app = Application(root)
    root.mainloop()