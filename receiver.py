import datetime
import http.client
import json
import queue
import socket
import ssl
import threading
import time
import tkinter as tk
import urllib.request
from tkinter import filedialog, messagebox, ttk
import os
# Must match tk.py
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
DATA_DIR = os.path.join(os.path.expanduser("~/Library/Application Support"), "MarkerTracker")
SAVE_PATH = os.path.join(DATA_DIR, "rsave.json")
REQUESTS_PATH = os.path.join(DATA_DIR, "requests.xlsx")  # every request, one row each
FONT = ("Helvetica Neue", 13)
FONT_BOLD = ("Helvetica Neue", 13, "bold")
COLORS = ["blue", "red", "green", "black", "orange", "amber", "amethyst", "apricot", "aqua",
          "aquamarine", "beige", "berry", "brick red", "bronze", "brown", "burgundy",
          "butter", "canary", "caramel", "cherry", "chestnut", "cobalt", "copper",
          "coral", "cream", "crimson", "cyan", "denim", "emerald", "forest", "fuchsia",
          "garnet", "gold", "gray", "honey", "indigo", "ivory", "jade", "khaki",
          "lavender", "lemon", "lilac", "lime", "magenta", "mahogany", "mango",
          "maroon", "mauve", "mint", "mocha", "moss", "mustard", "navy blue", "olive",
          "peach", "periwinkle", "pine", "pink", "plum", "pumpkin", "purple",
          "raspberry", "rose", "ruby", "rust", "sage", "salmon", "sand", "sapphire",
          "scarlet", "seafoam", "sienna", "silver", "sky blue", "slate", "tan",
          "tangerine", "taupe", "teal", "turquoise", "violet", "white", "wine", "yellow"]  # must match tk.py
ROOMS = [str(n) for n in (*range(200, 221), *range(300, 321))]
RAW_MARKER = "|DATA|"  # must match tk.py
SUPPLIES = ["Staples", "Erase Spray", "Batteries", "Binder Clips", "Calculators",
            "Chalk", "Disinfecting Wipes", "Erasers", "Folders", "Glue Sticks",
            "Hand Sanitizer", "Highlighters", "Index Cards", "Notebooks", "Paper",
            "Paper Clips", "Pencils", "Pens", "Printer Toner", "Rulers", "Scissors",
            "Sticky Notes", "Tape", "Tissues", "Whiteboard Erasers"]  # must match tk.py
TRACKED = COLORS + [s.lower() for s in SUPPLIES]
COMMON_COLORS = COLORS[:5]  # the buttons in tk.py; always kept in the save
MAX_PROCESSED = 1000  # how many message ids to remember


def add_color(room_data, color):
    room_data.setdefault(color + "start", 0)  # day number, from date.toordinal()
    room_data.setdefault(color + "time", [0, 0])  # [total days lasted, markers counted]


def new_room():
    data = {}
    for color in COMMON_COLORS:  # other colors get added the first time they're used
        add_color(data, color)
    return data


savedata = {
    "rooms": {room: new_room() for room in ROOMS},
    "orders": [],  # {"room", "item", "done"}, oldest first
    "processed": [],  # ids already saved, since ntfy replays messages after a restart
}


def upd_time(room, color, days):
    add_color(savedata["rooms"][room], color)
    entry = savedata["rooms"][room][color + "time"]
    entry[0] += days
    entry[1] += 1


def start_mark(room, color, day):
    add_color(savedata["rooms"][room], color)
    savedata["rooms"][room][color + "start"] = day


def stop_mark(room, color, days):
    upd_time(room, color, days)
    savedata["rooms"][room][color + "start"] = 0


def export_report(path):
    """Write used marker colors per room, plus averages, to an Excel file."""
    from openpyxl import Workbook  # pip install openpyxl
    from openpyxl.styles import Font

    rows, by_color, by_room = [], {}, {}
    for room, values in savedata["rooms"].items():
        for color in COLORS:
            start = values.get(color + "start", 0)
            days, count = values.get(color + "time", [0, 0])
            if not (start or count):
                continue  # never used in this room
            running = datetime.date.fromordinal(start) if start else None
            rows.append([int(room), color.title(), count, round(days, 2),
                         round(days / count, 2) if count else None, running])
            for totals, key in ((by_color, color.title()), (by_room, int(room))):
                entry = totals.setdefault(key, [0, 0])
                entry[0] += count
                entry[1] += days

    def add_sheet(sheet, headers, data):
        sheet.append(headers)
        for cell in sheet[1]:
            cell.font = Font(bold=True)
        for row in data:
            sheet.append(row)
        for i, header in enumerate(headers):
            longest = max([len(str(header))] + [len(str(row[i])) for row in data if row[i] is not None])
            sheet.column_dimensions[chr(65 + i)].width = longest + 2
        sheet.freeze_panes = "A2"

    def averages(totals):
        return [[key, count, round(days, 2), round(days / count, 2) if count else None]
                for key, (count, days) in totals.items()]

    book = Workbook()
    add_sheet(book.active, ["Room", "Color", "Markers Used", "Total Days", "Average Days", "Running Since"], rows)
    book.active.title = "By Room"
    add_sheet(book.create_sheet("Average by Color"),
              ["Color", "Markers Used", "Total Days", "Average Days"],
              sorted(averages(by_color)))
    add_sheet(book.create_sheet("Average by Room"),
              ["Room", "Markers Used", "Total Days", "Average Days"],
              averages(by_room))
    for row in book["By Room"].iter_rows(min_row=2, min_col=6, max_col=6):
        row[0].number_format = "yyyy-mm-dd"
    book.save(path)


def add_order(room, item):
    savedata["orders"].append({"room": room, "item": item, "done": False})


REQUEST_HEADERS = ["Sent", "Room", "Item", "Type", "Request", "Days Lasted", "From"]


def log_request(data, action, room, item, days):
    """Add one row to the requests spreadsheet, creating it the first time."""
    from openpyxl import Workbook, load_workbook  # pip install openpyxl
    from openpyxl.styles import Font

    if os.path.exists(REQUESTS_PATH):
        book = load_workbook(REQUESTS_PATH)
        sheet = book.active
    else:
        os.makedirs(DATA_DIR, exist_ok=True)
        book = Workbook()
        sheet = book.active
        sheet.title = "Requests"
        sheet.append(REQUEST_HEADERS)
        for cell in sheet[1]:
            cell.font = Font(bold=True)
        for column, width in zip("ABCDEFG", (18, 8, 20, 9, 10, 12, 16)):
            sheet.column_dimensions[column].width = width
        sheet.freeze_panes = "A2"

    sent = data.get("sent")
    sent = datetime.datetime.fromtimestamp(sent) if isinstance(sent, (int, float)) else datetime.datetime.now()
    sheet.append([sent, int(room), item, "Marker" if item.lower() in COLORS else "Supply",
                  "Started" if action == "start" else "Ran out", days, str(data.get("from", "unknown"))])
    sheet.cell(sheet.max_row, 1).number_format = "yyyy-mm-dd hh:mm"
    book.save(REQUESTS_PATH)


def intercept(data):
    """Save a message's raw data and return only the text that should be displayed."""
    shown, found, raw = data["text"].partition(RAW_MARKER)
    if found and data.get("id") not in savedata["processed"]:
        try:
            action, room, item, value = raw.split(":")
            color = item.lower()
            if color not in TRACKED:
                raise ValueError(color)
            logged = None
            if action == "start":
                start_mark(room, color, int(value))
                logged = None, item
            elif action == "stop" and value == "?":
                # The sender didn't have a start time, so use this room's start day if there is one
                start = savedata["rooms"][room].get(color + "start", 0)
                days = None
                if start:
                    sent = data.get("sent")
                    sent = datetime.date.fromtimestamp(sent) if isinstance(sent, (int, float)) else datetime.date.today()
                    days = max(0, sent.toordinal() - start)
                    stop_mark(room, color, days)
                add_order(room, f"{item} marker" if color in COLORS else item)
                logged = days, item
            elif action == "stop" and 0 <= float(value) < 36500:
                stop_mark(room, color, float(value))
                add_order(room, f"{item} marker" if color in COLORS else item)
                logged = round(float(value), 2), item
        except (ValueError, KeyError):
            logged = None  # not valid raw data, so there's nothing to save
        if logged:
            try:
                log_request(data, action, room, logged[1], logged[0])
            except Exception as e:  # e.g. a damaged file; the request is still in rsave.json
                print(f"Couldn't add request to {REQUESTS_PATH}: {e}")
        savedata["processed"] = (savedata["processed"] + [data.get("id")])[-MAX_PROCESSED:]
        write_save()
    return shown


def write_save():
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(SAVE_PATH, "w") as f:
        json.dump(savedata, f, indent=4)
    try:
        os.chmod(SAVE_PATH, 0o600)  # file holds the app password
    except OSError:
        pass
def load_save():
    try:
        with open(SAVE_PATH, "r") as f:
            data = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return
    for room, values in data.get("rooms", {}).items():
        if room in savedata["rooms"]:
            for color in TRACKED:
                start = values.get(color + "start", 0)
                days = values.get(color + "time", [0, 0])
                if color in COMMON_COLORS or start or days != [0, 0]:  # skip unused colors
                    savedata["rooms"][room][color + "start"] = start
                    savedata["rooms"][room][color + "time"] = days
    savedata["orders"] = data.get("orders", [])
    savedata["processed"] = data.get("processed", [])
def ntfy_post(payload):
    """Publish a JSON payload to the topic; returns ntfy's id for the message."""
    data = json.dumps(payload).encode()
    with urllib.request.urlopen(NTFY_URL, data=data, timeout=15, context=SSL_CONTEXT) as resp:
        return json.load(resp)["id"]


def ntfy_read(line):
    #"""Parse one line of ntfy's stream into (ntfy id, our payload), or (None, None)."""
    try:
        event = json.loads(line)
        if event.get("event") == "message":
            payload = json.loads(event["message"])
            if isinstance(payload, dict):
                payload["sent"] = event.get("time")  # when ntfy got it, in Unix seconds
                return event["id"], payload
    except (ValueError, KeyError, AttributeError):
        pass  # keepalives, or something on the topic that isn't ours
    return None, None


def listen(inbox):
    """Background thread: hand notifications to the UI through a queue."""
    since = "12h"  # catch up on anything ntfy still has stored
    while True:
        try:
            # Read what was missed in one go, so already-stopped ones don't pop up
            url = f"{NTFY_URL}/json?since={since}&poll=1"
            with urllib.request.urlopen(url, timeout=30, context=SSL_CONTEXT) as resp:
                missed = [ntfy_read(line) for line in resp]
            stopped = {data.get("id") for _, data in missed if data and data.get("type") == "stop"}
            for event_id, data in missed:
                if event_id is not None:
                    since = event_id
                    data["handled"] = data.get("id") in stopped
                    inbox.put(data)

            url = f"{NTFY_URL}/json?since={since}"
            with urllib.request.urlopen(url, timeout=90, context=SSL_CONTEXT) as stream:
                inbox.put({"type": "status", "text": "Listening for notifications…"})
                for line in stream:
                    event_id, data = ntfy_read(line)
                    if event_id is not None:
                        since = event_id
                        data["handled"] = False
                        inbox.put(data)
        except NET_ERRORS:
            inbox.put({"type": "status", "text": "Can't reach ntfy, retrying…"})
        time.sleep(NTFY_RETRY)


class Receiver:
    def __init__(self, root):
        self.root = root
        self.seen = set()  # message ids already shown
        self.popups = {}  # message id -> (window, stop button)
        root.title("Marker Receiver")
        root.resizable(False, False)

        frame = ttk.Frame(root, padding=8)
        frame.pack(fill="both", expand=True)
        top = ttk.Frame(frame)
        top.pack(fill="x")
        self.status = ttk.Label(top, text="Connecting to ntfy…", font=FONT_BOLD)
        self.status.pack(side="left")
        ttk.Button(top, text="Hide", command=root.withdraw).pack(side="right")
        ttk.Button(top, text="Export", command=self.export).pack(side="right")
        try:
            # Clicking the Dock icon brings the hidden window back
            root.createcommand("tk::mac::ReopenApplication", root.deiconify)
        except tk.TclError:
            pass
        self.log = tk.Listbox(frame, width=36, height=4, font=FONT)
        self.log.pack(fill="both", expand=True, pady=(6, 0))

        orders_top = ttk.Frame(frame)
        orders_top.pack(fill="x", pady=(8, 0))
        ttk.Label(orders_top, text="Orders", font=FONT_BOLD).pack(side="left")
        ttk.Button(orders_top, text="Clear Done", command=self.clear_done).pack(side="right")
        self.orders = tk.Listbox(frame, width=36, height=5, font=FONT, exportselection=False)
        self.orders.pack(fill="both", expand=True, pady=(6, 0))
        self.orders.bind("<<ListboxSelect>>", self.toggle_order)
        self.show_orders()

        self.inbox = queue.Queue()
        threading.Thread(target=listen, args=(self.inbox,), daemon=True).start()
        self._poll()

    def _poll(self):
        # Tk isn't thread-safe, so the background threads only fill the queue
        while not self.inbox.empty():
            data = self.inbox.get()
            kind, msg_id = data.get("type"), data.get("id")
            if kind == "status":
                self.status.configure(text=data["text"])
            elif kind in ("msg", "start") and isinstance(data.get("text"), str) and msg_id not in self.seen:
                self.seen.add(msg_id)
                data["text"] = intercept(data)
                self.show_orders()
                if kind != "msg":
                    data["handled"] = True  # goes in the list, with no popup to stop
                self.notify(msg_id, data)
            elif kind == "stop":
                self.close_popup(msg_id)
            elif kind == "stop_failed" and msg_id in self.popups:
                self.popups[msg_id][1].state(["!disabled"])
                self.status.configure(text="Couldn't send Stop, try again")
        self.root.after(200, self._poll)

    def export(self):
        path = filedialog.asksaveasfilename(
            parent=self.root, title="Export Marker Report", defaultextension=".xlsx",
            filetypes=[("Excel file", "*.xlsx")],
            initialdir=os.path.expanduser("~/Desktop"),
            initialfile=f"Marker Report {datetime.date.today()}.xlsx")
        if not path:
            return
        try:
            export_report(path)
            messagebox.showinfo("Exported", f"Saved to:\n{path}", parent=self.root)
        except (OSError, ImportError) as e:
            messagebox.showerror("Export failed", str(e), parent=self.root)

    def show_orders(self):
        top = self.orders.yview()[0]
        self.orders.delete(0, "end")
        for i, order in enumerate(savedata["orders"]):
            box = "☑" if order["done"] else "☐"
            self.orders.insert("end", f"{box} Room {order['room']}: {order['item']}")
            if order["done"]:
                self.orders.itemconfigure(i, foreground="gray")
        self.orders.yview_moveto(top)

    def toggle_order(self, _):
        picked = self.orders.curselection()
        if picked:
            order = savedata["orders"][picked[0]]
            order["done"] = not order["done"]
            write_save()
            self.show_orders()

    def clear_done(self):
        savedata["orders"] = [o for o in savedata["orders"] if not o["done"]]
        write_save()
        self.show_orders()

    def notify(self, msg_id, data):
        sender, text = str(data.get("from", "unknown")), data["text"]
        stamp = datetime.datetime.now().strftime("%H:%M")
        handled = "  (handled)" if data.get("handled") and data.get("type") == "msg" else ""
        self.log.insert(0, f"{stamp}  [{sender}]  {text}{handled}")
        if data.get("handled"):
            return

        win = tk.Toplevel(self.root)
        win.title("Marker Notification")
        win.resizable(False, False)
        win.attributes("-topmost", True)
        win.protocol("WM_DELETE_WINDOW", lambda: None)  # only Stop Broadcast closes it
        body = ttk.Frame(win, padding=16)
        body.pack()
        ttk.Label(body, text=f"From {sender}", font=FONT_BOLD).pack(anchor="w")
        ttk.Label(body, text=text, font=FONT, wraplength=320).pack(anchor="w", pady=(6, 14))

        buttons = ttk.Frame(body)
        buttons.pack(fill="x")
        stop_btn = ttk.Button(buttons, text="Stop Broadcast", command=lambda: self.send_stop(msg_id))
        stop_btn.pack(side="right")
        self.popups[msg_id] = (win, stop_btn)
        win.bell()
        win.lift()

    def send_stop(self, msg_id):
        self.popups[msg_id][1].state(["disabled"])

        def post():
            try:
                ntfy_post({"type": "stop", "id": msg_id, "from": socket.gethostname()})
                self.inbox.put({"type": "stop", "id": msg_id})
            except NET_ERRORS:
                self.inbox.put({"type": "stop_failed", "id": msg_id})

        threading.Thread(target=post, daemon=True).start()

    def close_popup(self, msg_id):
        popup = self.popups.pop(msg_id, None)
        if popup and popup[0].winfo_exists():
            popup[0].destroy()


if __name__ == "__main__":
    load_save()
    root = tk.Tk()
    Receiver(root)
    root.mainloop()
    write_save()
    print("closed")
