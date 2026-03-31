import json
import os
import sys
import threading
import time
from datetime import datetime
from enum import IntEnum

import pystray
from PIL import Image, ImageDraw, ImageFont


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

def load_config():
    config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")
    with open(config_path, "r") as f:
        return json.load(f)


# ---------------------------------------------------------------------------
# Calendar view modes (matches Outlook OlCalendarViewMode)
# ---------------------------------------------------------------------------

class CalendarViewMode(IntEnum):
    DAY = 0
    WEEK = 1
    MONTH = 2
    MULTI_DAY = 3
    WORK_WEEK = 4


# ---------------------------------------------------------------------------
# Outlook Calendar COM wrapper
# ---------------------------------------------------------------------------

class OutlookCalendar:
    def __init__(self, folder_path):
        self.folder_path = folder_path
        self.outlook = None
        self.namespace = None
        self.folder = None

    def init(self):
        """Initialize or reinitialize Outlook COM connection."""
        import win32com.client
        try:
            if not self._is_outlook_valid():
                self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
            self.folder = self._find_folder(self.namespace.Folders, self.folder_path)
            return self.folder is not None
        except Exception as e:
            print(f"Outlook init failed: {e}")
            return False

    def init_if_needed(self):
        """Reinitialize if COM connection is stale."""
        if not self._is_outlook_valid() or not self._is_folder_valid():
            return self.init()
        return True

    def _is_outlook_valid(self):
        try:
            return self.outlook is not None and self.outlook.Name is not None
        except Exception:
            return False

    def _is_folder_valid(self):
        try:
            return self.folder is not None and self.folder.Name is not None
        except Exception:
            return False

    def _find_folder(self, folders, path):
        for i in range(folders.Count):
            folder = folders.Item(i + 1)  # COM collections are 1-indexed
            try:
                if folder.FolderPath and folder.FolderPath == path:
                    return folder
                result = self._find_folder(folder.Folders, path)
                if result:
                    return result
            except Exception:
                continue
        return None

    def get_todays_remaining_items(self):
        """Get today's upcoming calendar items."""
        if not self._is_folder_valid():
            return []

        try:
            now = datetime.now()
            end_of_today = now.replace(hour=23, minute=59, second=59)
            now_str = now.strftime("%m/%d/%Y %I:%M %p")
            end_str = end_of_today.strftime("%m/%d/%Y %I:%M %p")
            query = f"[Start] >= '{now_str}' And [Start] < '{end_str}'"

            items = self.folder.Items
            items.IncludeRecurrences = True
            items.Sort("[Start]")
            restricted = items.Restrict(query)

            # Must iterate; .Count is unreliable with recurrences
            result = []
            for item in restricted:
                result.append({
                    "subject": item.Subject,
                    "start": datetime(
                        item.Start.year, item.Start.month, item.Start.day,
                        item.Start.hour, item.Start.minute
                    ),
                    "end": datetime(
                        item.End.year, item.End.month, item.End.day,
                        item.End.hour, item.End.minute
                    ),
                })
            return result
        except Exception as e:
            print(f"Failed to get items: {e}")
            return []

    def get_item_count(self):
        """Get count of remaining items today."""
        return len(self.get_todays_remaining_items())

    def focus(self, view_mode=None):
        """Open Outlook and focus the calendar view."""
        if not self._is_folder_valid():
            return

        try:
            explorer = self._get_or_open_explorer()
            if not explorer:
                return

            explorer.CurrentFolder = self.folder

            if view_mode is not None:
                view = explorer.CurrentView
                OL_CALENDAR_VIEW = 2
                if view.ViewType == OL_CALENDAR_VIEW:
                    now_com = datetime.now()
                    view.GoToDate(now_com)
                    view.CalendarViewMode = int(view_mode)
                    view.Save()

            explorer.Activate()
        except Exception as e:
            print(f"Focus failed: {e}")

    def _get_or_open_explorer(self):
        """Get active explorer or open one."""
        try:
            explorer = self.outlook.ActiveExplorer()
            if explorer:
                return explorer
        except Exception:
            pass

        try:
            self.folder.Display()
            return self.outlook.ActiveExplorer()
        except Exception:
            pass

        try:
            OL_FOLDER_INBOX = 6
            inbox = self.namespace.GetDefaultFolder(OL_FOLDER_INBOX)
            inbox.Display()
            return self.outlook.ActiveExplorer()
        except Exception:
            return None

    def create_new_appointment(self):
        if not self._is_folder_valid():
            return
        try:
            item = self.folder.Items.Add(1)  # olAppointmentItem
            item.Display()
        except Exception as e:
            print(f"Create appointment failed: {e}")

    def create_new_meeting(self):
        if not self._is_folder_valid():
            return
        try:
            item = self.folder.Items.Add(1)  # olAppointmentItem
            item.MeetingStatus = 1  # olMeeting
            item.Display()
        except Exception as e:
            print(f"Create meeting failed: {e}")


# ---------------------------------------------------------------------------
# Icon rendering
# ---------------------------------------------------------------------------

def render_icon(count, badge_config):
    """Render a calendar tray icon with optional event count badge."""
    size = 64
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    # Calendar base icon
    # Top bar (red header)
    draw.rounded_rectangle([4, 4, 60, 20], radius=4, fill=(200, 60, 60))
    # Calendar body (white)
    draw.rectangle([4, 18, 60, 58], fill=(255, 255, 255))
    draw.rectangle([4, 56, 60, 60], fill=(240, 240, 240))

    # Day number
    today = str(datetime.now().day)
    try:
        font = ImageFont.truetype("segoeui.ttf", 24)
    except OSError:
        font = ImageFont.load_default()
    bbox = draw.textbbox((0, 0), today, font=font)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    draw.text(((size - tw) / 2, 22), today, fill=(50, 50, 50), font=font)

    # Badge
    if badge_config.get("enabled", True) and count != 0:
        bg = tuple(badge_config.get("backgroundColor", [220, 80, 80]))
        fg = tuple(badge_config.get("textColor", [255, 255, 255]))

        badge_text = "E" if count < 0 else str(min(count, 99))
        try:
            badge_font = ImageFont.truetype("segoeui.ttf", 16)
        except OSError:
            badge_font = ImageFont.load_default()

        bb = draw.textbbox((0, 0), badge_text, font=badge_font)
        btw = bb[2] - bb[0]
        badge_w = max(btw + 8, 20)
        bx = size - badge_w - 2
        by = 2
        draw.rounded_rectangle([bx, by, bx + badge_w, by + 20], radius=8, fill=bg)
        draw.text((bx + (badge_w - btw) / 2, by + 1), badge_text, fill=fg, font=badge_font)

    return img


# ---------------------------------------------------------------------------
# Tooltip / preview text
# ---------------------------------------------------------------------------

def build_tooltip(items, max_items):
    """Build tooltip text from calendar items. Max 127 chars for Windows."""
    if not items:
        return "No upcoming events"

    now = datetime.now()
    header = now.strftime("TODAY %A %m/%d").upper()
    lines = [header]

    for item in items[:max_items]:
        start = item["start"].strftime("%I:%M").lstrip("0")
        end = item["end"].strftime("%I:%M %p").lstrip("0")
        subject = item["subject"]
        line = f"{start}\u2013{end} {subject}"
        lines.append(line)

    if len(items) > max_items:
        lines.append("...")

    text = "\n".join(lines)
    # Windows tooltip max is 127 chars
    if len(text) > 127:
        text = text[:124] + "..."
    return text


# ---------------------------------------------------------------------------
# App
# ---------------------------------------------------------------------------

class CalendarTrayApp:
    def __init__(self):
        self.config = load_config()
        self.calendar = None
        self.icon = None
        self.lock = threading.Lock()

    def _com_action(self, action, *args):
        """Run a COM action on a new thread with its own COM apartment."""
        def run():
            import pythoncom
            pythoncom.CoInitialize()
            try:
                with self.lock:
                    self.calendar.init_if_needed()
                    action(*args)
            finally:
                pythoncom.CoUninitialize()
        threading.Thread(target=run, daemon=True).start()

    def on_focus_calendar(self, icon=None, item=None):
        self._com_action(self.calendar.focus)

    def on_focus_day(self, icon=None, item=None):
        self._com_action(self.calendar.focus, CalendarViewMode.DAY)

    def on_focus_week(self, icon=None, item=None):
        self._com_action(self.calendar.focus, CalendarViewMode.WEEK)

    def on_focus_month(self, icon=None, item=None):
        self._com_action(self.calendar.focus, CalendarViewMode.MONTH)

    def on_new_appointment(self, icon=None, item=None):
        self._com_action(self.calendar.create_new_appointment)

    def on_new_meeting(self, icon=None, item=None):
        self._com_action(self.calendar.create_new_meeting)

    def on_exit(self, icon, item):
        icon.stop()

    def build_menu(self):
        return pystray.Menu(
            pystray.MenuItem("Open Calendar", self.on_focus_calendar, default=True),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Day View", self.on_focus_day),
            pystray.MenuItem("Week View", self.on_focus_week),
            pystray.MenuItem("Month View", self.on_focus_month),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("New Appointment", self.on_new_appointment),
            pystray.MenuItem("New Meeting", self.on_new_meeting),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Exit", self.on_exit),
        )

    def refresh_loop(self):
        """Background thread: refresh Outlook data and update icon/tooltip."""
        import pythoncom
        pythoncom.CoInitialize()

        try:
            self.calendar = OutlookCalendar(self.config["outlook"]["folderPath"])
            self.calendar.init()

            interval = self.config.get("updateIntervalSeconds", 30)
            badge_config = self.config.get("badge", {})
            max_items = self.config.get("tooltip", {}).get("maxItems", 6)

            while True:
                try:
                    with self.lock:
                        self.calendar.init_if_needed()
                        items = self.calendar.get_todays_remaining_items()

                    count = len(items)
                    self.icon.icon = render_icon(count, badge_config)
                    self.icon.title = build_tooltip(items, max_items)

                except Exception as e:
                    print(f"Refresh error: {e}")
                    self.icon.icon = render_icon(-1, badge_config)
                    self.icon.title = "Error connecting to Outlook"

                time.sleep(interval)
        finally:
            pythoncom.CoUninitialize()

    def run(self):
        initial_icon = render_icon(0, self.config.get("badge", {}))
        self.icon = pystray.Icon(
            name="Outlook Calendar",
            icon=initial_icon,
            title="Loading...",
            menu=self.build_menu(),
        )

        refresh_thread = threading.Thread(target=self.refresh_loop, daemon=True)
        refresh_thread.start()

        self.icon.run()


if __name__ == "__main__":
    app = CalendarTrayApp()
    app.run()
