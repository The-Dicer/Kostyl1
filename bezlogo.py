import re
import os
import sys
import json
import threading
import queue
import requests
import webbrowser
from tkinter import messagebox
import customtkinter as ctk
from urllib.parse import quote

CONFIG_FILE = "config.json"

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DEFAULT_VMIX_HOST = "192.168.1.5:8088"
TEAM_DB_JSON = os.path.join(BASE_DIR, "teams_colors.json")
TXT_FILE = os.path.join(BASE_DIR, "stream_keys.txt")
SHEET_NAME = "Sheet1"

HOME_DS_NAME = "Цвет Хозяева"
AWAY_DS_NAME = "Цвет Гости"

VMIX_INPUTS = {
    "Табло": {
        "home_name": "Хозяева.Text",
        "away_name": "Гости.Text",
        "home_score": "Счёт Хозяева.Text",
        "away_score": "Счёт Гости.Text",
    },
    "08) Большое Табло": {
        "home_name": "Хозяева.Text",
        "away_name": "Гости.Text",
        "home_score": "Счёт Хозяева.Text",
        "away_score": "Счёт Гости.Text",
    },
    "Заставка2025": {
        "home_name": "Хозяева Имя.Text",
        "away_name": "Гости Имя.Text",
    },
    "21) Стингер": {
        # Стингер оставляем пустым
    },
}

COLOR_ROW_MAP = {
    "Белый": 1, "Чёрный": 2, "Серый": 3, "Коричневый": 4, "Красный": 5,
    "Бордовый": 6, "Оранжевый": 7, "Жёлтый": 8, "Тёмно-Зеленый": 9,
    "Кислотно-Зеленый": 10, "Салатовый": 11, "Оливковый": 12, "Голубой": 13,
    "Синий": 14, "Фиолетовый": 15, "Розовый": 16, "Бледно-Розовый": 17,
}

UI_COLORS = {
    "Белый": "#FFFFFF", "Чёрный": "#000000", "Серый": "#999999",
    "Коричневый": "#734d00", "Красный": "#FF0000", "Бордовый": "#990000",
    "Оранжевый": "#FF9933", "Жёлтый": "#FFFF00", "Тёмно-Зеленый": "#336633",
    "Кислотно-Зеленый": "#00FF99", "Салатовый": "#CCFF00", "Оливковый": "#999933",
    "Голубой": "#4AB4FF", "Синий": "#0066FF", "Фиолетовый": "#9900CC",
    "Розовый": "#FF00FF", "Бледно-Розовый": "#FFCCFF"
}


def extract_markdown_parts(value: str):
    value = (value or "").strip()
    m = re.match(r'^\[(.*?)\]\((.*?)\)(.*)$', value)
    if not m:
        return None
    return m.group(1).strip(), m.group(2).strip(), m.group(3).strip()


def extract_video_url(value: str) -> str:
    value = (value or "").strip()
    if not value:
        return ""
    parts = extract_markdown_parts(value)
    if parts:
        _, link_part, _ = parts
        return link_part.strip()
    return value


def build_match_link(value: str) -> str:
    url = extract_video_url(value)
    if not url:
        return ""
    m = re.search(r'rutube\.ru/video/([A-Za-z0-9]+)/?', url, re.IGNORECASE)
    if m:
        return f"https://studio.rutube.ru/stream/{m.group(1)}"
    return url


def parse_match_block(block):
    fields = {}
    for raw_line in block.splitlines():
        line = raw_line.strip()
        if not line or ":" not in line:
            continue
        key, value = line.split(":", 1)
        fields[key.strip()] = value.strip()

    full_name = fields.get("Матч", "")
    if not full_name:
        return None

    video_raw = fields.get("URL видео", "")
    home_abbr = fields.get("Сокр. хозяев", "")
    away_abbr = fields.get("Сокр. гостей", "")
    server_url = fields.get("Сервер", "")
    stream_key = fields.get("Ключ", "")

    teams_part = re.sub(r"^.*?Day\s*\d+\.\s*", "", full_name).strip()
    teams = re.search(r"(.+?)\s+-\s+(.+?)(?:\s*\(|$)", teams_part)
    team1 = teams.group(1).strip() if teams else ""
    team2 = teams.group(2).strip() if teams else ""

    return {
        "full_name": full_name,
        "video_url": extract_video_url(video_raw),
        "match_link": build_match_link(video_raw),
        "server": server_url,
        "key": stream_key,
        "home_abbr": home_abbr,
        "away_abbr": away_abbr,
        "team1": team1,
        "team2": team2,
    }


def parse_all_matches(filepath):
    try:
        with open(filepath, encoding="utf-8") as f:
            content = f.read()
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось прочитать файл {filepath}:\n{e}")
        return []

    blocks = [b.strip() for b in re.split(r"-{10,}", content) if b.strip()]
    matches = []
    for block in blocks:
        item = parse_match_block(block)
        if item:
            matches.append(item)
    return matches


def normalize_team_name(name):
    return (name or "").strip().lower()


def load_team_db():
    try:
        with open(TEAM_DB_JSON, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def get_color_row(team_name, team_db=None):
    if team_db is None:
        team_db = load_team_db()
    target = normalize_team_name(team_name)
    for name_in_json, color in team_db.items():
        if normalize_team_name(name_in_json) == target:
            return color, COLOR_ROW_MAP.get(color, 1)
    return "Белый", 1


def get_api_url(host_str):
    host_str = host_str.strip()
    if not host_str.startswith("http://") and not host_str.startswith("https://"):
        host_str = "http://" + host_str
    if not host_str.endswith("/API/"):
        host_str = host_str.rstrip("/") + "/API/"
    return host_str


def vmix_send(api_url, params):
    r = requests.get(api_url, params=params, timeout=5)
    r.raise_for_status()
    return r


def vmix_select_ds_row(api_url, ds_name, sheet_name, row_index):
    value = f"{ds_name},{sheet_name},{row_index}"
    r = requests.get(f"{api_url}?Function=DataSourceSelectRow&Value={quote(value)}", timeout=5)
    r.raise_for_status()
    return r


def send_to_all_vmix_inputs(api_url, match):
    for input_name, fields in VMIX_INPUTS.items():
        if match["home_abbr"] and "home_name" in fields:
            vmix_send(api_url, {
                "Function": "SetText",
                "Input": input_name,
                "SelectedName": fields["home_name"],
                "Value": match["home_abbr"]
            })

        if match["away_abbr"] and "away_name" in fields:
            vmix_send(api_url, {
                "Function": "SetText",
                "Input": input_name,
                "SelectedName": fields["away_name"],
                "Value": match["away_abbr"]
            })

        if "home_score" in fields:
            vmix_send(api_url, {
                "Function": "SetText",
                "Input": input_name,
                "SelectedName": fields["home_score"],
                "Value": "0"
            })

        if "away_score" in fields:
            vmix_send(api_url, {
                "Function": "SetText",
                "Input": input_name,
                "SelectedName": fields["away_score"],
                "Value": "0"
            })


def load_config():
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return {}


def save_config(data):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"Не удалось сохранить конфиг: {e}")


def load_last_host():
    return load_config().get("last_host", DEFAULT_VMIX_HOST)


def save_last_host(host_str):
    cfg = load_config()
    cfg["last_host"] = host_str
    save_config(cfg)


class VmixApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Костыли (Без логотипов)")
        self.root.geometry("980x600")
        self.root.minsize(920, 550)

        self.matches = []
        self.radio_widgets = []
        self.selected_match_idx = ctk.IntVar(value=-1)

        self.log_queue = queue.Queue()
        self.team_db = load_team_db()
        self.current_match_link = ""

        # УБРАЛ ГЛОБАЛЬНЫЙ ПЕРЕХВАТ КЛИКА МЫШИ!
        # self.root.bind("<Button-1>", lambda e: self.root.focus_set())

        top = ctk.CTkFrame(root, fg_color="transparent")
        top.pack(fill="x", padx=14, pady=(12, 8))

        ctk.CTkLabel(
            top,
            text="vMix Управление Трансляциями",
            font=ctk.CTkFont(size=20, weight="bold")
        ).pack(anchor="w")

        host_row = ctk.CTkFrame(top, fg_color="transparent")
        host_row.pack(fill="x", pady=(10, 0))

        ctk.CTkLabel(
            host_row,
            text="Адрес vMix:",
            font=ctk.CTkFont(size=14)
        ).pack(side="left", padx=(0, 8))

        self.host_entry = ctk.CTkEntry(host_row, width=260, font=ctk.CTkFont(size=14))
        self.host_entry.insert(0, load_last_host())
        self.host_entry.pack(side="left")

        self.host_entry.bind("<Control-c>", lambda e: e.widget.event_generate("<<Copy>>"))
        self.host_entry.bind("<Control-v>", lambda e: e.widget.event_generate("<<Paste>>"))

        self.indicator = ctk.CTkLabel(
            host_row,
            text="●",
            font=ctk.CTkFont(size=18),
            text_color="gray"
        )
        self.indicator.pack(side="left", padx=(8, 8))

        ctk.CTkButton(
            host_row,
            text="Проверить",
            width=96,
            font=ctk.CTkFont(size=13),
            fg_color="#555",
            hover_color="#333",
            command=self.ping_vmix
        ).pack(side="left", padx=(0, 8))

        main = ctk.CTkFrame(root, fg_color="transparent")
        main.pack(fill="both", expand=True, padx=14, pady=(0, 10))

        left = ctk.CTkFrame(main, width=340)
        left.pack(side="left", fill="y", padx=(0, 10))
        left.pack_propagate(False)

        right = ctk.CTkFrame(main)
        right.pack(side="left", fill="both", expand=True)

        ctk.CTkLabel(
            left,
            text="Матчи",
            font=ctk.CTkFont(size=15, weight="bold")
        ).pack(anchor="w", padx=14, pady=(14, 6))

        # === ВЕРХНИЙ РЯД (Глобальные действия с базой) ===
        actions_global = ctk.CTkFrame(left, fg_color="transparent")
        actions_global.pack(fill="x", padx=14, pady=(0, 8))

        ctk.CTkButton(
            actions_global,
            text="Обновить список",
            command=self.load_matches,
            fg_color="#555555",
            hover_color="#333333",
            font=ctk.CTkFont(size=13, weight="bold"),
            width=310
        ).pack(fill="x")

        # === НИЖНИЙ РЯД (Отправка) ===
        actions_local = ctk.CTkFrame(left, fg_color="transparent")
        actions_local.pack(fill="x", padx=14, pady=(0, 8))

        ctk.CTkButton(
            actions_local,
            text="Отправить в vMix",
            command=self.send_to_vmix,
            fg_color="#28a745",
            hover_color="#218838",
            font=ctk.CTkFont(size=14, weight="bold"),
            width=310
        ).pack(fill="x")

        self.scroll_frame = ctk.CTkScrollableFrame(left, width=310, height=480)
        self.scroll_frame.pack(fill="both", expand=True, padx=14, pady=(0, 14))

        ctk.CTkLabel(
            right,
            text="Настройки матча",
            font=ctk.CTkFont(size=17, weight="bold")
        ).pack(anchor="w", padx=16, pady=(14, 4))

        self.match_link_button = ctk.CTkButton(
            right,
            text="Открыть матч в браузере",
            command=self.open_current_match_link,
            fg_color="transparent",
            hover_color="#1f1f1f",
            text_color="#4ea3ff",
            font=ctk.CTkFont(size=13, underline=True),
            anchor="w",
            state="disabled"
        )
        self.match_link_button.pack(fill="x", padx=16, pady=(0, 10))

        colors_grid = ctk.CTkFrame(right, fg_color="transparent")
        colors_grid.pack(fill="x", padx=16, pady=(0, 10))

        self.home_frame = ctk.CTkFrame(colors_grid)
        self.home_frame.pack(side="left", fill="both", expand=True, padx=(0, 8))

        self.away_frame = ctk.CTkFrame(colors_grid)
        self.away_frame.pack(side="left", fill="both", expand=True, padx=(8, 0))

        self.build_color_column(self.home_frame, "Хозяева", "home")
        self.build_color_column(self.away_frame, "Гости", "away")

        ctk.CTkLabel(
            right,
            text="Лог",
            font=ctk.CTkFont(size=15, weight="bold")
        ).pack(anchor="w", padx=16, pady=(4, 4))

        self.log_box = ctk.CTkTextbox(right, height=200)
        self.log_box.pack(fill="both", expand=True, padx=16, pady=(0, 12))
        self.log_box.configure(state="disabled")

        self.root.after(100, self.process_log_queue)

        bottom = ctk.CTkFrame(root, fg_color="transparent")
        bottom.pack(fill="x", padx=14, pady=(0, 12))

        self.status = ctk.CTkLabel(
            bottom,
            text="Готово",
            text_color="gray",
            font=ctk.CTkFont(size=13)
        )
        self.status.pack(side="left")

        # Привязка умных локальных горячих клавиш
        self.root.bind("<Up>", lambda e: self.handle_arrow_keys(-1, e))
        self.root.bind("<Down>", lambda e: self.handle_arrow_keys(1, e))
        self.root.bind("<Return>", lambda e: self.handle_enter_key(e))

        self.load_matches()

    def is_typing(self):
        """Проверяет, находится ли фокус внутри текстового поля"""
        focused = self.root.focus_get()
        return focused and type(focused).__name__ == 'Entry'

    def handle_arrow_keys(self, step, event):
        """Перехват стрелочек. Запрещаем переключение матчей, если вводим IP."""
        if self.is_typing():
            return
        self.navigate_matches(step)

    def handle_enter_key(self, event):
        """Перехват Enter. Если в поле IP - проверяем соединение. Если нет - отправляем матч."""
        if self.is_typing():
            self.ping_vmix()
        else:
            self.send_to_vmix()

    def build_color_column(self, parent, title, side):
        self.team_label = ctk.CTkLabel(
            parent,
            text=title,
            font=ctk.CTkFont(size=15, weight="bold")
        )
        self.team_label.pack(pady=(12, 4))

        color_frame = ctk.CTkFrame(parent, fg_color="transparent")
        color_frame.pack(pady=(4, 12))

        color_swatch = ctk.CTkFrame(color_frame, width=20, height=20, corner_radius=10, border_width=1,
                                    border_color="gray")
        color_swatch.pack(side="left", padx=(0, 10))

        color_var = ctk.StringVar(value="Белый")

        def update_swatch(choice, swatch=color_swatch):
            swatch.configure(fg_color=UI_COLORS.get(choice, "#FFFFFF"))

        color_menu = ctk.CTkOptionMenu(
            color_frame,
            values=list(COLOR_ROW_MAP.keys()),
            variable=color_var,
            command=update_swatch,
            width=140
        )
        color_menu.pack(side="left")
        color_menu.configure(state="disabled")

        if side == "home":
            self.home_team_label = self.team_label
            self.home_color_var = color_var
            self.home_color_swatch = color_swatch
            self.home_color_menu = color_menu
        else:
            self.away_team_label = self.team_label
            self.away_color_var = color_var
            self.away_color_swatch = color_swatch
            self.away_color_menu = color_menu

    def add_log(self, text):
        self.log_queue.put(text)

    def process_log_queue(self):
        while not self.log_queue.empty():
            msg = self.log_queue.get()
            self.log_box.configure(state="normal")
            self.log_box.insert("end", msg + "\n")
            self.log_box.see("end")
            self.log_box.configure(state="disabled")
        self.root.after(100, self.process_log_queue)

    def navigate_matches(self, step):
        if not self.matches:
            return

        current_idx = self.selected_match_idx.get()
        next_idx = 0 if current_idx == -1 else current_idx + step

        if next_idx < 0:
            next_idx = 0
        elif next_idx >= len(self.matches):
            next_idx = len(self.matches) - 1

        self.selected_match_idx.set(next_idx)
        self.on_match_select()

    def on_match_select(self):
        idx = self.selected_match_idx.get()
        if idx == -1 or not self.matches:
            return

        match = self.matches[idx]

        self.home_team_label.configure(text=match['team1'] if match['team1'] else "Хозяева")
        self.away_team_label.configure(text=match['team2'] if match['team2'] else "Гости")

        home_color, _ = get_color_row(match["team1"], self.team_db)
        away_color, _ = get_color_row(match["team2"], self.team_db)

        self.home_color_var.set(home_color)
        self.home_color_swatch.configure(fg_color=UI_COLORS.get(home_color, "#FFFFFF"))
        self.home_color_menu.configure(state="normal")

        self.away_color_var.set(away_color)
        self.away_color_swatch.configure(fg_color=UI_COLORS.get(away_color, "#FFFFFF"))
        self.away_color_menu.configure(state="normal")

        self.current_match_link = match.get("match_link", "")
        if self.current_match_link:
            self.match_link_button.configure(state="normal", text="Открыть матч в Rutube Studio")
        else:
            self.match_link_button.configure(state="disabled", text="Ссылка на матч недоступна")

        self.status.configure(text=f"Выбран матч: {match['team1']} vs {match['team2']}", text_color="gray")

    def open_current_match_link(self):
        if self.current_match_link:
            webbrowser.open_new_tab(self.current_match_link)

    def get_current_api(self):
        return get_api_url(self.host_entry.get())

    def ping_vmix(self):
        self.root.focus_set()  # Сбрасываем фокус с поля адреса после проверки
        api_url = self.get_current_api()
        try:
            r = requests.get(api_url, timeout=3)
            if r.status_code == 200:
                self.indicator.configure(text_color="#28a745")
                self.status.configure(text=f"Подключено: {self.host_entry.get()}", text_color="#28a745")
                save_last_host(self.host_entry.get())
                self.add_log(f"vMix доступен: {self.host_entry.get()}")
            else:
                self.indicator.configure(text_color="#ffcc00")
                self.status.configure(text=f"vMix ответил с кодом {r.status_code}", text_color="#ffcc00")
                self.add_log(f"vMix ответил с кодом {r.status_code}")
        except Exception:
            self.indicator.configure(text_color="#ff4444")
            self.status.configure(text=f"vMix недоступен: {self.host_entry.get()}", text_color="#ff4444")
            self.add_log(f"vMix недоступен: {self.host_entry.get()}")

    def load_matches(self):
        self.root.focus_set()  # Сбрасываем фокус
        self.team_db = load_team_db()

        for widget in self.radio_widgets:
            widget.destroy()
        self.radio_widgets.clear()
        self.selected_match_idx.set(-1)

        self.home_color_menu.configure(state="disabled")
        self.away_color_menu.configure(state="disabled")
        self.match_link_button.configure(state="disabled", text="Открыть матч в браузере")
        self.home_team_label.configure(text="Хозяева")
        self.away_team_label.configure(text="Гости")

        self.matches = parse_all_matches(TXT_FILE)

        if not self.matches:
            lbl = ctk.CTkLabel(
                self.scroll_frame,
                text="Файл stream_keys.txt пуст или не найден",
                text_color="#ffcc00"
            )
            lbl.pack(pady=20)
            self.radio_widgets.append(lbl)
            self.add_log("Файл stream_keys.txt пуст или не найден")
            self.status.configure(text="Матчи не загружены", text_color="#ffcc00")
            return

        for i, m in enumerate(self.matches):
            try:
                tournament = m["full_name"].split(".")[1].strip()
            except Exception:
                tournament = "Матч"

            display_name = f"{m['team1']} vs {m['team2']}   ({tournament})"
            rb = ctk.CTkRadioButton(
                self.scroll_frame,
                text=display_name,
                variable=self.selected_match_idx,
                value=i,
                command=self.on_match_select,
                font=ctk.CTkFont(size=14),
                hover_color="#28a745"
            )
            rb.pack(anchor="w", pady=6, padx=12)
            self.radio_widgets.append(rb)

        if self.matches:
            self.selected_match_idx.set(0)
            self.on_match_select()

        self.status.configure(text=f"Загружено матчей: {len(self.matches)}", text_color="gray")
        self.add_log(f"Загружено матчей: {len(self.matches)}")

    def send_to_vmix(self):
        self.root.focus_set()  # Сбрасываем фокус
        idx = self.selected_match_idx.get()
        if idx == -1:
            messagebox.showwarning("Внимание", "Сначала выберите матч из списка!")
            return

        match = self.matches[idx]
        api_url = self.get_current_api()

        home_color_ui = self.home_color_var.get()
        away_color_ui = self.away_color_var.get()

        self.status.configure(
            text=f"Отправка во все инпуты → {self.host_entry.get()}...",
            text_color="#3399ff"
        )
        self.add_log(f"Старт отправки в vMix: {match['team1']} vs {match['team2']}")

        threading.Thread(target=self.send_to_vmix_worker, args=(match, api_url, home_color_ui, away_color_ui),
                         daemon=True).start()

    def send_to_vmix_worker(self, match, api_url, home_color_ui, away_color_ui):
        try:
            home_row = COLOR_ROW_MAP.get(home_color_ui, 1)
            away_row = COLOR_ROW_MAP.get(away_color_ui, 1)

            self.add_log(f"Цвет хозяев (выбранный): {home_color_ui}, строка {home_row}")
            self.add_log(f"Цвет гостей (выбранный): {away_color_ui}, строка {away_row}")

            self.add_log("Отправляю данные во все инпуты vMix...")
            send_to_all_vmix_inputs(api_url, match)

            self.add_log("Переключаю строки DataSource...")
            vmix_select_ds_row(api_url, HOME_DS_NAME, SHEET_NAME, home_row)
            vmix_select_ds_row(api_url, AWAY_DS_NAME, SHEET_NAME, away_row)

            if match["server"] and match["key"]:
                self.add_log("Обновляю Streaming URL и Key...")
                vmix_send(api_url, {"Function": "StreamingSetURL", "Value": f"0,{match['server']}"})
                vmix_send(api_url, {"Function": "StreamingSetKey", "Value": f"0,{match['key']}"})

            self.root.after(0, lambda: self.indicator.configure(text_color="#28a745"))
            self.root.after(
                0,
                lambda: self.status.configure(
                    text=f"Успешно: {match['team1']} vs {match['team2']} ({home_color_ui}/{away_color_ui})",
                    text_color="#28a745"
                )
            )
            self.root.after(0, lambda: save_last_host(self.host_entry.get()))
            self.add_log("Отправка завершена успешно.")

            current_idx = self.selected_match_idx.get()
            if current_idx + 1 < len(self.matches):
                next_idx = current_idx + 1
                self.root.after(0, lambda: self.selected_match_idx.set(next_idx))
                self.root.after(0, self.on_match_select)
                self.add_log(
                    f"Автоматически выбран следующий матч: {self.matches[next_idx]['team1']} vs {self.matches[next_idx]['team2']}")
            else:
                self.add_log("Это был последний матч в списке.")

        except requests.exceptions.RequestException:
            self.add_log(f"Ошибка сети: vMix недоступен: {self.host_entry.get()}")
            self.root.after(0, lambda: self.indicator.configure(text_color="#ff4444"))
            self.root.after(0, lambda: self.status.configure(text="Ошибка: vMix недоступен", text_color="#ff4444"))
            self.root.after(0, lambda: messagebox.showerror("Ошибка сети", f"Не удалось подключиться:\n{api_url}"))
        except Exception as e:
            self.add_log(f"Ошибка выполнения: {e}")
            self.root.after(0, lambda: self.status.configure(text="Ошибка выполнения", text_color="#ff4444"))
            self.root.after(0, lambda: messagebox.showerror("Ошибка", str(e)))


if __name__ == "__main__":
    root = ctk.CTk()
    app = VmixApp(root)
    root.mainloop()