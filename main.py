import re
import os
import sys
import json
import tempfile
from io import BytesIO

import requests
import tkinter as tk
from tkinter import messagebox
import customtkinter as ctk
from urllib.parse import quote
from PIL import Image
from rembg import remove

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
TITLE_INPUT = "Табло"

HOME_NAME_FIELD = "Хозяева.Text"
AWAY_NAME_FIELD = "Гости.Text"
HOME_SCORE_FIELD = "СчётХозяева.Text"
AWAY_SCORE_FIELD = "СчётГости.Text"

# На будущее для vMix — замени на реальные имена полей image в титре
HOME_LOGO_FIELD = "ЛогоХозяева.Source"
AWAY_LOGO_FIELD = "ЛогоГости.Source"

COLOR_ROW_MAP = {
    "Белый": 1,
    "Чёрный": 2,
    "Серый": 3,
    "Коричневый": 4,
    "Красный": 5,
    "Бордовый": 6,
    "Оранжевый": 7,
    "Жёлтый": 8,
    "Тёмно-Зеленый": 9,
    "Кислотно-Зеленый": 10,
    "Салатовый": 11,
    "Оливковый": 12,
    "Голубой": 13,
    "Синий": 14,
    "Фиолетовый": 15,
    "Розовый": 16,
    "Бледно-Розовый": 17,
}


def download_image_to_temp(url, prefix="logo_original"):
    if not url:
        return None

    try:
        r = requests.get(url, timeout=15)
        r.raise_for_status()

        img = Image.open(BytesIO(r.content)).convert("RGBA")

        temp_dir = os.path.join(tempfile.gettempdir(), "vmix_logo_temp")
        os.makedirs(temp_dir, exist_ok=True)

        out_path = os.path.join(temp_dir, f"{prefix}.png")
        img.save(out_path, "PNG")
        return out_path
    except Exception as e:
        print(f"Ошибка скачивания логотипа {url}: {e}")
        return None


def prepare_logo_file(url, prefix="logo_processed"):
    if not url:
        return None

    try:
        r = requests.get(url, timeout=15)
        r.raise_for_status()

        img = Image.open(BytesIO(r.content)).convert("RGBA")
        result = remove(img)

        temp_dir = os.path.join(tempfile.gettempdir(), "vmix_logo_temp")
        os.makedirs(temp_dir, exist_ok=True)

        out_path = os.path.join(temp_dir, f"{prefix}.png")
        result.save(out_path, "PNG")
        return out_path
    except Exception as e:
        print(f"Ошибка обработки логотипа {url}: {e}")
        return None


def extract_url(value):
    if not value:
        return ""
    m = re.search(r"\((https?://[^)]+)\)", value)
    if m:
        return m.group(1).strip()
    value = value.strip()
    if value.startswith("http://") or value.startswith("https://"):
        return value
    return value


def parse_match_block(block):
    match_line = re.search(r"Матч:\s*(.+)", block)
    video_line = re.search(r"URL видео:\s*(.+)", block)
    server_line = re.search(r"Сервер:\s*(.+)", block)
    key_line = re.search(r"Ключ:\s*(.+)", block)
    home_logo_line = re.search(r"Лого хозяев:\s*(.+)", block)
    away_logo_line = re.search(r"Лого гостей:\s*(.+)", block)
    home_abbr_line = re.search(r"Сокр\. хозяев:\s*(.+)", block)
    away_abbr_line = re.search(r"Сокр\. гостей:\s*(.+)", block)

    if not match_line:
        return None

    full_name = match_line.group(1).strip()
    server_url = server_line.group(1).strip() if server_line else ""
    stream_key = key_line.group(1).strip() if key_line else ""
    video_url = extract_url(video_line.group(1).strip()) if video_line else ""
    home_logo = extract_url(home_logo_line.group(1).strip()) if home_logo_line else ""
    away_logo = extract_url(away_logo_line.group(1).strip()) if away_logo_line else ""
    home_abbr = home_abbr_line.group(1).strip() if home_abbr_line else ""
    away_abbr = away_abbr_line.group(1).strip() if away_abbr_line else ""

    teams_part = full_name
    teams_part = re.sub(r"^.*?Day\s*\d+\.\s*", "", teams_part).strip()

    teams = re.search(r"(.+?)\s+-\s+(.+?)(?:\s*\(|$)", teams_part)
    team1 = teams.group(1).strip() if teams else ""
    team2 = teams.group(2).strip() if teams else ""

    return {
        "full_name": full_name,
        "video_url": video_url,
        "server": server_url,
        "key": stream_key,
        "home_logo": home_logo,
        "away_logo": away_logo,
        "home_abbr": home_abbr,
        "away_abbr": away_abbr,
        "team1": team1,
        "team2": team2,
    }


def clean_team_name(name):
    if not name:
        return ""
    bad_words = ["AFL", "FC", "LFC", "LFK", "FK", "CF", "SC", "AC", "МФК", "ФК"]
    parts = name.strip().split()
    cleaned = []
    for part in parts:
        if part.upper().replace(".", "") not in bad_words:
            cleaned.append(part)
    return " ".join(cleaned).strip()


def parse_all_matches(filepath):
    try:
        with open(filepath, encoding="utf-8") as f:
            content = f.read()
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось прочитать файл {filepath}:\n{e}")
        return []

    blocks = re.split(r"-{10,}", content)
    matches = []

    for block in blocks:
        item = parse_match_block(block)
        if item:
            matches.append(item)

    return matches


def get_color_row(team_name):
    try:
        with open(TEAM_DB_JSON, encoding="utf-8") as f:
            team_db = json.load(f)
    except Exception:
        team_db = {}

    clean_name = clean_team_name(team_name).lower()
    color_word = "Белый"

    for name_in_json, color in team_db.items():
        clean_json_name = clean_team_name(name_in_json).lower()
        if clean_json_name == clean_name:
            color_word = color
            break

    return color_word, COLOR_ROW_MAP.get(color_word, 1)


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
    r = requests.get(
        f"{api_url}?Function=DataSourceSelectRow&Value={quote(value)}",
        timeout=5
    )
    r.raise_for_status()
    return r


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
        self.root.title("vMix Управление Трансляциями")
        self.root.geometry("1200x900")

        self.matches = []
        self.radio_widgets = []
        self.selected_match_idx = ctk.IntVar(value=-1)

        self.home_original_ctk = None
        self.home_processed_ctk = None
        self.away_original_ctk = None
        self.away_processed_ctk = None

        ctk.CTkLabel(
            root,
            text="Управление трансляцией vMix / Тест логотипов",
            font=ctk.CTkFont(size=20, weight="bold")
        ).pack(pady=(15, 5))

        host_frame = ctk.CTkFrame(root, fg_color="transparent")
        host_frame.pack(padx=20, pady=(0, 5), fill="x")

        ctk.CTkLabel(
            host_frame,
            text="Адрес vMix (IP:Порт):",
            font=ctk.CTkFont(size=14)
        ).pack(side="left", padx=(0, 10))

        self.host_entry = ctk.CTkEntry(
            host_frame,
            width=220,
            font=ctk.CTkFont(size=14)
        )
        self.host_entry.insert(0, load_last_host())
        self.host_entry.pack(side="left")

        self.indicator = ctk.CTkLabel(
            host_frame,
            text="●",
            font=ctk.CTkFont(size=18),
            text_color="gray"
        )
        self.indicator.pack(side="left", padx=8)

        self.btn_ping = ctk.CTkButton(
            host_frame,
            text="Проверить vMix",
            width=120,
            font=ctk.CTkFont(size=13),
            fg_color="#555",
            hover_color="#333",
            command=self.ping_vmix
        )
        self.btn_ping.pack(side="left")

        main_frame = ctk.CTkFrame(root, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=20, pady=10)

        left_panel = ctk.CTkFrame(main_frame)
        left_panel.pack(side="left", fill="y", padx=(0, 10))

        right_panel = ctk.CTkFrame(main_frame)
        right_panel.pack(side="left", fill="both", expand=True)

        ctk.CTkLabel(
            left_panel,
            text="Выберите матч:",
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(anchor="w", padx=15, pady=(15, 5))

        self.scroll_frame = ctk.CTkScrollableFrame(left_panel, width=360, height=540)
        self.scroll_frame.pack(fill="both", expand=True, padx=15, pady=(0, 10))

        btn_frame = ctk.CTkFrame(left_panel, fg_color="transparent")
        btn_frame.pack(pady=(0, 15))

        ctk.CTkButton(
            btn_frame,
            text="Обновить список",
            command=self.load_matches,
            fg_color="#555555",
            hover_color="#333333",
            font=ctk.CTkFont(size=14),
            width=160
        ).pack(pady=5)

        ctk.CTkButton(
            btn_frame,
            text="Проверить лого",
            command=self.preview_logos,
            fg_color="#1f6aa5",
            hover_color="#144870",
            font=ctk.CTkFont(size=14, weight="bold"),
            width=160
        ).pack(pady=5)

        ctk.CTkButton(
            btn_frame,
            text="Отправить в vMix",
            command=self.send_to_vmix,
            fg_color="#28a745",
            hover_color="#218838",
            font=ctk.CTkFont(size=14, weight="bold"),
            width=160
        ).pack(pady=5)

        ctk.CTkLabel(
            right_panel,
            text="Предпросмотр логотипов",
            font=ctk.CTkFont(size=18, weight="bold")
        ).pack(anchor="w", padx=20, pady=(15, 10))

        self.preview_info = ctk.CTkLabel(
            right_panel,
            text="Выбери матч и нажми «Проверить лого»",
            text_color="gray",
            font=ctk.CTkFont(size=13)
        )
        self.preview_info.pack(anchor="w", padx=20, pady=(0, 10))

        preview_grid = ctk.CTkFrame(right_panel, fg_color="transparent")
        preview_grid.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        self.home_frame = ctk.CTkFrame(preview_grid)
        self.home_frame.pack(side="left", fill="both", expand=True, padx=(0, 10))

        self.away_frame = ctk.CTkFrame(preview_grid)
        self.away_frame.pack(side="left", fill="both", expand=True, padx=(10, 0))

        self.build_preview_column(
            self.home_frame,
            "Хозяева",
            "home"
        )
        self.build_preview_column(
            self.away_frame,
            "Гости",
            "away"
        )

        self.status = ctk.CTkLabel(
            root,
            text="Готово",
            text_color="gray",
            font=ctk.CTkFont(size=13)
        )
        self.status.pack(side="bottom", pady=8)

        self.load_matches()

    def build_preview_column(self, parent, title, side):
        ctk.CTkLabel(
            parent,
            text=title,
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(pady=(15, 8))

        ctk.CTkLabel(
            parent,
            text="Оригинал",
            font=ctk.CTkFont(size=13)
        ).pack()

        original_label = ctk.CTkLabel(parent, text="")
        original_label.pack(pady=(5, 15))

        ctk.CTkLabel(
            parent,
            text="После удаления фона",
            font=ctk.CTkFont(size=13)
        ).pack()

        processed_label = ctk.CTkLabel(parent, text="")
        processed_label.pack(pady=(5, 15))

        url_label = ctk.CTkTextbox(parent, width=300, height=85)
        url_label.pack(padx=10, pady=(0, 15))
        url_label.insert("1.0", "")
        url_label.configure(state="disabled")

        if side == "home":
            self.home_original_label = original_label
            self.home_processed_label = processed_label
            self.home_url_box = url_label
        else:
            self.away_original_label = original_label
            self.away_processed_label = processed_label
            self.away_url_box = url_label

    def set_textbox_value(self, textbox, text):
        textbox.configure(state="normal")
        textbox.delete("1.0", "end")
        textbox.insert("1.0", text)
        textbox.configure(state="disabled")

    def load_ctk_image(self, file_path, size=(180, 180), checkerboard=False):
        img = Image.open(file_path).convert("RGBA")
        img.thumbnail(size, Image.LANCZOS)

        if checkerboard:
            bg = Image.new("RGBA", size, (220, 220, 220, 255))
            cell = 18
            pixels = bg.load()

            for yy in range(size[1]):
                for xx in range(size[0]):
                    if ((xx // cell) + (yy // cell)) % 2 == 0:
                        pixels[xx, yy] = (235, 235, 235, 255)
                    else:
                        pixels[xx, yy] = (180, 180, 180, 255)

            x = (size[0] - img.width) // 2
            y = (size[1] - img.height) // 2
            bg.paste(img, (x, y), img)
            final_img = bg
        else:
            canvas = Image.new("RGBA", size, (255, 255, 255, 255))
            x = (size[0] - img.width) // 2
            y = (size[1] - img.height) // 2
            canvas.paste(img, (x, y), img)
            final_img = canvas

        return ctk.CTkImage(light_image=final_img, dark_image=final_img, size=size)

    def clear_previews(self):
        self.home_original_label.configure(image=None, text="")
        self.home_processed_label.configure(image=None, text="")
        self.away_original_label.configure(image=None, text="")
        self.away_processed_label.configure(image=None, text="")

        self.home_original_ctk = None
        self.home_processed_ctk = None
        self.away_original_ctk = None
        self.away_processed_ctk = None

        self.set_textbox_value(self.home_url_box, "")
        self.set_textbox_value(self.away_url_box, "")

    def get_current_api(self):
        return get_api_url(self.host_entry.get())

    def ping_vmix(self):
        api_url = self.get_current_api()
        try:
            r = requests.get(api_url, timeout=3)
            if r.status_code == 200:
                self.indicator.configure(text_color="#28a745")
                self.status.configure(
                    text=f"Подключено: {self.host_entry.get()}",
                    text_color="#28a745"
                )
                save_last_host(self.host_entry.get())
            else:
                self.indicator.configure(text_color="#ffcc00")
                self.status.configure(
                    text=f"vMix ответил с кодом {r.status_code}",
                    text_color="#ffcc00"
                )
        except Exception:
            self.indicator.configure(text_color="#ff4444")
            self.status.configure(
                text=f"vMix недоступен: {self.host_entry.get()}",
                text_color="#ff4444"
            )

    def load_matches(self):
        for widget in self.radio_widgets:
            widget.destroy()
        self.radio_widgets.clear()
        self.selected_match_idx.set(-1)
        self.clear_previews()

        self.matches = parse_all_matches(TXT_FILE)

        if not self.matches:
            lbl = ctk.CTkLabel(
                self.scroll_frame,
                text="Файл stream_keys.txt пуст или не найден",
                text_color="#ffcc00"
            )
            lbl.pack(pady=20)
            self.radio_widgets.append(lbl)
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
                font=ctk.CTkFont(size=14),
                hover_color="#28a745"
            )
            rb.pack(anchor="w", pady=6, padx=15)
            self.radio_widgets.append(rb)

        self.status.configure(
            text=f"Загружено матчей: {len(self.matches)}",
            text_color="gray"
        )

    def preview_logos(self):
        idx = self.selected_match_idx.get()
        if idx == -1:
            messagebox.showwarning("Внимание", "Сначала выберите матч из списка!")
            return

        self.clear_previews()
        match = self.matches[idx]

        self.preview_info.configure(
            text=f"Проверка логотипов: {match['team1']} vs {match['team2']}",
            text_color="#cccccc"
        )
        self.status.configure(text="Скачивание и обработка логотипов...", text_color="#3399ff")
        self.root.update()

        try:
            home_original = download_image_to_temp(match["home_logo"], "home_original")
            home_processed = prepare_logo_file(match["home_logo"], "home_processed")
            away_original = download_image_to_temp(match["away_logo"], "away_original")
            away_processed = prepare_logo_file(match["away_logo"], "away_processed")

            if home_original:
                self.home_original_ctk = self.load_ctk_image(home_original, checkerboard=False)
                self.home_original_label.configure(image=self.home_original_ctk, text="")
            else:
                self.home_original_label.configure(text="Не удалось загрузить", image=None)

            if home_processed:
                self.home_processed_ctk = self.load_ctk_image(home_processed, checkerboard=True)
                self.home_processed_label.configure(image=self.home_processed_ctk, text="")
            else:
                self.home_processed_label.configure(text="Не удалось обработать", image=None)

            if away_original:
                self.away_original_ctk = self.load_ctk_image(away_original, checkerboard=False)
                self.away_original_label.configure(image=self.away_original_ctk, text="")
            else:
                self.away_original_label.configure(text="Не удалось загрузить", image=None)

            if away_processed:
                self.away_processed_ctk = self.load_ctk_image(away_processed, checkerboard=True)
                self.away_processed_label.configure(image=self.away_processed_ctk, text="")
            else:
                self.away_processed_label.configure(text="Не удалось обработать", image=None)

            self.set_textbox_value(
                self.home_url_box,
                f"Команда: {match['team1']}\nURL: {match['home_logo']}"
            )
            self.set_textbox_value(
                self.away_url_box,
                f"Команда: {match['team2']}\nURL: {match['away_logo']}"
            )

            self.status.configure(text="Предпросмотр логотипов готов", text_color="#28a745")

        except Exception as e:
            self.status.configure(text="Ошибка предпросмотра", text_color="#ff4444")
            messagebox.showerror("Ошибка", str(e))

    def send_to_vmix(self):
        idx = self.selected_match_idx.get()
        if idx == -1:
            messagebox.showwarning("Внимание", "Сначала выберите матч из списка!")
            return

        match = self.matches[idx]
        api_url = self.get_current_api()

        self.status.configure(
            text=f"Отправка на {self.host_entry.get()}...",
            text_color="#3399ff"
        )
        self.root.update()

        try:
            home_color, home_row = get_color_row(match["team1"])
            away_color, away_row = get_color_row(match["team2"])

            home_logo_file = prepare_logo_file(match["home_logo"], "home_logo")
            away_logo_file = prepare_logo_file(match["away_logo"], "away_logo")

            if match["home_abbr"]:
                vmix_send(api_url, {
                    "Function": "SetText",
                    "Input": TITLE_INPUT,
                    "SelectedName": HOME_NAME_FIELD,
                    "Value": match["home_abbr"]
                })

            if match["away_abbr"]:
                vmix_send(api_url, {
                    "Function": "SetText",
                    "Input": TITLE_INPUT,
                    "SelectedName": AWAY_NAME_FIELD,
                    "Value": match["away_abbr"]
                })

            vmix_send(api_url, {
                "Function": "SetText",
                "Input": TITLE_INPUT,
                "SelectedName": HOME_SCORE_FIELD,
                "Value": "0"
            })

            vmix_send(api_url, {
                "Function": "SetText",
                "Input": TITLE_INPUT,
                "SelectedName": AWAY_SCORE_FIELD,
                "Value": "0"
            })

            if home_logo_file:
                vmix_send(api_url, {
                    "Function": "SetImage",
                    "Input": TITLE_INPUT,
                    "SelectedName": HOME_LOGO_FIELD,
                    "Value": home_logo_file
                })

            if away_logo_file:
                vmix_send(api_url, {
                    "Function": "SetImage",
                    "Input": TITLE_INPUT,
                    "SelectedName": AWAY_LOGO_FIELD,
                    "Value": away_logo_file
                })

            vmix_select_ds_row(api_url, HOME_DS_NAME, SHEET_NAME, home_row)
            vmix_select_ds_row(api_url, AWAY_DS_NAME, SHEET_NAME, away_row)

            if match["server"] and match["key"]:
                vmix_send(api_url, {
                    "Function": "StreamingSetURL",
                    "Value": f"0,{match['server']}"
                })
                vmix_send(api_url, {
                    "Function": "StreamingSetKey",
                    "Value": f"0,{match['key']}"
                })

            self.indicator.configure(text_color="#28a745")
            self.status.configure(
                text=f"[{self.host_entry.get()}] {match['team1']} vs {match['team2']} ({home_color}/{away_color})",
                text_color="#28a745"
            )

            save_last_host(self.host_entry.get())

        except requests.exceptions.RequestException:
            self.indicator.configure(text_color="#ff4444")
            self.status.configure(
                text=f"Ошибка: vMix недоступен по адресу {self.host_entry.get()}",
                text_color="#ff4444"
            )
            messagebox.showerror("Ошибка сети", f"Не удалось подключиться:\n{api_url}")
        except Exception as e:
            self.status.configure(text="Ошибка выполнения", text_color="#ff4444")
            messagebox.showerror("Ошибка", str(e))


if __name__ == "__main__":
    root = ctk.CTk()
    app = VmixApp(root)
    root.mainloop()
