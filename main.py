import re
import json
import requests
import tkinter as tk
from tkinter import messagebox
import customtkinter as ctk
from urllib.parse import quote

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# --- Настройки ---
DEFAULT_VMIX_HOST = "192.168.1.5:8088"
TEAM_DB_JSON = "teams_colors.json"
TXT_FILE = "stream_keys.txt"
SHEET_NAME = "Sheet1"
HOME_DS_NAME = "Цвет Хозяева"
AWAY_DS_NAME = "Цвет Гости"
TITLE_INPUT = "Табло"

COLOR_ROW_MAP = {
    "Белый": 1, "Чёрный": 2, "Серый": 3, "Коричневый": 4,
    "Красный": 5, "Бордовый": 6, "Оранжевый": 7, "Жёлтый": 8,
    "Тёмно-Зеленый": 9, "Кислотно-Зеленый": 10, "Салатовый": 11,
    "Оливковый": 12, "Голубой": 13, "Синий": 14, "Фиолетовый": 15,
    "Розовый": 16, "Бледно-Розовый": 17,
}


# --- Функция очистки названий команд ---
def clean_team_name(name):
    if not name:
        return ""
    bad_words = ["FC", "LFC", "LFK", "FK", "CF", "SC", "AC", "МФК", "ФК"]
    parts = name.strip().split()
    cleaned = []
    for part in parts:
        if part.upper().replace(".", "") not in bad_words:
            cleaned.append(part)
    return " ".join(cleaned).strip()


# --- Логика работы с данными ---
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
        match_line  = re.search(r"Матч:\s*(.+)", block)
        server_line = re.search(r"Сервер:\s*(.+)", block)
        key_line    = re.search(r"Ключ:\s*(.+)", block)

        if not match_line:
            continue

        match_name = match_line.group(1).strip()
        server_url = server_line.group(1).strip() if server_line else ""
        stream_key = key_line.group(1).strip() if key_line else ""

        teams_str = match_name.split("Day 2.")[-1].strip() if "Day 2." in match_name else match_name.split(".")[-1].strip()
        teams = re.search(r"([^-]+)\s*-\s*(.+?)(?:\s*\(|$)", teams_str)

        team1 = clean_team_name(teams.group(1).strip() if teams else "")
        team2 = clean_team_name(teams.group(2).strip() if teams else "")

        matches.append({
            "full_name": match_name,
            "team1": team1, "team1_abbr": team1[:3].upper() if team1 else "",
            "team2": team2, "team2_abbr": team2[:3].upper() if team2 else "",
            "server": server_url, "key": stream_key,
        })

    return matches


def get_color_row(team_name):
    try:
        with open(TEAM_DB_JSON, encoding="utf-8") as f:
            team_db = json.load(f)
    except:
        team_db = {}

    color_word = "Белый"
    for name, color in team_db.items():
        if name.strip().lower() == team_name.strip().lower():
            color_word = color
            break

    return color_word, COLOR_ROW_MAP.get(color_word, 1)


# --- Логика vMix ---
def get_api_url(host_str):
    host_str = host_str.strip()
    if not host_str.startswith("http://") and not host_str.startswith("https://"):
        host_str = "http://" + host_str
    if not host_str.endswith("/API/"):
        host_str = host_str.rstrip("/") + "/API/"
    return host_str


def vmix_send(api_url, params):
    requests.get(api_url, params=params, timeout=5)

def vmix_select_ds_row(api_url, ds_name, sheet_name, row_index):
    value = f"{ds_name},{sheet_name},{row_index}"
    requests.get(f"{api_url}?Function=DataSourceSelectRow&Value={quote(value)}", timeout=5)


# --- Графический интерфейс ---
class VmixApp:
    def __init__(self, root):
        self.root = root
        self.root.title("vMix Управление Трансляциями")
        self.root.geometry("700x560")

        self.matches = []
        self.radio_widgets = []
        self.selected_match_idx = ctk.IntVar(value=-1)

        # Заголовок
        ctk.CTkLabel(root, text="Управление трансляцией vMix", font=ctk.CTkFont(size=18, weight="bold")).pack(pady=(20, 5))

        # --- Ввод IP и порта ---
        host_frame = ctk.CTkFrame(root, fg_color="transparent")
        host_frame.pack(padx=20, pady=(0, 5), fill="x")

        ctk.CTkLabel(host_frame, text="Адрес vMix (IP:Порт):", font=ctk.CTkFont(size=14)).pack(side="left", padx=(0, 10))

        self.host_entry = ctk.CTkEntry(host_frame, width=200, font=ctk.CTkFont(size=14))
        self.host_entry.insert(0, DEFAULT_VMIX_HOST)
        self.host_entry.pack(side="left")

        # Индикатор статуса подключения
        self.indicator = ctk.CTkLabel(host_frame, text="●", font=ctk.CTkFont(size=18), text_color="gray")
        self.indicator.pack(side="left", padx=8)
        self.btn_ping = ctk.CTkButton(host_frame, text="Проверить", width=100, font=ctk.CTkFont(size=13),
                                      fg_color="#555", hover_color="#333", command=self.ping_vmix)
        self.btn_ping.pack(side="left")

        # --- Список матчей ---
        ctk.CTkLabel(root, text="Выберите матч:", font=ctk.CTkFont(size=14)).pack(anchor="w", padx=20)
        self.scroll_frame = ctk.CTkScrollableFrame(root, width=600, height=270)
        self.scroll_frame.pack(fill="both", expand=True, padx=20, pady=5)

        # --- Кнопки ---
        btn_frame = ctk.CTkFrame(root, fg_color="transparent")
        btn_frame.pack(pady=12)

        ctk.CTkButton(btn_frame, text="Обновить список", command=self.load_matches,
                      fg_color="#555555", hover_color="#333333", font=ctk.CTkFont(size=14)).pack(side="left", padx=10)

        ctk.CTkButton(btn_frame, text="Отправить в vMix", command=self.send_to_vmix,
                      fg_color="#28a745", hover_color="#218838",
                      font=ctk.CTkFont(size=14, weight="bold"), width=200).pack(side="left", padx=10)

        # --- Статус ---
        self.status = ctk.CTkLabel(root, text="Готово", text_color="gray", font=ctk.CTkFont(size=13))
        self.status.pack(side="bottom", pady=8)

        self.load_matches()

    def get_current_api(self):
        return get_api_url(self.host_entry.get())

    def ping_vmix(self):
        """Проверяет доступность vMix на указанном адресе"""
        api_url = self.get_current_api()
        try:
            r = requests.get(api_url, timeout=3)
            if r.status_code == 200:
                self.indicator.configure(text_color="#28a745")  # Зеленый
                self.status.configure(text=f"Подключено: {self.host_entry.get()}", text_color="#28a745")
            else:
                self.indicator.configure(text_color="#ffcc00")
                self.status.configure(text=f"vMix ответил с кодом {r.status_code}", text_color="#ffcc00")
        except:
            self.indicator.configure(text_color="#ff4444")  # Красный
            self.status.configure(text=f"vMix недоступен: {self.host_entry.get()}", text_color="#ff4444")

    def load_matches(self):
        for widget in self.radio_widgets:
            widget.destroy()
        self.radio_widgets.clear()
        self.selected_match_idx.set(-1)

        self.matches = parse_all_matches(TXT_FILE)

        if not self.matches:
            lbl = ctk.CTkLabel(self.scroll_frame, text="Файл stream_keys.txt пуст или не найден", text_color="#ffcc00")
            lbl.pack(pady=20)
            self.radio_widgets.append(lbl)
            return

        for i, m in enumerate(self.matches):
            try:
                tournament = m['full_name'].split('.')[1].strip()
            except:
                tournament = "Матч"

            display_name = f"{m['team1']} vs {m['team2']}   ({tournament})"
            rb = ctk.CTkRadioButton(self.scroll_frame, text=display_name,
                                    variable=self.selected_match_idx, value=i,
                                    font=ctk.CTkFont(size=14), hover_color="#28a745")
            rb.pack(anchor="w", pady=6, padx=15)
            self.radio_widgets.append(rb)

        self.status.configure(text=f"Загружено матчей: {len(self.matches)}", text_color="gray")

    def send_to_vmix(self):
        idx = self.selected_match_idx.get()
        if idx == -1:
            messagebox.showwarning("Внимание", "Сначала выберите матч из списка!")
            return

        match = self.matches[idx]
        api_url = self.get_current_api()

        self.status.configure(text=f"Отправка на {self.host_entry.get()}...", text_color="#3399ff")
        self.root.update()

        try:
            home_color, home_row = get_color_row(match["team1"])
            away_color, away_row = get_color_row(match["team2"])

            vmix_send(api_url, {"Function": "SetText", "Input": TITLE_INPUT, "SelectedName": "Хозяева.Text", "Value": match["team1_abbr"]})
            vmix_send(api_url, {"Function": "SetText", "Input": TITLE_INPUT, "SelectedName": "Гости.Text",   "Value": match["team2_abbr"]})

            vmix_select_ds_row(api_url, HOME_DS_NAME, SHEET_NAME, home_row)
            vmix_select_ds_row(api_url, AWAY_DS_NAME, SHEET_NAME, away_row)

            if match["server"] and match["key"]:
                vmix_send(api_url, {"Function": "StreamingSetURL", "Value": f"0,{match['server']}"})
                vmix_send(api_url, {"Function": "StreamingSetKey", "Value": f"0,{match['key']}"})

            self.indicator.configure(text_color="#28a745")
            self.status.configure(
                text=f"[{self.host_entry.get()}]  {match['team1_abbr']} ({home_color}) vs {match['team2_abbr']} ({away_color})",
                text_color="#28a745")

        except requests.exceptions.RequestException:
            self.indicator.configure(text_color="#ff4444")
            self.status.configure(text=f"Ошибка: vMix недоступен по адресу {self.host_entry.get()}", text_color="#ff4444")
            messagebox.showerror("Ошибка сети", f"Не удалось подключиться:\n{api_url}")
        except Exception as e:
            self.status.configure(text="Ошибка выполнения", text_color="#ff4444")
            messagebox.showerror("Ошибка", str(e))


if __name__ == "__main__":
    root = ctk.CTk()
    app = VmixApp(root)
    root.mainloop()