import re
import os
import sys
import json
import tempfile
from io import BytesIO
import threading
import queue
import requests
import webbrowser
from tkinter import messagebox
import customtkinter as ctk
from urllib.parse import quote
from collections import deque
from PIL import Image, ImageFilter

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
    "07) Табло": {
        "home_logo": "ЛогоХозяева.Source",
        "away_logo": "ЛогоГости.Source",
        "home_name": "Хозяева.Text",
        "away_name": "Гости.Text",
        "home_score": "Счёт Хозяева.Text",
        "away_score": "Счёт Гости.Text",
    },
    "08) Большое Табло": {
        "home_logo": "ЛогоХозяева.Source",
        "away_logo": "ЛогоГости.Source",
        "home_name": "Хозяева.Text",
        "away_name": "Гости.Text",
        "home_score": "Счёт Хозяева.Text",
        "away_score": "Счёт Гости.Text",
    },
    "Заставка2025": {
        "home_logo": "Лого Хозяева.Source",
        "away_logo": "Лого Гости.Source",
        "home_name": "Хозяева Имя.Text",
        "away_name": "Гости Имя.Text",
    },
    "21) Стингер": {
        "home_logo": "ЛогоХ.Source",
        "away_logo": "ЛогоГ.Source",
    },
}

COLOR_ROW_MAP = {
    "Белый": 1, "Чёрный": 2, "Серый": 3, "Коричневый": 4, "Красный": 5,
    "Бордовый": 6, "Оранжевый": 7, "Жёлтый": 8, "Тёмно-Зеленый": 9,
    "Кислотно-Зеленый": 10, "Салатовый": 11, "Оливковый": 12, "Голубой": 13,
    "Синий": 14, "Фиолетовый": 15, "Розовый": 16, "Бледно-Розовый": 17,
}

UI_COLORS = {
    "Белый": "#FFFFFF",
    "Чёрный": "#000000",
    "Серый": "#999999",
    "Коричневый": "#734d00",
    "Красный": "#FF0000",
    "Бордовый": "#990000",
    "Оранжевый": "#FF9933",
    "Жёлтый": "#FFFF00",
    "Тёмно-Зеленый": "#336633",
    "Кислотно-Зеленый": "#00FF99",
    "Салатовый": "#CCFF00",
    "Оливковый": "#999933",
    "Голубой": "#4AB4FF",
    "Синий": "#0066FF",
    "Фиолетовый": "#9900CC",
    "Розовый": "#FF00FF",
    "Бледно-Розовый": "#FFCCFF"
}


def needs_background_removal(img: Image.Image, white_threshold=220) -> bool:
    img = img.convert("RGBA")
    w, h = img.size

    border_pixels = []
    border_pixels.extend([(x, 0) for x in range(w)])
    border_pixels.extend([(x, h - 1) for x in range(w)])
    border_pixels.extend([(0, y) for y in range(1, h - 1)])
    border_pixels.extend([(w - 1, y) for y in range(1, h - 1)])

    white_opaque_count = 0
    transparent_count = 0

    for x, y in border_pixels:
        r, g, b, a = img.getpixel((x, y))
        if a < 50:
            transparent_count += 1
        elif r >= white_threshold and g >= white_threshold and b >= white_threshold:
            white_opaque_count += 1

    total_border = len(border_pixels)
    if transparent_count > total_border * 0.3:
        return False
    return white_opaque_count > total_border * 0.15


def remove_white_fringe(img: Image.Image) -> Image.Image:
    r, g, b, a = img.split()
    a = a.filter(ImageFilter.MinFilter(3))
    a = a.filter(ImageFilter.GaussianBlur(0.5))
    return Image.merge("RGBA", (r, g, b, a))


def process_floodfill(img_raw: Image.Image, white_threshold=220, alpha_threshold=10) -> Image.Image:
    img = img_raw.convert("RGBA")
    w, h = img.size
    pixels = img.load()

    def is_light_pixel(px):
        r, g, b, a = px
        if a <= alpha_threshold:
            return False
        return r >= white_threshold and g >= white_threshold and b >= white_threshold

    visited = [[False] * h for _ in range(w)]
    q = deque()

    def try_add(x, y):
        if 0 <= x < w and 0 <= y < h and not visited[x][y]:
            if is_light_pixel(pixels[x, y]):
                visited[x][y] = True
                q.append((x, y))

    for x in range(w):
        try_add(x, 0)
        try_add(x, h - 1)
    for y in range(h):
        try_add(0, y)
        try_add(w - 1, y)

    dirs = [(1, 0), (-1, 0), (0, 1), (0, -1)]

    while q:
        x, y = q.popleft()
        for dx, dy in dirs:
            nx, ny = x + dx, y + dy
            if 0 <= nx < w and 0 <= ny < h and not visited[nx][ny]:
                if is_light_pixel(pixels[nx, ny]):
                    visited[nx][ny] = True
                    q.append((nx, ny))

    for x in range(w):
        for y in range(h):
            if visited[x][y]:
                r, g, b, a = pixels[x, y]
                pixels[x, y] = (r, g, b, 0)

    img = remove_white_fringe(img)
    return img


def normalize_url(url: str) -> str:
    url = (url or "").strip()
    if not url:
        return ""
    return quote(url, safe=':/?=&%.-_~()')


def extract_markdown_parts(value: str):
    value = (value or "").strip()
    m = re.match(r'^\[(.*?)\]\((.*?)\)(.*)$', value)
    if not m:
        return None
    return m.group(1).strip(), m.group(2).strip(), m.group(3).strip()


def extract_logo_url(value: str) -> str:
    value = (value or "").strip()
    if not value:
        return ""

    parts = extract_markdown_parts(value)
    if parts:
        text_part, link_part, tail_part = parts

        if tail_part:
            base = text_part.rstrip("/")
            full = f"{base} {tail_part}".strip()
            return normalize_url(full)

        return normalize_url(link_part)

    return normalize_url(value)


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


def prepare_logo_file(url, prefix="logo_processed"):
    if not url:
        return None
    try:
        if "-min" in url:
            url = url.replace("-min", "-max")

        r = requests.get(url, timeout=15)
        r.raise_for_status()

        img = Image.open(BytesIO(r.content)).convert("RGBA")
        if needs_background_removal(img):
            img = process_floodfill(img)

        temp_dir = os.path.join(tempfile.gettempdir(), "vmix_logo_temp")
        os.makedirs(temp_dir, exist_ok=True)
        out_path = os.path.join(temp_dir, f"{prefix}.png")
        img.save(out_path, "PNG")
        return out_path
    except Exception as e:
        print(f"Ошибка обработки логотипа {url}: {e}")
        return None


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
    home_logo = extract_logo_url(fields.get("Лого хозяев", ""))
    away_logo = extract_logo_url(fields.get("Лого гостей", ""))
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
        "home_logo": home_logo,
        "away_logo": away_logo,
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


def send_to_all_vmix_inputs(api_url, match, home_logo_file, away_logo_file):
    for input_name, fields in VMIX_INPUTS.items():

        # Отправка аббревиатур (названий) команд
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

        # Обнуление счета
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

        # Отправка логотипов
        if home_logo_file and "home_logo" in fields:
            vmix_send(api_url, {
                "Function": "SetImage",
                "Input": input_name,
                "SelectedName": fields["home_logo"],
                "Value": home_logo_file
            })

        if away_logo_file and "away_logo" in fields:
            vmix_send(api_url, {
                "Function": "SetImage",
                "Input": input_name,
                "SelectedName": fields["away_logo"],
                "Value": away_logo_file
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
        self.root.title("vMix Управление Трансляциями")
        self.root.geometry("980x650")
        self.root.minsize(920, 620)

        self.matches = []
        self.radio_widgets = []
        self.selected_match_idx = ctk.IntVar(value=-1)

        self.home_processed_ctk = None
        self.away_processed_ctk = None
        self.log_queue = queue.Queue()
        self.team_db = load_team_db()
        self.current_match_link = ""

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
        self.host_entry.bind("<Control-C>", lambda e: e.widget.event_generate("<<Copy>>"))
        self.host_entry.bind("<Control-x>", lambda e: e.widget.event_generate("<<Cut>>"))
        self.host_entry.bind("<Control-X>", lambda e: e.widget.event_generate("<<Cut>>"))
        self.host_entry.bind("<Control-v>", lambda e: e.widget.event_generate("<<Paste>>"))
        self.host_entry.bind("<Control-V>", lambda e: e.widget.event_generate("<<Paste>>"))

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
            text="Обновить",
            command=self.load_matches,
            fg_color="#555555",
            hover_color="#333333",
            font=ctk.CTkFont(size=13),
            width=145
        ).pack(side="left")

        ctk.CTkButton(
            actions_global,
            text="Pre-flight Check",
            command=self.run_preflight_check,
            fg_color="#d35400",
            hover_color="#e67e22",
            font=ctk.CTkFont(size=13, weight="bold"),
            width=155
        ).pack(side="right")

        # === НИЖНИЙ РЯД (Локальные действия с выбранным матчем) ===
        actions_local = ctk.CTkFrame(left, fg_color="transparent")
        actions_local.pack(fill="x", padx=14, pady=(0, 10))

        ctk.CTkButton(
            actions_local,
            text="Предпросмотр",
            command=self.preview_logos,
            fg_color="#1f6aa5",
            hover_color="#144870",
            font=ctk.CTkFont(size=13),
            width=145
        ).pack(side="left")

        ctk.CTkButton(
            actions_local,
            text="Отправить в vMix",
            command=self.send_to_vmix,
            fg_color="#28a745",
            hover_color="#218838",
            font=ctk.CTkFont(size=13, weight="bold"),
            width=155
        ).pack(side="right")

        self.scroll_frame = ctk.CTkScrollableFrame(left, width=310, height=480)
        self.scroll_frame.pack(fill="both", expand=True, padx=14, pady=(0, 14))

        ctk.CTkLabel(
            right,
            text="Предпросмотр логотипов",
            font=ctk.CTkFont(size=17, weight="bold")
        ).pack(anchor="w", padx=16, pady=(14, 4))

        self.preview_info = ctk.CTkLabel(
            right,
            text="Выбери матч и нажми «Проверить лого»",
            text_color="gray",
            font=ctk.CTkFont(size=13)
        )
        self.preview_info.pack(anchor="w", padx=16, pady=(0, 6))

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

        preview_grid = ctk.CTkFrame(right, fg_color="transparent")
        preview_grid.pack(fill="both", expand=True, padx=16, pady=(0, 10))

        self.home_frame = ctk.CTkFrame(preview_grid)
        self.home_frame.pack(side="left", fill="both", expand=True, padx=(0, 8))

        self.away_frame = ctk.CTkFrame(preview_grid)
        self.away_frame.pack(side="left", fill="both", expand=True, padx=(8, 0))

        self.build_preview_column(self.home_frame, "Хозяева", "home")
        self.build_preview_column(self.away_frame, "Гости", "away")

        ctk.CTkLabel(
            right,
            text="Лог",
            font=ctk.CTkFont(size=15, weight="bold")
        ).pack(anchor="w", padx=16, pady=(4, 4))

        self.log_box = ctk.CTkTextbox(right, height=150)
        self.log_box.pack(fill="x", padx=16, pady=(0, 12))
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

        # Привязка локальных горячих клавиш
        self.root.bind("<Up>", lambda e: self.navigate_matches(-1))
        self.root.bind("<Down>", lambda e: self.navigate_matches(1))
        self.root.bind("<Return>", lambda e: self.send_to_vmix())  # Клавиша Enter
        self.root.bind("<space>", lambda e: self.preview_logos())  # Клавиша Пробел

        # Загрузка матчей при старте
        self.load_matches()

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

    def build_preview_column(self, parent, title, side):
        ctk.CTkLabel(
            parent,
            text=title,
            font=ctk.CTkFont(size=15, weight="bold")
        ).pack(pady=(12, 8))

        processed_label = ctk.CTkLabel(parent, text="")
        processed_label.pack(pady=(6, 10))

        # Блок выбора цвета
        color_frame = ctk.CTkFrame(parent, fg_color="transparent")
        color_frame.pack(pady=(0, 10))

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

        # По умолчанию скрываем выбор цвета, пока не загрузим предпросмотр
        color_menu.configure(state="disabled")

        if side == "home":
            self.home_processed_label = processed_label
            self.home_color_var = color_var
            self.home_color_swatch = color_swatch
            self.home_color_menu = color_menu
        else:
            self.away_processed_label = processed_label
            self.away_color_var = color_var
            self.away_color_swatch = color_swatch
            self.away_color_menu = color_menu

    def load_ctk_image(self, file_path, size=(170, 170), checkerboard=True):
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
            final_img = img

        return ctk.CTkImage(light_image=final_img, dark_image=final_img, size=size)

    def navigate_matches(self, step):
        """Перемещение по списку матчей с помощью стрелок Вверх/Вниз"""
        if not self.matches:
            return

        current_idx = self.selected_match_idx.get()

        # Если ни один матч еще не выбран, стартуем с первого
        if current_idx == -1:
            next_idx = 0
        else:
            # Вычисляем следующий индекс
            next_idx = current_idx + step

            # Защита от выхода за пределы списка (чтобы программа не выдала ошибку)
            if next_idx < 0:
                next_idx = 0
            elif next_idx >= len(self.matches):
                next_idx = len(self.matches) - 1

        # Устанавливаем новый выбор
        self.selected_match_idx.set(next_idx)

        # Очищаем предпросмотр логотипов от предыдущего матча
        self.clear_previews()

        # Обновляем статус-бар
        match = self.matches[next_idx]
        self.status.configure(text=f"Выбран матч: {match['team1']} vs {match['team2']}", text_color="gray")

    def clear_previews(self):
        self.home_processed_ctk = None
        self.away_processed_ctk = None
        self.home_processed_label.configure(image="", text="")
        self.away_processed_label.configure(image="", text="")
        self.home_color_menu.configure(state="disabled")
        self.away_color_menu.configure(state="disabled")
        self.current_match_link = ""
        self.match_link_button.configure(
            state="disabled",
            text="Открыть матч в браузере"
        )

    def apply_preview_results(self, match, home_processed, away_processed):
        self.preview_info.configure(
            text=f"Проверка логотипов: {match['team1']} vs {match['team2']}",
            text_color="#cccccc"
        )

        home_color, _ = get_color_row(match["team1"], self.team_db)
        away_color, _ = get_color_row(match["team2"], self.team_db)

        self.home_color_var.set(home_color)
        self.home_color_swatch.configure(fg_color=UI_COLORS.get(home_color, "#FFFFFF"))
        self.home_color_menu.configure(state="normal")

        self.away_color_var.set(away_color)
        self.away_color_swatch.configure(fg_color=UI_COLORS.get(away_color, "#FFFFFF"))
        self.away_color_menu.configure(state="normal")

        if home_processed:
            self.home_processed_ctk = self.load_ctk_image(home_processed, checkerboard=True)
            self.home_processed_label.configure(image=self.home_processed_ctk, text="")
        else:
            self.home_processed_label.configure(text="Не удалось обработать", image="")

        if away_processed:
            self.away_processed_ctk = self.load_ctk_image(away_processed, checkerboard=True)
            self.away_processed_label.configure(image=self.away_processed_ctk, text="")
        else:
            self.away_processed_label.configure(text="Не удалось обработать", image="")

        self.current_match_link = match.get("match_link", "")

        if self.current_match_link:
            self.match_link_button.configure(
                state="normal",
                text="Открыть матч в Rutube Studio"
            )
        else:
            self.match_link_button.configure(
                state="disabled",
                text="Ссылка на матч недоступна"
            )

        self.status.configure(text="Предпросмотр логотипов готов", text_color="#28a745")

    def open_current_match_link(self):
        if self.current_match_link:
            webbrowser.open_new_tab(self.current_match_link)

    def preview_logos(self):
        idx = self.selected_match_idx.get()
        if idx == -1:
            messagebox.showwarning("Внимание", "Сначала выберите матч из списка!")
            return

        self.clear_previews()
        self.status.configure(text="Скачивание и обработка логотипов...", text_color="#3399ff")
        self.add_log("Старт обработки логотипов...")
        threading.Thread(target=self.preview_logos_worker, args=(idx,), daemon=True).start()

    def preview_logos_worker(self, idx):
        try:
            match = self.matches[idx]
            self.add_log(f"Скачиваю лого хозяев: {match['team1']}")
            home_processed = prepare_logo_file(match["home_logo"], "home_processed")

            self.add_log(f"Скачиваю лого гостей: {match['team2']}")
            away_processed = prepare_logo_file(match["away_logo"], "away_processed")

            self.root.after(0, lambda: self.apply_preview_results(match, home_processed, away_processed))
            self.add_log("Готово.")
        except Exception as e:
            self.add_log(f"Ошибка: {e}")
            self.root.after(0, lambda: messagebox.showerror("Ошибка", str(e)))
            self.root.after(0, lambda: self.status.configure(text="Ошибка предпросмотра", text_color="#ff4444"))

    def get_current_api(self):
        return get_api_url(self.host_entry.get())

    def ping_vmix(self):
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
        self.team_db = load_team_db()

        for widget in self.radio_widgets:
            widget.destroy()
        self.radio_widgets.clear()
        self.selected_match_idx.set(-1)

        if hasattr(self, "home_processed_label"):
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
                font=ctk.CTkFont(size=14),
                hover_color="#28a745"
            )
            rb.pack(anchor="w", pady=6, padx=12)
            self.radio_widgets.append(rb)

        if self.matches:
            self.selected_match_idx.set(0)

        self.status.configure(text=f"Загружено матчей: {len(self.matches)}", text_color="gray")
        self.add_log(f"Загружено матчей: {len(self.matches)}")

    def run_preflight_check(self):
        if not self.matches:
            messagebox.showwarning("Внимание", "Сначала загрузите матчи из файла stream_keys.txt!")
            return

        self.status.configure(text="Выполняется Pre-flight Check...", text_color="#d35400")
        self.add_log("Запущена проверка всех ссылок и ключей. Пожалуйста, подождите...")

        # Запускаем в отдельном потоке, чтобы интерфейс не завис
        threading.Thread(target=self.preflight_worker, daemon=True).start()

    def check_url_fast(self, url):
        """Быстрая проверка доступности ссылки без скачивания всей картинки"""
        if not url:
            return False, "Нет ссылки"
        if "-min" in url:
            url = url.replace("-min", "-max")
        try:
            # Используем stream=True, чтобы скачать только заголовки, а не сам файл
            r = requests.get(url, stream=True, timeout=5)
            if r.status_code == 200:
                return True, "OK"
            else:
                return False, f"Ошибка {r.status_code}"
        except requests.exceptions.Timeout:
            return False, "Сервер не ответил (Timeout)"
        except Exception:
            return False, "Недоступен (Сбой соединения)"

    def preflight_worker(self):
        errors_count = 0

        self.add_log("\n=== СТАРТ PRE-FLIGHT CHECK ===")

        for i, match in enumerate(self.matches):
            match_name = f"{match['team1']} vs {match['team2']}"
            match_issues = []

            # 1. Проверка логотипов
            home_ok, home_msg = self.check_url_fast(match['home_logo'])
            if not home_ok:
                match_issues.append(f"Лого хозяев: {home_msg} ({match['home_logo']})")

            away_ok, away_msg = self.check_url_fast(match['away_logo'])
            if not away_ok:
                match_issues.append(f"Лого гостей: {away_msg} ({match['away_logo']})")

            # 2. Проверка ключа трансляции
            if not match['key'] or not match['server']:
                match_issues.append("Отсутствует RTMP сервер или ключ трансляции")

            # Если есть проблемы, выводим их в лог
            if match_issues:
                self.add_log(f"[Матч {i + 1}] {match_name}:")
                for issue in match_issues:
                    self.add_log(f"   {issue}")
                errors_count += len(match_issues)

        # Вывод итогов проверки
        if errors_count == 0:
            self.add_log("Все матчи готовы к эфиру! Проблем не найдено.")
            self.root.after(0, lambda: self.status.configure(text="Pre-flight: ОК", text_color="#28a745"))
        else:
            self.add_log(f"Проверка завершена. Найдено проблем: {errors_count}")
            self.root.after(0, lambda: self.status.configure(text=f"Pre-flight: Найдено проблем - {errors_count}",
                                                             text_color="#ffcc00"))

        self.add_log("=== КОНЕЦ ПРОВЕРКИ ===\n")

    def send_to_vmix(self):
        idx = self.selected_match_idx.get()
        if idx == -1:
            messagebox.showwarning("Внимание", "Сначала выберите матч из списка!")
            return

        match = self.matches[idx]
        api_url = self.get_current_api()

        self.status.configure(
            text=f"Отправка во все инпуты → {self.host_entry.get()}...",
            text_color="#3399ff"
        )
        self.add_log(f"Старт отправки в vMix: {match['team1']} vs {match['team2']}")
        threading.Thread(target=self.send_to_vmix_worker, args=(match, api_url), daemon=True).start()

    def send_to_vmix_worker(self, match, api_url):
        try:
            home_color, home_row = get_color_row(match["team1"], self.team_db)
            away_color, away_row = get_color_row(match["team2"], self.team_db)

            self.add_log(f"Цвет хозяев: {home_color}, строка {home_row}")
            self.add_log(f"Цвет гостей: {away_color}, строка {away_row}")

            self.add_log("Обрабатываю логотип хозяев...")
            home_logo_file = prepare_logo_file(match["home_logo"], "home_logo")

            self.add_log("Обрабатываю логотип гостей...")
            away_logo_file = prepare_logo_file(match["away_logo"], "away_logo")

            self.add_log("Отправляю данные во все инпуты vMix...")
            send_to_all_vmix_inputs(api_url, match, home_logo_file, away_logo_file)

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
                    text=f"Успешно: {match['team1']} vs {match['team2']} ({home_color}/{away_color})",
                    text_color="#28a745"
                )
            )
            self.root.after(0, lambda: save_last_host(self.host_entry.get()))
            self.add_log("Отправка завершена успешно.")

            current_idx = self.selected_match_idx.get()
            if current_idx + 1 < len(self.matches):
                next_idx = current_idx + 1
                # Через after безопасно переключаем UI из фонового потока
                self.root.after(0, lambda: self.selected_match_idx.set(next_idx))
                # Очищаем предпросмотр, чтобы старые логотипы не сбивали с толку
                self.root.after(0, self.clear_previews)
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
