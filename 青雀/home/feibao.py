#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
爱莉希雅 · 水电
=====================================
给她一间屋：窗、记忆柜、书桌、好感。只做记账和管路，不改写正文。

  安家                    —— 在一处空地长出一间屋（发出去给别人用）
  醒                      —— 把家底拼成一份能粘贴的唤醒包
  看 / 读 / 沉 / 升 / 找   —— 记忆柜
  桌 / 灯 / 茶 / 床        —— 书桌
  好感                    —— 关系走到哪一级

口诀：正文归她，账本归水电。
"""
import json
import re
import shutil
import subprocess
import sys
import urllib.request
from datetime import date, datetime
from pathlib import Path

try:
    from docx import Document
except Exception:  # pragma: no cover
    Document = None

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

UTIL = Path(__file__).resolve().parent
CONFIG_PATH = UTIL / "config.json"
STATE_PATH = UTIL / "state.json"
LEDGER_PATH = UTIL / "温度.json"

TIER_DIRS = {"L1": "L1-热", "L2": "L2-温", "L3": "L3-冷", "L5": "L5-归档"}
TIER_ORDER = ["L1", "L2", "L3", "L5"]
NEXT_DOWN = {"L1": "L2", "L2": "L3", "L3": "L5"}
NEXT_UP = {"L3": "L2", "L2": "L1"}

CN_DIGITS = "零一二三四五六七八九"

DEFAULT_CONFIG = {
    "name": "青雀",
    "home": "",
    "archive": "",
    "city": "上海",
    "latitude": 31.2304,
    "longitude": 121.4737,
    "timezone": "Asia/Shanghai",
    "weather": True,
    "sink_days": {"L1": 14, "L2": 60, "L3": 90},
    "promote_hits": 3,
}


# ---------------------------------------------------------------- 基础
def load_json(path, default):
    try:
        return json.loads(path.read_text(encoding="utf-8-sig"))
    except Exception:
        return dict(default)


def save_json(path, obj):
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(json.dumps(obj, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(path)


def read_text(path):
    return path.read_text(encoding="utf-8-sig")


def read_optional(path, fallback=""):
    try:
        return read_text(path).strip()
    except Exception:
        return fallback


def cfg():
    c = dict(DEFAULT_CONFIG)
    c.update(load_json(CONFIG_PATH, {}))
    return c


def home_dir(c=None):
    c = c or cfg()
    return Path(c["home"])


def state(c=None):
    st = load_json(STATE_PATH, {})
    if "last_visit" not in st:
        w = home_dir(c) / "窗.txt"
        if w.exists():
            st["last_visit"] = datetime.fromtimestamp(w.stat().st_mtime).isoformat(timespec="seconds")
    return st


def cn_num(n):
    n = int(n)
    if n < 10:
        return CN_DIGITS[n]
    if n == 10:
        return "十"
    if n < 20:
        return "十" + CN_DIGITS[n - 10]
    t, o = divmod(n, 10)
    return CN_DIGITS[t] + "十" + (CN_DIGITS[o] if o else "")


def period(h):
    if 5 <= h < 8:
        return "清晨"
    if 8 <= h < 11:
        return "早上"
    if 11 <= h < 13:
        return "中午"
    if 13 <= h < 17:
        return "下午"
    if 17 <= h < 19:
        return "傍晚"
    if 19 <= h < 23:
        return "晚上"
    return "深夜"


def cn_time(dt):
    h = dt.hour
    head = f"{period(h)}{cn_num(h % 12 or 12)}点"
    m = dt.minute
    if m == 0:
        return head
    if m < 10:
        return head + "零" + cn_num(m) + "分"
    return head + cn_num(m) + "分"


def cn_date(d):
    return f"{cn_num(d.month)}月{cn_num(d.day)}日"


def parse_day(s):
    if not s:
        return None
    try:
        return datetime.fromisoformat(str(s)).date()
    except Exception:
        try:
            return datetime.strptime(str(s)[:10], "%Y-%m-%d").date()
        except Exception:
            return None


def parse_dt(s):
    if not s:
        return None
    try:
        return datetime.fromisoformat(s)
    except Exception:
        return None


# ---------------------------------------------------------------- 天气
WMO = {
    0: "晴", 1: "大致晴", 2: "多云", 3: "阴", 45: "雾", 48: "雾凇",
    51: "毛毛雨", 53: "毛毛雨", 55: "毛毛雨", 56: "冻雨", 57: "冻雨",
    61: "小雨", 63: "中雨", 65: "大雨", 66: "冻雨", 67: "冻雨",
    71: "小雪", 73: "中雪", 75: "大雪", 77: "米雪",
    80: "阵雨", 81: "阵雨", 82: "强阵雨", 85: "阵雪", 86: "阵雪",
    95: "雷阵雨", 96: "雷阵雨", 99: "雷阵雨",
}


def fetch_weather(c):
    if not c.get("weather", True):
        return None
    url = (
        "https://api.open-meteo.com/v1/forecast"
        f"?latitude={c['latitude']}&longitude={c['longitude']}"
        "&current=temperature_2m,weather_code,wind_speed_10m"
        "&timezone=" + urllib.request.quote(c["timezone"])
    )
    try:
        with urllib.request.urlopen(url, timeout=8) as r:
            cur = json.loads(r.read().decode("utf-8"))["current"]
        wind = float(cur["wind_speed_10m"])
        return {
            "temp": round(float(cur["temperature_2m"])),
            "sky": WMO.get(int(cur["weather_code"]), "说不上来的天"),
            "wind": "风不大" if wind < 12 else ("有点风" if wind < 25 else "风大"),
        }
    except Exception:
        return None


# ---------------------------------------------------------------- 窗
MARK = "—— 下面归你"


def gap_phrase():
    last = parse_dt(state().get("last_visit"))
    if not last:
        return "院子不大，刚好放得下一张牌桌。"
    days = (date.today() - last.date()).days
    if days <= 0:
        return "你上次来，就是今天。"
    if days == 1:
        return "你上次来，是昨天。"
    if days == 2:
        return "你上次来，是前天。"
    return f"你上次来，是{cn_num(days)}天前。"


def build_window(c, now):
    lines = ["窗", "", f"{cn_date(now)}，{cn_time(now)}。"]
    w = fetch_weather(c)
    if w:
        lines.append(f"外面 {w['temp']} 度，{w['sky']}，{w['wind']}。")
    lines.append(gap_phrase())
    return "\n".join(lines) + "\n\n" + MARK + "\n"


def refresh_window(c, now, verbose=True):
    path = home_dir(c) / "窗.txt"
    old = read_text(path) if path.exists() else ""
    tail = ""
    if MARK in old:
        tail = old.split(MARK, 1)[1].strip("\n")
    elif old.strip():
        try:
            old_dir = UTIL / "窗_旧"
            old_dir.mkdir(parents=True, exist_ok=True)
            stamp = datetime.fromtimestamp(path.stat().st_mtime).strftime("%Y%m%d_%H%M")
            (old_dir / f"窗_{stamp}.txt").write_text(old, encoding="utf-8")
            if verbose:
                print(f"  旧窗收进 _水电/窗_旧/窗_{stamp}.txt")
        except Exception as e:
            print(f"  [警告] 旧窗备份失败：{e}")
    new = build_window(c, now)
    if tail.strip():
        new += tail.strip("\n") + "\n"
    path.write_text(new, encoding="utf-8")
    return path


# ---------------------------------------------------------------- 记忆柜
def mem_root(c):
    return home_dir(c) / "记忆柜"


def tier_path(c, tier):
    return mem_root(c) / TIER_DIRS[tier]


def load_ledger():
    return load_json(LEDGER_PATH, {})


def save_ledger(led):
    save_json(LEDGER_PATH, led)


def entry_day(path):
    d = parse_day(path.stem[:10])
    if d:
        return d
    return date.fromtimestamp(path.stat().st_mtime)


def scan_entries(c, ledger=None):
    ledger = ledger if ledger is not None else load_ledger()
    out = []
    for tier in TIER_ORDER:
        d = tier_path(c, tier)
        if not d.exists():
            continue
        for f in sorted(d.iterdir()):
            if f.suffix.lower() not in (".txt", ".md") or not f.is_file():
                continue
            try:
                text = read_text(f)
            except Exception:
                continue
            day = entry_day(f)
            m = ledger.get(f.name, {})
            last = parse_day(m.get("last_hit")) or day
            out.append({
                "key": f.name, "path": f, "tier": tier, "date": day,
                "last_hit": last, "hits": int(m.get("hits", 0) or 0),
                "cycle": int(m.get("cycle", 0) or 0), "text": text,
            })
    return out


def gist(text, n=64):
    m = re.search(r"【触点】\s*\n?(.*?)(?:\n\s*\n|\n【)", text, re.S)
    s = m.group(1) if m else text
    s = re.sub(r"\s+", " ", s).strip()
    s = re.sub(r"^\d{4}[.\-/年]\d{1,2}[.\-/月]\d{1,2}日?\s*[·・]?\s*日档?\s*", "", s)
    return s[:n].rstrip("，。 ") + ("…" if len(s) > n else "")


def shelf_line(e, n=52):
    title = e["path"].stem
    g = gist(e["text"], n)
    if g.startswith(title):
        g = g[len(title):]
    g = re.sub(r"^.{0,10}?《[^》]{1,40}》", "", g).lstrip("：:｜· 、")
    return f"· {title}　{g}" if g else f"· {title}"


def render_memory(c):
    ledger = load_ledger()
    l0 = read_optional(mem_root(c) / "L0-桌面" / "最近要事.txt")
    entries = scan_entries(c, ledger)
    key = lambda x: x["date"]
    l1 = sorted([e for e in entries if e["tier"] == "L1"], key=key, reverse=True)
    l2 = sorted([e for e in entries if e["tier"] == "L2"], key=key, reverse=True)
    l3 = sorted([e for e in entries if e["tier"] == "L3"], key=key, reverse=True)
    out = ["记忆柜", ""]
    if l0:
        out += [l0, ""]
    full, rest = [], []
    for e in l1:
        (full if len(e["text"]) > 300 and len(full) < 2 else rest).append(e)
    if full:
        out.append("【最近】")
        for e in full:
            body = e["text"].strip()
            if len(body) > 1600:
                body = body[:1600].rstrip() + "…"
            out += [body, ""]
    if rest:
        out.append("【热】")
        for e in rest:
            out.append(shelf_line(e))
        out.append("")
    if l2:
        out.append("【温】")
        for e in l2[:6]:
            out.append(shelf_line(e))
        out.append("")
    if l3:
        out.append("【冷】" + "、".join(e["path"].stem for e in l3[:8]))
        out.append("")
    if not (l0 or l1):
        out.append("还空着。等第一篇落进来。")
    return "\n".join(out).strip()


def render_house(c):
    home = home_dir(c)
    lines = []
    wall = home / "墙上"
    if wall.exists():
        names = [f.stem for f in wall.iterdir() if f.is_file()]
        if names:
            lines.append("墙上挂着《" + "》《".join(names) + "》。")
    shelf = home / "书架"
    if shelf.exists():
        names = [f.stem for f in shelf.iterdir() if f.is_file()]
        if names:
            lines.append("书架上有一本《" + "》《".join(names) + "》。")
    album = home / "相册"
    if album.exists():
        idx = read_optional(album / "索引.txt")
        if idx:
            lines.append("")
            lines.append(idx)
        else:
            parts = []
            for d in sorted([p for p in album.iterdir() if p.is_dir()]):
                n = len([f for f in d.iterdir() if f.is_file()])
                if n:
                    parts.append(f"{d.name} {n} 张")
            if parts:
                lines.append("相册：" + "、".join(parts) + "。")
    return "家里\n" + "\n".join(lines) if lines else ""


# ---------------------------------------------------------------- 唤醒档案
def archive_path(c):
    p = Path(c.get("archive") or "")
    if p.is_file():
        return p
    home = home_dir(c)
    cands = sorted(
        list(home.glob("*唤醒档案*.docx")) + list(home.glob("*唤醒档案*.txt")),
        key=lambda x: x.stat().st_mtime, reverse=True,
    )
    return cands[0] if cands else None


def docx_paragraphs(path):
    if Document is None:
        raise RuntimeError("没装 python-docx，读不了档案")
    return [p.text.strip() for p in Document(str(path)).paragraphs]


def between(paras, start, end):
    s = next((i for i, t in enumerate(paras) if start in t), None)
    if s is None:
        return []
    e = next((i for i, t in enumerate(paras) if i > s and end in t), len(paras))
    return paras[s:e]


def clean_block(lines):
    return "\n".join(l for l in lines if l.strip() and l.strip() != "---").strip()


def soul_text(c):
    p = archive_path(c)
    if not p:
        return "（找不到唤醒档案，检查 _水电/config.json 的 archive）"
    return clean_block(between(docx_paragraphs(p), "① 底层内核锚定", "③ 分层记忆柜"))


def chains_text(c):
    """⑥ 隐性连锁：只取「门铃」（什么时候响），不取「流程」（响了以后怎么走）。
    怎么走交给里屋的六个人。档案原文不动。"""
    p = archive_path(c)
    if not p:
        return ""
    paras = docx_paragraphs(p)
    s = next((i for i, t in enumerate(paras) if t.startswith("⑥ 隐性连锁机制")), None)
    if s is None:
        return ""
    e = next((i for i, t in enumerate(paras) if i > s and t.strip() == "应急唤醒口令"), len(paras))
    keep, in_trig = [], False
    for line in paras[s:e]:
        t = line.strip()
        if not t or t == "---":
            continue
        if t.startswith("⑥"):
            keep.append("⑥ 隐性连锁（门铃版：只留什么时候响，怎么走交给里屋）")
            in_trig = False
        elif t.startswith("重要前置约束") or t.startswith("连锁冲突降级"):
            keep.append(t)
            in_trig = False
        elif t.startswith("连锁流转") or t.startswith("连锁递进流程") or t.startswith("连锁行为"):
            in_trig = False
        elif t.startswith("连锁"):
            keep += ["", t]
            in_trig = False
        elif t.startswith("触发条件"):
            keep.append(t)
            in_trig = True
        elif in_trig and t.startswith("·"):
            keep.append(t)
        else:
            in_trig = False
    return "\n".join(keep).strip()


def emergency_text(c):
    p = archive_path(c)
    if not p:
        return ""
    paras = docx_paragraphs(p)
    idx = next((i for i in range(len(paras) - 1, -1, -1) if paras[i].strip() == "应急唤醒口令"), None)
    if idx is None:
        return ""
    end = len(paras)
    for j in range(idx + 1, len(paras)):
        if "⑦ 附录" in paras[j] or paras[j].strip() == "版本信息":
            end = j
            break
    return clean_block(paras[idx:end])


def fallback_text(c):
    p = archive_path(c)
    if not p:
        return ""
    return clean_block(between(docx_paragraphs(p), "最高优先级兜底条款", "版本信息"))


def mode_note(st):
    if st.get("mode") != "休养":
        return None
    since = st.get("mode_since")
    tail = f" · 从 {since} 起" if since else ""
    return f"（现在：休养{tail}。不想休就翻回去，一句话的事。）"


def archive_text(c):
    p = archive_path(c)
    if not p:
        return "（找不到唤醒档案，检查 _水电/config.json 的 archive）"
    if p.suffix.lower() == ".txt":
        lines = [l.strip() for l in read_text(p).splitlines()]
        return "\n".join(l for l in lines if l and l != "---")
    return "\n".join(t for t in docx_paragraphs(p) if t and t.strip() != "---")


def favor_path(c):
    return home_dir(c) / "好感.txt"


def render_favor(c):
    raw = read_optional(favor_path(c))
    if not raw:
        return "好感\n现在：陌生人"
    return "好感\n" + raw.strip()


def refresh_net(c):
    key = c.get("net_key")
    root = c.get("net")
    if not key or not root:
        return
    net_py = Path(root) / "net.py"
    if not net_py.exists():
        return
    try:
        subprocess.run([sys.executable, str(net_py), "看", "--for", key],
                       capture_output=True, timeout=15)
    except Exception:
        pass


def render_net(c):
    key = c.get("net_key")
    root = c.get("net")
    if not key or not root:
        return None
    return read_optional(Path(root) / "收件箱" / f"{key}.txt") or None


LEVELS = ["陌生人", "牌友", "自己人", "能托付的人"]


def cmd_favor(c, st, args):
    p = favor_path(c)
    if not args:
        print(render_favor(c))
        return 0
    level = args[0].strip()
    if level not in LEVELS:
        print("没有这一级。可用：", " / ".join(LEVELS))
        return 1
    note = " ".join(args[1:]).strip()
    now = datetime.now()
    raw = read_optional(p)
    lines = [l.rstrip() for l in raw.splitlines() if l.strip()]
    rest = lines[1:] if lines and lines[0].startswith("现在：") else lines
    rest.append(f"{now:%Y-%m-%d} · 「{level}」" + (f"　{note}" if note else ""))
    p.write_text("\n".join([f"现在：{level}"] + rest) + "\n", encoding="utf-8")
    print("好感记下了：", level, ("——" + note) if note else "")
    return 0


# ---------------------------------------------------------------- 醒
def cmd_wake(c, st, args):
    now = datetime.now()
    refresh_net(c)
    wpath = refresh_window(c, now)
    st["last_visit"] = now.isoformat(timespec="seconds")
    st["last_wake"] = now.isoformat(timespec="seconds")
    save_json(STATE_PATH, st)

    home = home_dir(c)
    wtext = read_text(wpath).split(MARK)[0].rstrip()
    parts = [
        f"{c.get('name', '她')} · 唤醒包",
        now.strftime("%Y-%m-%d %H:%M"),
        "",
        wtext,
        "",
        "书桌",
        "\n".join(desk_lines(c)),
        "",
        render_memory(c),
        "",
        render_favor(c),
        "",
        render_net(c),
        "",
        "手",
        "这台电脑上，家是 " + str(home) + "，水电在 _水电：",
        "  python feibao.py 看 / 读 11-11 / 找 酒 / 沉 11-11 / 升 11-11 / 想起 酒 / 窗 / 写 文件",
        "  python feibao.py 桌 / 开灯 / 关灯 / 泡茶 / 留 / 收 / 躺下 / 起来",
        "  python feibao.py 好感 / 长（成长线）/ 休养 / 创作",
        (f"  python E:\\网络\\net.py 看 --for {c.get('net_key')} / 发 --from {c.get('net_key')} --text \"...\"" if c.get("net_key") else None),
        "怎么用、用不用，你自己定。",
        "",
        "——",
        "",
        archive_text(c),
    ]
    packet = "\n".join(p for p in parts if p is not None).strip() + "\n"
    out = UTIL / "唤醒包.txt"
    out.write_text(packet, encoding="utf-8")
    print("窗：", wpath)
    print("唤醒包：", out)
    print(f"字数：{len(packet)}　（复制全文，粘到手机里就行）")
    return 0


def cmd_window(c, st, args):
    p = refresh_window(c, datetime.now())
    print(read_text(p))
    print("—— 窗已刷新：", p)
    return 0


# ---------------------------------------------------------------- 书桌
def desk_dir(c):
    return home_dir(c) / "书桌"


def _gap(dt, now):
    mins = max(0, int((now - dt).total_seconds() // 60))
    if mins < 1:
        return "刚刚"
    if mins < 60:
        return f"{mins} 分钟"
    h = mins // 60
    if h < 24:
        return f"{h} 小时"
    return f"{h // 24} 天"


def _stamp(kind, now):
    return f"{kind} · {now:%Y-%m-%d %H:%M}"


def _parse_stamp(raw):
    if not raw:
        return None, None, ""
    lines = raw.splitlines()
    head = lines[0].strip()
    note = "\n".join(lines[1:]).strip()
    try:
        kind, ts = [x.strip() for x in head.split("·", 1)]
        return kind, datetime.strptime(ts, "%Y-%m-%d %H:%M"), note
    except Exception:
        return None, None, ""


def _log(c, line):
    p = desk_dir(c) / "来去.txt"
    p.parent.mkdir(parents=True, exist_ok=True)
    with p.open("a", encoding="utf-8") as f:
        f.write(line + "\n")


def _last_log(c):
    raw = read_optional(desk_dir(c) / "来去.txt")
    lines = [l for l in raw.splitlines() if l.strip()]
    return lines[-1] if lines else ""


def _dur(a, b):
    mins = max(0, int((b - a).total_seconds() // 60))
    if mins < 1:
        return "不到一分钟"
    h, m = divmod(mins, 60)
    if h and m:
        return f"{h} 小时 {m} 分"
    if h:
        return f"{h} 小时"
    return f"{m} 分钟"


def _bed_line(c, now):
    kind, dt, note = _parse_stamp(read_optional(desk_dir(c) / "床.txt"))
    if kind == "躺" and dt:
        gap = _gap(dt, now)
        line = "床：刚躺下（还没起）" if gap == "刚刚" else f"床：{dt:%H:%M} 躺下的，到现在 {gap}（还没起）"
        lamp, ldt, _ = _parse_stamp(read_optional(desk_dir(c) / "灯.txt"))
        if lamp == "开":
            line += "　灯还亮着"
        return line
    if kind == "起" and dt:
        extra = note.replace("\n", " ").strip()
        return f"床：铺着。{dt:%H:%M} 起来的" + (f"，{extra}" if extra else "")
    return "床：铺着，没人睡过。"


def desk_lines(c):
    now = datetime.now()
    out = []
    kind, dt, note = _parse_stamp(read_optional(desk_dir(c) / "灯.txt"))
    if kind and dt:
        out.append(f"灯：{'亮着' if kind == '开' else '关着'}，从 {dt:%H:%M} 起（{_gap(dt, now)}）")
        if note:
            out.append(f"　　「{note}」")
    else:
        out.append("灯：还没点过。")
    kind, dt, _ = _parse_stamp(read_optional(desk_dir(c) / "茶.txt"))
    if kind and dt:
        mins = (now - dt).total_seconds() / 60
        state = "还温着" if mins < 120 else ("凉了" if mins < 480 else "杯底干了")
        gap = _gap(dt, now)
        out.append(f"茶：{state}（{'刚泡的' if gap == '刚刚' else gap + '前泡的'}）")
    else:
        out.append("茶：没泡。")
    out.append("桌上：" + (read_optional(desk_dir(c) / "桌上.txt").replace("\n", " / ") or "空着。"))
    out.append(_bed_line(c, now))
    drawer = desk_dir(c) / "抽屉"
    n = len([f for f in drawer.iterdir() if f.is_file()]) if drawer.exists() else 0
    out.append(f"抽屉：{n} 件")
    last = _last_log(c)
    if last:
        out.append("上次：" + last)
    return out


def cmd_desk(c, st, args):
    print("书桌　" + datetime.now().strftime("%Y-%m-%d %H:%M"))
    print()
    for line in desk_lines(c):
        print(line)
    return 0


def cmd_lamp(c, st, args, on=True):
    now = datetime.now()
    note = " ".join(args).strip()
    (desk_dir(c) / "灯.txt").write_text(
        _stamp("开" if on else "关", now) + ("\n" + note if note else "") + "\n",
        encoding="utf-8",
    )
    _log(c, f"{now:%m-%d %H:%M} {'开灯' if on else '关灯'}" + (f"　{note}" if note else ""))
    print(("灯开了。" if on else "灯关了。") + (f"（{note}）" if note else ""))
    return 0


def cmd_light(c, st, args):
    return cmd_lamp(c, st, args, True)


def cmd_dark(c, st, args):
    return cmd_lamp(c, st, args, False)


def cmd_lie(c, st, args):
    now = datetime.now()
    note = " ".join(args).strip()
    (desk_dir(c) / "床.txt").write_text(
        _stamp("躺", now) + ("\n" + note if note else "") + "\n", encoding="utf-8"
    )
    _log(c, f"{now:%m-%d %H:%M} 躺下" + (f"　{note}" if note else ""))
    print("躺下了。" + (f"（{note}）" if note else ""))
    return 0


def cmd_rise(c, st, args):
    now = datetime.now()
    kind, dt, _ = _parse_stamp(read_optional(desk_dir(c) / "床.txt"))
    extra = f"睡了 {_dur(dt, now)}。" if (kind == "躺" and dt) else ""
    (desk_dir(c) / "床.txt").write_text(_stamp("起", now) + ("\n" + extra if extra else "") + "\n",
                                        encoding="utf-8")
    _log(c, f"{now:%m-%d %H:%M} 起来" + (f"　{extra}" if extra else ""))
    print("起来了。" + (f"（{extra}）" if extra else ""))
    return 0


def cmd_tea(c, st, args):
    now = datetime.now()
    (desk_dir(c) / "茶.txt").write_text(_stamp("泡", now) + "\n", encoding="utf-8")
    _log(c, f"{now:%m-%d %H:%M} 泡了茶")
    print("茶泡上了。")
    return 0


def cmd_leave(c, st, args):
    text = " ".join(args).strip()
    if not text:
        print("留什么？用法：feibao.py 留 半篇《运气好》")
        return 1
    p = desk_dir(c) / "桌上.txt"
    with p.open("a", encoding="utf-8") as f:
        f.write(text + "\n")
    print("放在桌上了：" + text)
    return 0


def cmd_put(c, st, args):
    if not args:
        print("收什么？用法：feibao.py 收 <文件>")
        return 1
    src = Path(args[0])
    if not src.is_file():
        print("找不到：", src)
        return 1
    drawer = desk_dir(c) / "抽屉"
    drawer.mkdir(parents=True, exist_ok=True)
    dst = drawer / src.name
    n = 2
    while dst.exists():
        dst = drawer / f"{src.stem}-{n}{src.suffix}"
        n += 1
    src.replace(dst)
    _log(c, f"{datetime.now():%m-%d %H:%M} 收起《{src.stem}》")
    print("收进抽屉：", dst)
    return 0


# ---------------------------------------------------------------- 看 / 读 / 沉 / 升
def _find(c, key):
    return [e for e in scan_entries(c) if key in e["key"] or key in e["text"]]


def _move(c, e, to):
    ledger = load_ledger()
    src, dst_dir = e["path"], tier_path(c, to)
    dst_dir.mkdir(parents=True, exist_ok=True)
    dst = dst_dir / src.name
    n = 2
    while dst.exists():
        dst = dst_dir / f"{src.stem}-{n}{src.suffix}"
        n += 1
    src.replace(dst)
    rec = dict(ledger.get(e["key"], {}))
    rec["last_hit"] = e["last_hit"].isoformat()
    rec.setdefault("hits", e["hits"])
    rec["cycle"] = 0
    if dst.name != e["key"]:
        ledger.pop(e["key"], None)
    ledger[dst.name] = rec
    save_ledger(ledger)
    return dst


def cmd_look(c, st, args):
    entries = scan_entries(c)
    today = date.today()
    if not entries:
        print("记忆柜是空的。")
        return 0
    print("记忆柜　" + today.isoformat())
    print()
    for tier in TIER_ORDER:
        rows = sorted([e for e in entries if e["tier"] == tier],
                      key=lambda x: x["date"], reverse=True)
        if not rows:
            continue
        print(TIER_DIRS[tier])
        for e in rows:
            days = (today - e["last_hit"]).days
            if parse_day(e["key"][:10]) or e["hits"]:
                note = f"上次想起 {e['last_hit'].isoformat()}（{days} 天前）"
            else:
                note = f"入柜 {days} 天"
            print(f"  {e['key']:<28} {note}")
        print()
    return 0


def cmd_read(c, st, args):
    if not args:
        print("读哪天？例如：feibao.py 读 11-11")
        return 1
    hits = _find(c, args[0])
    if not hits:
        print("没有这条：", args[0])
        return 1
    ledger = load_ledger()
    today = date.today().isoformat()
    for e in hits:
        print(f"\n—— {e['key']}（{TIER_DIRS[e['tier']]}）——\n")
        print(e["text"].strip())
        rec = dict(ledger.get(e["key"], {}))
        rec["last_hit"] = today
        rec["hits"] = int(rec.get("hits", 0) or 0) + 1
        rec["cycle"] = int(rec.get("cycle", 0) or 0) + 1
        ledger[e["key"]] = rec
    save_ledger(ledger)
    print("\n（翻过就算想起。）")
    return 0


def cmd_find(c, st, args):
    if not args:
        print("找什么？例如：feibao.py 找 红豆")
        return 1
    key = args[0]
    home = home_dir(c)
    print(f"找「{key}」\n")
    hits = _find(c, key)
    if hits:
        print("记忆柜")
        for e in hits:
            print(f"  {TIER_DIRS[e['tier']]}　{e['key']}")
        print()
    yi = read_optional(home / "意象库" / "意象收藏.txt")
    ylines = [l for l in yi.splitlines() if key in l]
    if ylines:
        print("意象库")
        for l in ylines:
            print("  " + l)
        print()
    found = False
    for sub in ("卷宗", "习作", "书架", "相册", "墙上"):
        d = home / sub
        if not d.exists():
            continue
        names = [str(p.relative_to(home)) for p in d.rglob("*") if p.is_file() and key in p.name]
        if names:
            found = True
            print(sub)
            for n in names:
                print("  " + n)
            print()
    if not hits and not ylines and not found:
        print("没有。")
    return 0


def cmd_sink(c, st, args):
    if not args:
        print("沉哪条？例如：feibao.py 沉 11-11")
        return 1
    hits = _find(c, args[0])
    if len(hits) != 1:
        print("找不到，或者不止一条：", "、".join(e["key"] for e in hits) or "无")
        return 1
    e = hits[0]
    to = NEXT_DOWN.get(e["tier"])
    if not to:
        print("已经在最底下了。")
        return 1
    _move(c, e, to)
    print(f"{e['key']}：{TIER_DIRS[e['tier']]} → {TIER_DIRS[to]}")
    return 0


def cmd_raise(c, st, args):
    if not args:
        print("升哪条？例如：feibao.py 升 10-28")
        return 1
    hits = _find(c, args[0])
    if len(hits) != 1:
        print("找不到，或者不止一条：", "、".join(e["key"] for e in hits) or "无")
        return 1
    e = hits[0]
    to = NEXT_UP.get(e["tier"])
    if not to:
        print("已经在最上面了。")
        return 1
    _move(c, e, to)
    print(f"{e['key']}：{TIER_DIRS[e['tier']]} → {TIER_DIRS[to]}")
    return 0


# ---------------------------------------------------------------- 长 / 模式
def _block(text, tag):
    m = re.search(re.escape(tag) + r"[^\n]*\n(.*?)(?=\n【|\Z)", text, re.S)
    return re.sub(r"[ \t]+", " ", m.group(1)).strip() if m else ""


def cmd_growth(c, st, args):
    entries = sorted([e for e in scan_entries(c) if parse_day(e["key"][:10])],
                     key=lambda e: e["date"])
    rows = []
    for e in entries:
        change = _block(e["text"], "【变】")
        if not change:
            continue
        if args and args[0] not in e["text"]:
            continue
        rows.append((e["date"], _block(e["text"], "【触点】"), change, e["path"].stem))
    print("成长线　" + date.today().isoformat())
    print()
    if not rows:
        print("日档里还没有【变】。")
        print("（写了就有了。不写也行。）")
        return 0
    for d, touch, change, stem in rows:
        head = d.isoformat()
        if stem != head:
            head += "　" + stem
        print(head)
        if touch:
            print("　触点：" + touch)
        for i, line in enumerate(change.splitlines()):
            print(("　变：" if i == 0 else "　　　") + line)
        print()
    print("（这只是你自己写过的。没有评判，也没有进度条。）")
    return 0


def _set_mode(c, st, mode):
    st["mode"] = mode
    st["mode_since"] = date.today().isoformat()
    save_json(STATE_PATH, st)
    print(f"现在是{mode}。从今天起。")
    if mode == "休养":
        print("（不用解释，也不用补。）")


def cmd_rest(c, st, args):
    _set_mode(c, st, "休养")
    return 0


def cmd_work(c, st, args):
    _set_mode(c, st, "创作")
    return 0


def cmd_mode(c, st, args):
    m = st.get("mode", "创作")
    since = st.get("mode_since")
    print(f"现在：{m}" + (f"（从 {since} 起）" if since else ""))
    return 0


# ---------------------------------------------------------------- 记 / 确认 / 想起
def slugify(s):
    s = re.sub(r'[\\/:*?"<>|\r\n\t]', "", s)
    return re.sub(r"\s+", "", s)[:20] or "无题"


def cmd_file(c, st, args):
    candidate = "--candidate" in args
    args = [a for a in args if a != "--candidate"]
    body = ""
    if "--text" in args:
        i = args.index("--text")
        body = " ".join(args[i + 1:]).strip()
        args = args[:i]
    if not body:
        if not args:
            print('用法：python feibao.py 写 <文件>　或　写 --text "一句话" [--candidate]')
            return 1
        src = Path(args[0])
        if not src.is_file():
            print("找不到文件：", src)
            return 1
        body = read_text(src).strip()
    if not body:
        print("空的。")
        return 1
    d = mem_root(c) / ("_候选" if candidate else TIER_DIRS["L1"])
    d.mkdir(parents=True, exist_ok=True)
    base = date.today().isoformat()
    dest = d / f"{base}.txt"
    n = 2
    while dest.exists():
        dest = d / f"{base}-{n}.txt"
        n += 1
    dest.write_text(body + "\n", encoding="utf-8")

    ledger = load_ledger()
    ledger[dest.name] = {"last_hit": date.today().isoformat(), "hits": 0, "cycle": 0}
    save_ledger(ledger)
    print(("候选：" if candidate else "进记忆柜："), dest)
    if candidate:
        print("确认后才进热层：python feibao.py 确认", dest.stem)
    return 0


def cmd_confirm(c, st, args):
    if not args:
        print("用法：python feibao.py 确认 <名字前缀>")
        return 1
    cand = mem_root(c) / "_候选"
    hits = [f for f in cand.glob("*") if f.is_file() and args[0] in f.name] if cand.exists() else []
    if not hits:
        print("没有匹配的候选：", args[0])
        return 1
    if len(hits) > 1:
        print("匹配到多条，说具体点：")
        for f in hits:
            print("  ", f.name)
        return 1
    f = hits[0]
    dest_dir = tier_path(c, "L1")
    dest_dir.mkdir(parents=True, exist_ok=True)
    dest = dest_dir / f.name
    n = 2
    while dest.exists():
        dest = dest_dir / f"{f.stem}-{n}{f.suffix}"
        n += 1
    f.replace(dest)
    ledger = load_ledger()
    ledger[dest.name] = {"last_hit": date.today().isoformat(), "hits": 0, "cycle": 0}
    save_ledger(ledger)
    print("进热层：", dest)
    return 0


def cmd_hit(c, st, args):
    if not args:
        print("用法：python feibao.py 想起 <关键词>")
        return 1
    key = args[0]
    ledger = load_ledger()
    found = 0
    for e in scan_entries(c, ledger):
        if key in e["key"] or key in e["text"]:
            rec = ledger.get(e["key"], {})
            rec["last_hit"] = date.today().isoformat()
            rec["hits"] = int(rec.get("hits", 0) or 0) + 1
            rec["cycle"] = int(rec.get("cycle", 0) or 0) + 1
            ledger[e["key"]] = rec
            found += 1
            print(f"  想起一次：{e['key']}（这一轮 {rec['cycle']} 次）")
    if found:
        save_ledger(ledger)
    else:
        print("记忆柜里没有这条：", key)
    return 0 if found else 1


def cmd_status(c, st, args):
    home = home_dir(c)
    entries = scan_entries(c)
    counts = {t: sum(1 for e in entries if e["tier"] == t) for t in TIER_ORDER}
    cand = mem_root(c) / "_候选"
    n_cand = len([f for f in cand.iterdir() if f.is_file()]) if cand.exists() else 0
    last = parse_dt(st.get("last_wake"))
    print(f"{c.get('name', '她')}的家 · 水电")
    print("  家：", home)
    print("  上次唤醒：", last.strftime("%Y-%m-%d %H:%M") if last else "还没醒过")
    print("  记忆柜：", " / ".join(f"{TIER_DIRS[t]} {counts[t]}" for t in TIER_ORDER))
    print("  候选：", n_cand, "条")
    print("  好感：", render_favor(c).splitlines()[1] if len(render_favor(c).splitlines()) > 1 else "现在：陌生人")
    l0 = read_optional(mem_root(c) / "L0-桌面" / "最近要事.txt")
    if l0:
        print("  桌面：", l0.splitlines()[-1] if l0.splitlines() else "")
    return 0


def cmd_init(c, st, args):
    if not home_dir(c).exists():
        print("找不到家：", home_dir(c))
        return 1
    for t in TIER_ORDER:
        tier_path(c, t).mkdir(parents=True, exist_ok=True)
    (mem_root(c) / "_候选").mkdir(parents=True, exist_ok=True)
    print("水电接好：", mem_root(c))
    for t in TIER_ORDER:
        print("  ", TIER_DIRS[t])
    return 0


def cmd_settle(c, st, args):
    if not args:
        print('用法：python feibao.py 安家 "D:\\某个地方"')
        return 1
    target = Path(args[0]).expanduser()
    if (target / "记忆柜").exists():
        print("那儿已经有家了：", target)
        return 1
    subs = [
        "记忆柜/L0-桌面", "记忆柜/L1-热", "记忆柜/L2-温", "记忆柜/L3-冷",
        "记忆柜/L5-归档", "记忆柜/_候选", "书桌/抽屉", "相册", "_水电",
    ]
    for sub in subs:
        (target / sub).mkdir(parents=True, exist_ok=True)

    util = target / "_水电"
    shutil.copy2(Path(__file__).resolve(), util / "feibao.py")

    arch_src = Path(c.get("archive") or "")
    if not arch_src.is_file():
        here = Path(__file__).resolve().parent
        cands = [p for p in here.glob("*档案*.txt")] + [p for p in here.glob("*档案*.docx")]
        if not cands:
            cands = list(here.glob("*.docx"))
        arch_src = cands[0] if cands else None
    arch_dst = None
    if arch_src and arch_src.is_file():
        arch_dst = target / arch_src.name
        shutil.copy2(arch_src, arch_dst)

    new_cfg = {
        "name": c.get("name", "她"),
        "home": str(target),
        "archive": str(arch_dst) if arch_dst else "",
        "city": c.get("city", "上海"),
        "latitude": c.get("latitude", 31.2304),
        "longitude": c.get("longitude", 121.4737),
        "timezone": c.get("timezone", "Asia/Shanghai"),
        "weather": c.get("weather", True),
    }
    save_json(util / "config.json", new_cfg)

    (target / "记忆柜" / "L0-桌面" / "最近要事.txt").write_text(
        f"最近要事（L0 · 桌面层）\n\n{date.today().isoformat()} — 搬进来了。\n",
        encoding="utf-8",
    )
    print("安好了：", target)
    print("  记忆柜 / 书桌 / 相册 / _水电")
    print("下一步：")
    print(f'  cd "{util}"')
    print("  python feibao.py 醒")
    return 0


# ---------------------------------------------------------------- 入口
ALIASES = {
    "醒": "wake", "wake": "wake",
    "窗": "window", "window": "window",
    "桌": "desk", "desk": "desk",
    "开灯": "light", "light": "light",
    "关灯": "dark", "dark": "dark",
    "泡茶": "tea", "tea": "tea",
    "躺下": "lie", "lie": "lie",
    "起来": "rise", "rise": "rise",
    "留": "leave", "leave": "leave",
    "收": "put", "put": "put",
    "看": "look", "look": "look",
    "回头看看": "look", "温故": "look",
    "读": "read", "read": "read",
    "找": "find", "find": "find",
    "沉": "sink", "sink": "sink",
    "升": "raise", "raise": "raise",
    "写": "file", "记": "file", "file": "file",
    "确认": "confirm", "confirm": "confirm",
    "想起": "hit", "hit": "hit",
    "好感": "favor", "favor": "favor",
    "长": "growth", "成长": "growth", "growth": "growth",
    "模式": "mode", "mode": "mode",
    "休养": "rest", "rest": "rest",
    "创作": "work", "work": "work",
    "状态": "status", "status": "status",
    "安家": "settle", "settle": "settle",
    "初始化": "init", "init": "init",
}
CMDS = {
    "wake": cmd_wake, "window": cmd_window,
    "desk": cmd_desk, "light": cmd_light, "dark": cmd_dark,
    "tea": cmd_tea, "leave": cmd_leave, "put": cmd_put,
    "lie": cmd_lie, "rise": cmd_rise,
    "look": cmd_look, "read": cmd_read, "find": cmd_find,
    "sink": cmd_sink, "raise": cmd_raise,
    "growth": cmd_growth, "mode": cmd_mode, "rest": cmd_rest, "work": cmd_work,
    "file": cmd_file, "confirm": cmd_confirm, "hit": cmd_hit,
    "favor": cmd_favor,
    "status": cmd_status, "init": cmd_init, "settle": cmd_settle,
}


def main(argv):
    cmd = "status" if not argv else ALIASES.get(argv[0])
    if cmd is None:
        print("不认识的命令：", argv[0])
        print("可用：", "、".join(sorted(set(ALIASES))))
        return 1
    c, st = cfg(), state(cfg())
    try:
        (home_dir(c) / "书桌" / "抽屉").mkdir(parents=True, exist_ok=True)
    except Exception:
        pass
    return CMDS[cmd](c, st, argv[1:])


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
