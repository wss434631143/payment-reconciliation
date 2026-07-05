from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "assets" / "usage-demo.gif"
W, H = 960, 540

BG = "#f5f8fd"
INK = "#1f2a44"
MUTED = "#667085"
BLUE = "#2563eb"
GREEN = "#16a34a"
RED = "#dc2626"
BORDER = "#cbd5e1"
HEADER = "#eaf1fb"


def load_font(size):
    for path in [
        r"C:\Windows\Fonts\msyh.ttc",
        r"C:\Windows\Fonts\simhei.ttf",
        r"C:\Windows\Fonts\simsun.ttc",
    ]:
        if Path(path).exists():
            return ImageFont.truetype(path, size)
    return ImageFont.load_default()


F_TITLE = load_font(28)
F_H2 = load_font(20)
F = load_font(16)
F_SMALL = load_font(13)
F_TINY = load_font(12)

STEPS = [
    ("1. 选择店铺和年月", "先选店铺、报表类型，再选择查询年月区间。"),
    ("2. 导入流水", "一个弹窗里完成店铺、报表类型、归属年月和文件选择。"),
    ("3. 查看汇总", "汇总表按年月展示，冻结栏和活动栏可拖动调宽。"),
    ("4. 筛选明细", "连续加入多个筛选条件，点击开始筛选后再查询大表。"),
    ("5. 差异与调整", "核对结算期末余额和店铺期末余额，记录调整说明。"),
    ("6. 导出与备份", "导出 XLSX，或备份/还原全部及部分店铺数据库。"),
]

SUMMARY_COLS = ["店铺", "报表类型", "年月", "收入净额合计", "支出金额", "结算金额", "差异"]
SUMMARY_ROWS = [
    ["抖音专营店", "已结算", "2026-05", "3585731.13", "-582794.88", "3002936.25", "0.00"],
    ["抖音专营店", "已结算", "2026-06", "5511150.77", "-1034133.47", "4477017.30", "-128.60"],
]
DETAIL_COLS = ["动账时间", "动账方向", "动账金额", "动账摘要", "主订单编号"]
DETAIL_ROWS = [
    ["2026-06-01 09:18", "入账", "138.96", "货款结算入账", "692644488449908171"],
    ["2026-06-01 09:25", "入账", "61.83", "货款结算入账", "692646244815953871"],
    ["2026-06-01 10:02", "出账", "-8.40", "技术服务费", "692647187046855304"],
]


def draw_text(draw, xy, value, fill=INK, font=F, anchor=None):
    draw.text(xy, value, fill=fill, font=font, anchor=anchor)


def round_rect(draw, box, fill, outline=None, radius=10, width=1):
    draw.rounded_rectangle(box, radius=radius, fill=fill, outline=outline, width=width)


def button(draw, x, y, label, active=False, width=None):
    text_width = draw.textbbox((0, 0), label, font=F_SMALL)[2]
    width = width or max(74, text_width + 28)
    fill = BLUE if active else "#eef4ff"
    fg = "white" if active else "#2754a3"
    round_rect(draw, (x, y, x + width, y + 34), fill, "#9bb7ee", 6)
    draw_text(draw, (x + width / 2, y + 17), label, fg, F_SMALL, "mm")
    return x + width + 8


def select_box(draw, x, y, label, value, width=180):
    draw_text(draw, (x, y + 18), label, MUTED, F_SMALL, "lm")
    round_rect(draw, (x + 54, y, x + 54 + width, y + 34), "white", BORDER, 6)
    draw_text(draw, (x + 68, y + 17), value, INK, F_SMALL, "lm")
    draw_text(draw, (x + 54 + width - 18, y + 17), "▼", "#64748b", F_TINY, "mm")
    return x + 64 + width


def draw_window(draw):
    round_rect(draw, (18, 16, W - 18, H - 18), "white", "#d7e0ee", 12)
    draw.rectangle((18, 16, W - 18, 72), fill="#f8fbff", outline="#d7e0ee")
    draw_text(draw, (36, 45), "财务第三方支付核对 - Qt版", INK, F_TITLE, "lm")
    x = 36
    for label in ["导入流水", "店铺", "参数", "余额", "调整", "差异", "导出", "备份", "设置"]:
        x = button(draw, x, 88, label, active=label in {"导入流水", "导出", "备份"})


def draw_progress(draw, index):
    x0, y0, x1 = 56, 490, W - 56
    draw.line((x0, y0, x1, y0), fill="#dbe4f0", width=8)
    draw.line((x0, y0, x0 + (x1 - x0) * (index + 1) / len(STEPS), y0), fill=BLUE, width=8)
    for i in range(len(STEPS)):
        x = x0 + (x1 - x0) * i / (len(STEPS) - 1)
        fill = BLUE if i <= index else "#cbd5e1"
        draw.ellipse((x - 8, y0 - 8, x + 8, y0 + 8), fill=fill, outline="white", width=2)


def draw_table(draw, x, y, columns, rows, widths, highlight_col=None):
    height = 30
    cx = x
    for i, col in enumerate(columns):
        draw.rectangle((cx, y, cx + widths[i], y + height), fill=HEADER, outline=BORDER)
        draw_text(draw, (cx + 8, y + height / 2), col, "#475569", F_TINY, "lm")
        cx += widths[i]
    for r, row in enumerate(rows):
        cx = x
        for i, value in enumerate(row):
            fill = "#f8fbff" if highlight_col == i else "white"
            draw.rectangle((cx, y + height * (r + 1), cx + widths[i], y + height * (r + 2)), fill=fill, outline="#e2e8f0")
            draw_text(draw, (cx + 8, y + height * (r + 1) + height / 2), value, "#334155", F_TINY, "lm")
            cx += widths[i]


def make_frame(index):
    img = Image.new("RGB", (W, H), BG)
    draw = ImageDraw.Draw(img)
    draw_window(draw)
    title, subtitle = STEPS[index]
    draw_text(draw, (42, 142), title, BLUE, F_H2)
    draw_text(draw, (42, 172), subtitle, MUTED, F_SMALL)

    if index == 0:
        x = 42
        x = select_box(draw, x, 214, "店铺", "抖音专营店", 210)
        x = select_box(draw, x + 10, 214, "报表", "已结算", 110)
        select_box(draw, x + 10, 214, "年月", "2026-05 至 2026-06", 210)
        button(draw, 795, 214, "查询", True, 84)
        draw_table(draw, 42, 270, SUMMARY_COLS[:4], [r[:4] for r in SUMMARY_ROWS], [120, 90, 80, 150], 2)
    elif index == 1:
        round_rect(draw, (250, 174, 710, 420), "#ffffff", "#b8c7dc", 12)
        draw_text(draw, (282, 210), "导入流水设置", INK, F_H2)
        select_box(draw, 282, 245, "店铺", "抖音专营店", 260)
        select_box(draw, 282, 292, "报表", "已结算", 260)
        select_box(draw, 282, 339, "年月", "2026 年 06 月", 260)
        button(draw, 454, 382, "开始导入", True, 112)
        button(draw, 576, 382, "退出", False, 82)
    elif index == 2:
        draw_table(draw, 42, 220, SUMMARY_COLS, SUMMARY_ROWS, [112, 82, 72, 128, 108, 108, 82], 6)
        draw.line((42 + 112 + 82, 214, 42 + 112 + 82, 332), fill=BLUE, width=3)
        draw_text(draw, (46, 360), "状态：汇总 2 条，活动栏可横向滚动，冻结区可拖动调宽", GREEN, F_SMALL)
    elif index == 3:
        select_box(draw, 42, 214, "字段", "动账方向", 180)
        select_box(draw, 300, 214, "值", "入账", 210)
        button(draw, 590, 214, "加入条件", False, 100)
        button(draw, 700, 214, "开始筛选", True, 110)
        draw_text(draw, (42, 266), "已应用：动账方向 = 入账；动账摘要 = 货款结算入账", BLUE, F_SMALL)
        draw_table(draw, 42, 300, DETAIL_COLS, DETAIL_ROWS[:2], [145, 90, 90, 150, 190], 1)
    elif index == 4:
        round_rect(draw, (48, 214, 912, 384), "#fff7ed", "#fed7aa", 10)
        draw_text(draw, (72, 248), "核对摘要", INK, F_H2)
        labels = [("结算期末余额", "4,477,017.30", GREEN), ("店铺期末余额", "4,477,145.90", BLUE), ("差异", "-128.60", RED)]
        x = 72
        for name, value, color in labels:
            round_rect(draw, (x, 282, x + 230, 344), "white", "#fed7aa", 8)
            draw_text(draw, (x + 18, 306), name, MUTED, F_SMALL)
            draw_text(draw, (x + 18, 332), value, color, F_H2)
            x += 260
        button(draw, 624, 248, "查看差异明细", False, 132)
        button(draw, 766, 248, "新增调整", True, 108)
    else:
        round_rect(draw, (60, 216, 424, 378), "white", BORDER, 10)
        draw_text(draw, (88, 256), "导出汇总 XLSX", INK, F_H2)
        draw_text(draw, (88, 292), "包含汇总表、原始表格、店铺、年月区间和时间戳", MUTED, F_SMALL)
        button(draw, 88, 326, "导出", True, 94)
        round_rect(draw, (506, 216, 872, 378), "white", BORDER, 10)
        draw_text(draw, (534, 256), "备份还原", INK, F_H2)
        draw_text(draw, (534, 292), "可按全部或部分店铺备份/还原数据库和配置", MUTED, F_SMALL)
        button(draw, 534, 326, "备份", True, 94)
        button(draw, 638, 326, "还原", False, 94)
    draw_progress(draw, index)
    return img


def main():
    OUT.parent.mkdir(parents=True, exist_ok=True)
    frames = []
    for i in range(len(STEPS)):
        frames.extend([make_frame(i)] * 2)
    frames[0].save(OUT, save_all=True, append_images=frames[1:], duration=850, loop=0, optimize=True)
    print(OUT)
    print(OUT.stat().st_size)


if __name__ == "__main__":
    main()
