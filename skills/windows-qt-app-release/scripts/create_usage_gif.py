import os
import shutil
import sys
import tempfile
from pathlib import Path

from PIL import Image


ROOT = Path(__file__).resolve().parents[1]
LOCAL_PYSIDE = ROOT.parents[1] / "work" / "pyside6_pkg"
if LOCAL_PYSIDE.exists():
    sys.path.insert(0, str(LOCAL_PYSIDE))
sys.path.insert(0, str(ROOT))

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication

from qt_app import BackupRestoreDialog, GlobalSettingsDialog, ImportOptionsDialog, MainWindow


OUT = ROOT / "assets" / "usage-demo.gif"
FRAME_SIZE = (1280, 760)


def process_events(app, rounds=8):
    for _ in range(rounds):
        app.processEvents()


def capture_widget(widget, app, frames_dir, name):
    process_events(app)
    pixmap = widget.grab()
    png_path = frames_dir / f"{name}.png"
    pixmap.save(str(png_path))
    image = Image.open(png_path).convert("RGB")
    image.thumbnail(FRAME_SIZE, Image.Resampling.LANCZOS)
    canvas = Image.new("RGB", FRAME_SIZE, "#f3f6fb")
    x = (FRAME_SIZE[0] - image.width) // 2
    y = (FRAME_SIZE[1] - image.height) // 2
    canvas.paste(image, (x, y))
    return canvas


def select_if_exists(combo, text):
    index = combo.findText(text)
    if index >= 0:
        combo.setCurrentIndex(index)
        return True
    return False


def prepare_main_window(app):
    win = MainWindow()
    win.resize(1480, 920)
    win.show()
    process_events(app, 12)

    preferred_store = "抖音专营店"
    if hasattr(win, "store_combo"):
        if not select_if_exists(win.store_combo, preferred_store) and win.store_combo.count():
            win.store_combo.setCurrentIndex(0)
    if hasattr(win, "report_type_combo"):
        select_if_exists(win.report_type_combo, "已结算")

    months = []
    try:
        months = win.repo.months()
    except Exception:
        months = []
    months = [
        str(item.get("month_sort") or item.get("month_label") or item.get("month") or "")
        if isinstance(item, dict) else str(item)
        for item in months
    ]
    months = [item for item in months if item]
    if months:
        win.month_filter.setText(",".join(months[-2:]))
        try:
            win.refresh_all()
        except Exception as exc:
            win.set_status(f"演示截图：数据加载失败，{exc}")
    process_events(app, 12)
    return win


def main():
    OUT.parent.mkdir(parents=True, exist_ok=True)
    app = QApplication.instance() or QApplication([])
    frames_dir = Path(tempfile.mkdtemp(prefix="payment_recon_gif_"))
    frames = []
    try:
        win = prepare_main_window(app)
        frames.extend([capture_widget(win, app, frames_dir, "01-main")] * 2)

        stores = [win.store_combo.itemText(i) for i in range(win.store_combo.count())] or ["默认店铺"]
        import_dialog = ImportOptionsDialog(win, stores, win.current_store())
        import_dialog.resize(520, 260)
        import_dialog.show()
        process_events(app, 8)
        frames.extend([capture_widget(import_dialog, app, frames_dir, "02-import-dialog")] * 2)
        import_dialog.close()

        try:
            if win.summary_rows:
                win.select_summary_row(0)
            select_if_exists(win.detail_filter_column_combo, "动账方向")
            win.refresh_detail_filter_values()
            select_if_exists(win.detail_filter_value_combo, "入账")
            win.add_pending_detail_filter()
            process_events(app, 8)
        except Exception:
            pass
        frames.extend([capture_widget(win, app, frames_dir, "03-filter-pending")] * 2)

        try:
            win.apply_detail_filter()
            process_events(app, 8)
        except Exception:
            pass
        frames.extend([capture_widget(win, app, frames_dir, "04-filter-applied")] * 2)

        backup_dialog = BackupRestoreDialog(win, win.repo)
        backup_dialog.resize(920, 620)
        backup_dialog.show()
        process_events(app, 8)
        frames.extend([capture_widget(backup_dialog, app, frames_dir, "05-backup")] * 2)
        backup_dialog.close()

        settings_dialog = GlobalSettingsDialog(win, win.app_settings)
        settings_dialog.resize(420, 180)
        settings_dialog.show()
        process_events(app, 8)
        frames.extend([capture_widget(settings_dialog, app, frames_dir, "06-settings")] * 2)
        settings_dialog.close()

        win.set_status("演示：导入、查询、筛选、导出和备份流程已完成。")
        process_events(app, 8)
        frames.extend([capture_widget(win, app, frames_dir, "07-finish")] * 2)

        frames[0].save(
            OUT,
            save_all=True,
            append_images=frames[1:],
            duration=900,
            loop=0,
            optimize=True,
        )
        print(OUT)
        print(OUT.stat().st_size)
    finally:
        shutil.rmtree(frames_dir, ignore_errors=True)


if __name__ == "__main__":
    main()
