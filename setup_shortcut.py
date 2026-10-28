"""
カレンダーアイコン付きショートカットを作成するセットアップスクリプト
実行方法: このフォルダで「python setup_shortcut.py」を実行
"""
import sys, os, subprocess
from pathlib import Path

APP_DIR = Path(__file__).parent.resolve()

# ── Pillow が無ければインストール ──
try:
    from PIL import Image, ImageDraw, ImageFont
except ImportError:
    print("Pillowをインストール中...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pillow"])
    from PIL import Image, ImageDraw, ImageFont

# ── pywin32 が無ければインストール ──
try:
    import win32com.client
except ImportError:
    print("pywin32をインストール中...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pywin32"])
    import win32com.client


def make_calendar_icon(size):
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    pad = max(2, size // 16)
    r   = max(4, size // 12)

    # 本体（白・角丸）
    body = [pad, pad + size // 8, size - pad, size - pad]
    d.rounded_rectangle(body, radius=r,
                        fill=(255, 255, 255, 255),
                        outline=(180, 180, 190, 255),
                        width=max(1, size // 32))

    # ヘッダー（青）
    hdr = [pad, pad + size // 8, size - pad, pad + size // 8 + size // 4]
    d.rounded_rectangle(hdr, radius=r, fill=(25, 118, 210, 255))
    fix = [pad, pad + size // 8 + size // 8, size - pad, pad + size // 8 + size // 4]
    d.rectangle(fix, fill=(25, 118, 210, 255))

    # リング（上部2本）
    rw = max(2, size // 20)
    rh = size // 7
    for cx in [size // 3, size * 2 // 3]:
        d.rounded_rectangle(
            [cx - rw, pad - rh // 3, cx + rw, pad - rh // 3 + rh],
            radius=rw, fill=(100, 149, 237, 255))

    # 日付グリッド
    cy0  = pad + size // 8 + size // 4 + size // 16
    ch   = (size - pad - cy0) // 3
    cw   = (size - pad * 2) // 4
    fsz  = max(6, size // 9)
    try:
        font = ImageFont.truetype("arial.ttf", fsz)
    except Exception:
        font = ImageFont.load_default()

    for i, day in enumerate(["1","2","3","4","5","6"]):
        col, row = i % 3, i // 3
        tx = pad + cw // 2 + col * cw + cw // 3
        ty = cy0 + row * ch + ch // 4
        color = (25, 118, 210, 255) if i == 0 else (60, 60, 60, 255)
        try:
            d.text((tx, ty), day, fill=color, font=font, anchor="mm")
        except Exception:
            d.text((tx, ty), day, fill=color, font=font)
    return img


def create_icon():
    ico_path = APP_DIR / "calendar.ico"
    sizes = [256, 64, 32, 16]
    frames = [make_calendar_icon(s) for s in sizes]
    frames[0].save(str(ico_path), format="ICO",
                   sizes=[(s, s) for s in sizes],
                   append_images=frames[1:])
    print(f"アイコン作成: {ico_path}")
    return ico_path


def create_shortcut(ico_path: Path):
    import win32com.client as wc
    shell   = wc.Dispatch("WScript.Shell")
    lnk_path = APP_DIR / "勤務スケジュール.lnk"
    sc = shell.CreateShortcut(str(lnk_path))
    sc.TargetPath      = str(APP_DIR / "起動.bat")
    sc.WorkingDirectory = str(APP_DIR)
    sc.IconLocation    = str(ico_path)
    sc.Description     = "大西 勤務スケジュール管理"
    sc.WindowStyle     = 7   # 最小化して起動（コンソール非表示）
    sc.Save()
    print(f"ショートカット作成: {lnk_path}")
    return lnk_path


if __name__ == "__main__":
    print("セットアップを開始します...")
    ico  = create_icon()
    lnk  = create_shortcut(ico)
    print()
    print("完了！")
    print(f"「{lnk.name}」をダブルクリックしてアプリを起動してください。")
    input("\nEnterキーで閉じる...")
