"""生成应用图标 assets/app_icon.ico（Element 蓝 + 白，无红色）。"""
from pathlib import Path

from PIL import Image, ImageDraw

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "assets" / "app_icon.ico"

PRIMARY = "#409EFF"
PRIMARY_DARK = "#3A8EE6"
WHITE = (255, 255, 255, 255)
WHITE_SOFT = (255, 255, 255, 200)


def _draw_icon(size: int) -> Image.Image:
    w = h = size
    img = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    pad = max(2, w // 14)
    radius = max(4, w // 6)

    # 圆角底：Element 蓝渐变感（外深内亮）
    d.rounded_rectangle(
        (pad, pad, w - pad, h - pad),
        radius=radius,
        fill=PRIMARY,
    )
    inset = max(1, w // 40)
    d.rounded_rectangle(
        (pad + inset, pad + inset, w - pad - inset, h - pad - inset),
        radius=max(3, radius - 1),
        fill=PRIMARY_DARK,
    )
    d.rounded_rectangle(
        (pad + inset * 2, pad + inset * 2, w - pad - inset * 2, h - pad - inset * 2),
        radius=max(2, radius - 2),
        fill=PRIMARY,
    )

    # 白色表格（3x3 网格）
    m = int(w * 0.24)
    inner = w - 2 * m
    cell = inner / 3
    x0, y0 = float(m), float(m)
    lw = max(1, w // 28)
    for i in range(4):
        x = x0 + i * cell
        d.line([(x, y0), (x, y0 + inner)], fill=WHITE, width=lw)
    for j in range(4):
        y = y0 + j * cell
        d.line([(x0, y), (x0 + inner, y)], fill=WHITE, width=lw)

    # 右下角：白色「拆分」双页（不用红/绿）
    if w >= 24:
        sx, sy = w * 0.58, h * 0.52
        sw, sh = w * 0.28, h * 0.30
        off = max(2, w // 20)
        d.rounded_rectangle(
            (sx, sy, sx + sw, sy + sh),
            radius=max(1, w // 32),
            fill=WHITE_SOFT,
            outline=WHITE,
            width=max(1, lw),
        )
        d.rounded_rectangle(
            (sx + off, sy - off, sx + off + sw, sy - off + sh),
            radius=max(1, w // 32),
            fill=WHITE,
            outline=WHITE,
            width=max(1, lw),
        )

    return img


def main() -> None:
    OUT.parent.mkdir(parents=True, exist_ok=True)
    sizes = [(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)]
    images = [_draw_icon(s) for s, _ in sizes]

    images[0].save(
        OUT,
        format="ICO",
        sizes=[(im.width, im.height) for im in images],
        append_images=images[1:],
    )
    print(f"Wrote {OUT}")


if __name__ == "__main__":
    main()
