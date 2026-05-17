"""生成应用图标 assets/app_icon.ico"""
from pathlib import Path

from PIL import Image, ImageDraw

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "assets" / "app_icon.ico"


def main() -> None:
    OUT.parent.mkdir(parents=True, exist_ok=True)
    sizes = [(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)]
    images: list[Image.Image] = []

    for w, h in sizes:
        img = Image.new("RGBA", (w, h), (0, 0, 0, 0))
        d = ImageDraw.Draw(img)
        pad = max(2, w // 16)
        r = pad
        # 圆角矩形底（Element 蓝）
        d.rounded_rectangle(
            (pad, pad, w - pad, h - pad),
            radius=max(4, w // 8),
            fill="#409EFF",
        )
        # 白色表格网格
        margin = w * 0.22
        inner = w - 2 * margin
        cell = inner / 3
        x0, y0 = margin, margin
        line_w = max(1, w // 32)
        white = (255, 255, 255, 240)
        for i in range(4):
            x = x0 + i * cell
            d.line([(x, y0), (x, y0 + inner)], fill=white, width=line_w)
        for j in range(4):
            y = y0 + j * cell
            d.line([(x0, y), (x0 + inner, y)], fill=white, width=line_w)
        # 右下角拆分箭头（浅绿点缀）
        if w >= 32:
            ax, ay = w * 0.62, h * 0.62
            d.polygon(
                [
                    (ax, ay),
                    (ax + w * 0.22, ay + h * 0.08),
                    (ax + w * 0.08, ay + h * 0.22),
                ],
                fill="#A0E7A0",
            )
        images.append(img)

    images[0].save(
        OUT,
        format="ICO",
        sizes=[(im.width, im.height) for im in images],
        append_images=images[1:],
    )
    print(f"Wrote {OUT}")


if __name__ == "__main__":
    main()
