from pathlib import Path
from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "xiaohongshu" / "v2-source"
OUT = ROOT / "xiaohongshu" / "editorial"

W, H = 1200, 1600
IVORY = "#F3F0E8"
INK = "#173B68"
MUTED = "#6B7580"
PALE_BLUE = "#DCE8F2"
MID_BLUE = "#5F88B2"
LIGHT_BLUE = "#AFC9DD"
CORAL = "#CB725E"
GREEN = "#769486"

FONT_SONG = Path(r"C:\Windows\Fonts\STSONG.TTF")
FONT_HEI = Path(r"C:\Windows\Fonts\msyhbd.ttc")
FONT_YAHEI = Path(r"C:\Windows\Fonts\msyh.ttc")
FONT_SERIF = Path(r"C:\Windows\Fonts\times.ttf")
FONT_SERIF_ITALIC = Path(r"C:\Windows\Fonts\timesi.ttf")


def font(path: Path, size: int):
    return ImageFont.truetype(str(path), size)


def fit_text(draw, text, box, font_path, start, minimum=24, spacing=8, align="left"):
    x1, y1, x2, y2 = box
    for size in range(start, minimum - 1, -1):
        f = font(font_path, size)
        bbox = draw.multiline_textbbox(
            (0, 0), text, font=f, spacing=spacing, align=align
        )
        if bbox[2] - bbox[0] <= x2 - x1 and bbox[3] - bbox[1] <= y2 - y1:
            return f
    return font(font_path, minimum)


def contain(im, size, background=IVORY):
    w, h = size
    result = Image.new("RGB", (w, h), background)
    source = im.copy()
    source.thumbnail((w, h), Image.Resampling.LANCZOS)
    px = (w - source.width) // 2
    py = (h - source.height) // 2
    result.paste(source, (px, py))
    return result


def cover_crop(im, size):
    tw, th = size
    scale = max(tw / im.width, th / im.height)
    resized = im.resize(
        (round(im.width * scale), round(im.height * scale)),
        Image.Resampling.LANCZOS,
    )
    left = (resized.width - tw) // 2
    top = (resized.height - th) // 2
    return resized.crop((left, top, left + tw, top + th))


def draw_header(draw, index, label):
    draw.text((72, 54), label, font=font(FONT_YAHEI, 25), fill=MID_BLUE)
    draw.text(
        (1128, 53),
        f"{index:02d}",
        font=font(FONT_SERIF, 25),
        fill=MUTED,
        anchor="ra",
    )


def draw_footer(draw, text="nanfengy · AI 科研作图"):
    draw.line((72, 1520, 1128, 1520), fill="#D8D3C9", width=2)
    draw.text((72, 1540), text, font=font(FONT_YAHEI, 19), fill=MUTED)


def save(canvas, index):
    OUT.mkdir(parents=True, exist_ok=True)
    canvas.save(OUT / f"xhs-{index:02d}.png", quality=95)


def build_overview():
    thumb_w, thumb_h = 270, 360
    gap = 28
    margin = 36
    cols, rows = 4, 2
    overview = Image.new(
        "RGB",
        (
            margin * 2 + cols * thumb_w + (cols - 1) * gap,
            margin * 2 + rows * thumb_h + (rows - 1) * gap,
        ),
        IVORY,
    )
    for index in range(1, 8):
        image = Image.open(OUT / f"xhs-{index:02d}.png").convert("RGB")
        image.thumbnail((thumb_w, thumb_h), Image.Resampling.LANCZOS)
        col = (index - 1) % cols
        row = (index - 1) // cols
        x = margin + col * (thumb_w + gap)
        y = margin + row * (thumb_h + gap)
        overview.paste(image, (x, y))
    overview.save(ROOT / "xiaohongshu" / "editorial-overview.jpg", quality=92)


def build_cover():
    # Keep one real PPT page fully intact as the sole source region.
    source = Image.open(SOURCE / "slide-23.png").convert("RGB")
    top_h = 675
    canvas = Image.new("RGB", (W, H), IVORY)
    canvas.paste(source.resize((W, top_h), Image.Resampling.LANCZOS), (0, 0))
    draw = ImageDraw.Draw(canvas)

    # Abstract memory panel derived from the blue modules, arrows, and central hub.
    cy = 955
    draw.line((245, cy, 955, cy), fill=LIGHT_BLUE, width=7)
    for x, y, color, width in [
        (250, 885, MID_BLUE, 135),
        (435, 955, PALE_BLUE, 165),
        (640, 895, CORAL, 135),
        (820, 970, GREEN, 150),
    ]:
        draw.rounded_rectangle(
            (x, y, x + width, y + 54), radius=16, fill=color
        )
    draw.ellipse((555, 860, 675, 980), outline=INK, width=7)
    draw.ellipse((590, 895, 640, 945), fill=INK)
    draw.line((615, 860, 615, 815), fill=INK, width=6)
    draw.line((675, 920, 730, 920), fill=INK, width=6)
    draw.line((555, 920, 500, 920), fill=INK, width=6)

    title = "导师让我给组里讲\nAI 作图"
    title_font = fit_text(
        draw, title, (72, 1070, 1128, 1360), FONT_SONG, 88, 52, spacing=10
    )
    draw.multiline_text(
        (72, 1070),
        title,
        font=title_font,
        fill=INK,
        spacing=10,
    )
    draw.text(
        (76, 1435),
        "FROM IMAGE THINKING TO WORKFLOW",
        font=font(FONT_SERIF, 24),
        fill=CORAL,
    )
    save(canvas, 1)


def editorial_intro(index, source_name, label, title, body, abstract_kind):
    im = Image.open(SOURCE / source_name).convert("RGB")
    canvas = Image.new("RGB", (W, H), IVORY)
    draw = ImageDraw.Draw(canvas)
    draw_header(draw, index, label)

    image_box = (72, 135, 1056, 585)
    fitted = contain(im, (image_box[2], image_box[3]), background="#E6EDF3")
    canvas.paste(fitted, (image_box[0], image_box[1]))

    # Sparse abstract motif, always derived from the PPT's blue modular grammar.
    if abstract_kind == "layers":
        for i, (x, y, w, color) in enumerate(
            [
                (160, 815, 245, MID_BLUE),
                (430, 850, 305, PALE_BLUE),
                (770, 805, 205, CORAL),
            ]
        ):
            draw.rounded_rectangle((x, y, x + w, y + 52), radius=14, fill=color)
            if i < 2:
                draw.line((x + w, y + 26, x + w + 48, y + 26), fill=INK, width=4)
    elif abstract_kind == "routes":
        for r in range(2):
            for c in range(3):
                x = 240 + c * 245
                y = 780 + r * 110
                color = [MID_BLUE, LIGHT_BLUE, CORAL, GREEN, PALE_BLUE, MID_BLUE][
                    r * 3 + c
                ]
                draw.rounded_rectangle((x, y, x + 145, y + 62), radius=16, fill=color)
                if c < 2:
                    draw.line((x + 145, y + 31, x + 190, y + 31), fill=INK, width=4)
    else:
        draw.line((245, 840, 955, 840), fill=LIGHT_BLUE, width=7)
        for i, color in enumerate([MID_BLUE, LIGHT_BLUE, CORAL, GREEN]):
            x = 260 + i * 190
            draw.rounded_rectangle((x, 805, x + 118, 875), radius=16, fill=color)

    title_font = fit_text(
        draw, title, (72, 980, 1128, 1160), FONT_SONG, 62, 38
    )
    draw.multiline_text((72, 980), title, font=title_font, fill=INK, spacing=6)

    body_font = fit_text(
        draw, body, (72, 1210, 1128, 1455), FONT_YAHEI, 34, 25, spacing=14
    )
    draw.multiline_text(
        (72, 1215),
        body,
        font=body_font,
        fill=MUTED,
        spacing=14,
    )
    draw_footer(draw)
    save(canvas, index)


def real_slide(index, slide_num, label, title, crop_box, caption):
    im = Image.open(SOURCE / f"slide-{slide_num}.png").convert("RGB")
    canvas = Image.new("RGB", (W, H), IVORY)
    draw = ImageDraw.Draw(canvas)
    draw_header(draw, index, label)

    title_font = fit_text(draw, title, (72, 105, 1128, 220), FONT_SONG, 52, 34)
    draw.text((72, 115), title, font=title_font, fill=INK)

    full = contain(im, (1056, 594), background="#E4ECF2")
    canvas.paste(full, (72, 245))

    # Magnify the most important region from the exact same slide.
    x1, y1, x2, y2 = crop_box
    crop = im.crop((x1, y1, x2, y2))
    zoom = contain(crop, (1056, 420), background="#E4ECF2")
    canvas.paste(zoom, (72, 870))

    draw.text(
        (72, 1325),
        "原页 + 关键区域放大",
        font=font(FONT_YAHEI, 22),
        fill=CORAL,
    )
    body_font = fit_text(
        draw, caption, (72, 1370, 1128, 1490), FONT_YAHEI, 29, 22, spacing=8
    )
    draw.multiline_text(
        (72, 1370), caption, font=body_font, fill=MUTED, spacing=8
    )
    draw_footer(draw, "完整材料：github.com/nanfengy/ai-research-visualization")
    save(canvas, index)


def build_all():
    build_cover()
    editorial_intro(
        2,
        "slide-7.png",
        "这份分享从哪里来",
        "一场组会，最后整理成了 73 页",
        "最初只想给师弟师妹讲几个 AI 作图工具。\n"
        "真正整理时，我把重点放在了图形思维、路线选择和正式交付。",
        "layers",
    )
    editorial_intro(
        3,
        "slide-23.png",
        "整套内容讲什么",
        "不只是文生图，而是一套能力体系",
        "从图的本质与类型，到六条技术路线、工具选型、完整工作流，\n"
        "再到质量判断、真实案例和科研场景中的边界。",
        "routes",
    )
    editorial_intro(
        4,
        "slide-42.png",
        "最核心的方法",
        "AI 负责提效，人负责判断",
        "大模型梳理结构 → 结构化工具搭骨架 → 生成式探索视觉 → "
        "可编辑化精修。\n完整 PPT 与 PDF 已整理到 GitHub。",
        "workflow",
    )
    real_slide(
        5,
        23,
        "真实 PPT 页面 01",
        "AI 作图的六条技术路线",
        (95, 260, 1825, 920),
        "结构化、代码化、生成式、可控工作流、可编辑化与视频化，"
        "并不是互相替代的关系。",
    )
    real_slide(
        6,
        52,
        "真实 PPT 页面 02",
        "从需求到成图的完整工作流",
        (115, 255, 1805, 940),
        "高质量科研图不是一次生成，而是从任务定义、逻辑拆解到"
        "精修校验的多阶段迭代。",
    )
    real_slide(
        7,
        64,
        "真实 PPT 页面 03",
        "一张科研图怎么做质量检查",
        (125, 245, 1795, 1030),
        "核心信息、阅读顺序、主次层级、视觉干扰、风格统一与准确无误，"
        "都可以变成可执行的检查项。",
    )
    build_overview()


if __name__ == "__main__":
    build_all()
    print(OUT)
