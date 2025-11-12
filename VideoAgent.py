"""
VideoShortAgent — модуль/агент для створення коротких відео з картинки, тексту та аудіо.

Особливості:
- Автоматичне визначення шляху до ImageMagick (для MoviePy TextClip).
- Акуратні 3D-плашки (paper3D / glass) під текстом.
- Багатошаровий текст із екструзією та мʼякою тінню.
- Компактні субтитри (.srt) з пакуванням у 1–2 рядки, опціональний overlay/burn-in.
- Єдиний високорівневий метод create_video(...) + допоміжні ф-ції як «інструменти агента».
- CLI на Typer (python video_short_agent.py create ...).

Залежності:
  pip install moviepy pillow numpy typer[all]
  # Також потрібен ImageMagick (convert) для деяких режимів TextClip.

Приклад CLI:
  python video_short_agent.py create \
    --image examples/bg.jpg \
    --audio examples/voice.mp3 \
    --title "Головний заголовок" \
    --subtitle "Короткий опис, який переноситься у два рядки" \
    --transcript "Повний текст для субтитрів..." \
    --output out.mp4

Приклад з Python:
  from video_short_agent import VideoShortAgent
  agent = VideoShortAgent()
  agent.create_video(
      image_path="examples/bg.jpg",
      text_blocks=["Головний заголовок", "Короткий опис..."],
      audio_path="examples/voice.mp3",
      output_path="out.mp4",
      subtitle_text="Повний текст для субтитрів...",
      burn_subtitles=False,
  )
"""
from __future__ import annotations

import os
import re
import shutil
from dataclasses import dataclass
from typing import Iterable, List, Optional, Tuple

import numpy as np
from PIL import Image, ImageDraw, ImageFilter

from pathlib import Path

import pysrt
from datetime import timedelta


# ==== Ледачий імпорт MoviePy (щоб модуль можна було імпортувати без повної ініціалізації) ====
try:
    from moviepy.config import change_settings as _mp_change_settings
    from moviepy.editor import (
        AudioFileClip,
        ColorClip,
        CompositeVideoClip,
        ImageClip,
        TextClip,
        concatenate_audioclips,
    )
    from moviepy.video.fx.all import fadein, fadeout
except Exception as e:  # pragma: no cover
    AudioFileClip = None  # type: ignore
    ColorClip = None  # type: ignore
    CompositeVideoClip = None  # type: ignore
    ImageClip = None  # type: ignore
    TextClip = None  # type: ignore
    fadein = None  # type: ignore
    fadeout = None  # type: ignore


# =============================================================================================
#                                      КОНФІГУРАЦІЯ
# =============================================================================================

def _autoconfig_imagemagick() -> None:
    """Автоматично налаштувати шлях до ImageMagick для MoviePy (якщо доступний)."""
    try:
        convert_path = shutil.which("convert") or shutil.which("magick")
        if convert_path and '_mp_change_settings' in globals():
            _mp_change_settings({"IMAGEMAGICK_BINARY": convert_path})
            print(f"ImageMagick знайдено: {convert_path}")
        elif '_mp_change_settings' in globals():
            # Популярні шляхи (наприклад, macOS Apple Silicon Homebrew)
            fallback = "/opt/homebrew/bin/convert"
            if os.path.exists(fallback):
                _mp_change_settings({"IMAGEMAGICK_BINARY": fallback})
                print(f"ImageMagick (fallback): {fallback}")
    except Exception:
        pass


# Виклик при імпорті модуля (не критично, якщо не вдасться)
_autoconfig_imagemagick()


# =============================================================================================
#                                  ДОПОМОЖНІ СТРУКТУРИ
# =============================================================================================

@dataclass
class Typography:
    font: str = "Arial-Bold"
    title_size: int = 80
    subtitle_size: int = 55
    title_color: str = "white"
    subtitle_color: str = "white"
    stroke_color: str = "#1a1a1a"
    stroke_width: int = 3
    title_method: str = "caption"     # НОВЕ
    subtitle_method: str = "caption"  # НОВЕ


@dataclass
class Layout:
    size: Tuple[int, int] = (1280, 720)
    title_left_margin: float = 0.08
    subtitle_left_margin: float = 0.08
    subtitle_bottom_margin: float = 0.08
    subtitle_width_ratio: float = 0.62


@dataclass
class BackgroundPlates:
    enabled: bool = True
    opacity: float = 0.65
    padding_left: int = 32
    padding_right: int = 32
    padding_top: int = 32
    padding_bottom: int = 12
    title_radius: int = 22
    subtitle_radius: int = 18


@dataclass
class Anim:
    intro_fade: float = 1.5
    outro_start: float = -2.0
    outro_fade: float = 1.5


@dataclass
class Subtitles:
    text: Optional[str] = None
    srt_path: Optional[str] = None
    burn_in: bool = False
    max_line_chars: int = 38
    max_lines: int = 2
    cps: float = 16.0
    min_dur: float = 0.8
    max_dur: float = 3.2


# # ADD: CLI
# @app.command("create-bullets")
# def cli_create_bullets(
#     image: str = typer.Option(..., help="Шлях до фонового зображення"),
#     title: str = typer.Option(..., help="Title (угорі ліворуч)"),
#     title_audio: str = typer.Option(..., help="MP3 для Title"),
#     title_subs: str = typer.Option(None, help="Текст субтитрів для Title"),
#     bullets_json: str = typer.Option(..., help='JSON: [{"text": "...", "audio": "...", "subs": "..."}]'),
#     output: str = typer.Option("output_bullets.mp4", help="Куди зберегти відео"),
#     burn_subs: bool = typer.Option(False, help="Підпалити субтитри"),
# ):
#     import json
#     bullets = json.loads(bullets_json)
#     agent = VideoShortAgent()
#     agent.create_video_bulleted_sequence(
#         image_path=image,
#         title_text=title,
#         title_audio_path=title_audio,
#         title_subs_text=title_subs,
#         bullets=bullets,
#         output_path=output,
#         burn_subtitles=burn_subs,
#     )
    

# =============================================================================================
#                                    SRT / ТЕКСТ УТИЛІТИ
# =============================================================================================

def _format_srt_time(t: float) -> str:
    if t < 0:
        t = 0
    h = int(t // 3600)
    m = int((t % 3600) // 60)
    s = int(t % 60)
    ms = int(round((t - int(t)) * 1000))
    return f"{h:02d}:{m:02d}:{s:02d},{ms:03d}"


def _write_srt(segments: List[Tuple[float, float, str]], path: str) -> None:
    with open(path, "w", encoding="utf-8") as f:
        for i, (start, end, text) in enumerate(segments, 1):
            f.write(f"{i}\n{_format_srt_time(start)} --> {_format_srt_time(end)}\n{text}\n\n")


def _sentences(text: str) -> List[str]:
    text = text.strip()
    parts = re.split(r'(?<=[\.!?])\s+|\n+', text)
    return [p.strip() for p in parts if p and p.strip()]


def _clauses(sentence: str) -> List[str]:
    parts = re.split(r'(?<=[,;:—–\-])\s+', sentence)
    return [p.strip() for p in parts if p.strip()]


def _pack_words_to_chunks(text: str, max_line_chars: int = 38, max_lines: int = 2) -> List[str]:
    words = text.split()
    chunks: List[str] = []
    buf: List[str] = []
    max_chars = max_line_chars * max_lines

    def cur_len(lst: List[str]) -> int:
        return len(" ".join(lst))

    for w in words:
        if not buf:
            buf.append(w)
            continue
        if cur_len(buf + [w]) <= max_chars:
            buf.append(w)
        else:
            chunks.append(" ".join(buf))
            buf = [w]
    if buf:
        chunks.append(" ".join(buf))

    MIN_CHARS = max(8, max_line_chars // 2)
    if len(chunks) >= 2 and len(chunks[-1]) < MIN_CHARS:
        tail = chunks.pop()
        chunks[-1] = f"{chunks[-1]} {tail}"
    return chunks


def _compact_segments_from_text(text: str, *, max_line_chars: int, max_lines: int) -> List[str]:
    segs: List[str] = []
    for sent in _sentences(text):
        for cl in _clauses(sent):
            segs.extend(_pack_words_to_chunks(cl, max_line_chars=max_line_chars, max_lines=max_lines))
    return [s for s in segs if s.strip()]


def _assign_times_by_cps(
    segments: List[str],
    total_duration: float,
    *,
    cps: float = 15.0,
    min_dur: float = 0.8,
    max_dur: float = 3.5,
) -> List[Tuple[float, float, str]]:
    if not segments:
        return []
    base = [max(1, len(re.sub(r"\s+", "", s))) / max(1e-9, cps) for s in segments]
    clamped = [min(max_dur, max(min_dur, d)) for d in base]
    scale = total_duration / sum(clamped)
    durs = [d * scale for d in clamped]

    out: List[Tuple[float, float, str]] = []
    t = 0.0
    for s, d in zip(segments, durs):
        start = t
        end = min(total_duration, t + d)
        out.append((start, end, s))
        t = end
    if out:
        out[-1] = (out[-1][0], total_duration, out[-1][2])
    return out


def build_srt_segments_compact(
    full_text: str,
    total_duration: float,
    *,
    max_line_chars: int = 38,
    max_lines: int = 2,
    cps: float = 15.0,
    min_dur: float = 0.8,
    max_dur: float = 3.5,
) -> List[Tuple[float, float, str]]:
    chunks = _compact_segments_from_text(full_text, max_line_chars=max_line_chars, max_lines=max_lines)
    return _assign_times_by_cps(chunks, total_duration=total_duration, cps=cps, min_dur=min_dur, max_dur=max_dur)


from PIL import ImageEnhance  # Додайте для контрасту/яскравості


def rounded_box_clip_3d(
        box_size: Tuple[int, int],
        radius: int = 20,
        *,
        fill: Tuple[int, int, int, int] = (255, 255, 255, 210),  # Базовий fill
        border_color: Tuple[int, int, int, int] = (255, 255, 255, 255),
        border_width: int = 2,
        shadow_offset: Tuple[int, int] = (1, 1),
        shadow_blur: int = 8,
        shadow_opacity: int = 100,
        enable_inner_stroke: bool = True,
        inner_stroke_color: Tuple[int, int, int, int] = (255, 255, 255, 90),
        inner_stroke_width: int = 1,
        enable_alpha_gradient: bool = True,
        alpha_grad_strength: float = 0.14,
        mode: str = "tahoe_glass",  # Додайте "tahoe_glass"
        glass_blur: int = 10,
        glass_tint: Tuple[int, int, int, int] = (255, 255, 255, 120),
        background_image_path: Optional[str] = None,
        # Нові параметри для Tahoe
        enable_refraction: bool = True,  # Симуляція рефракції
        specular_strength: float = 0.4,  # Блиск (0–1)
        adaptive_tint: bool = True,  # Адаптивний колір
):
    # print(f"Mode is: {mode} (Tahoe-like: {mode == 'tahoe_glass'})")
    if ImageClip is None:
        raise RuntimeError("MoviePy is required for rounded_box_clip_3d")

    w, h = box_size
    pad = max(8, shadow_blur * 2 + glass_blur)
    W = w + pad + max(0, shadow_offset[0])
    H = h + pad + max(0, shadow_offset[1])

    x0 = pad // 2
    y0 = pad // 2
    x1 = x0 + w
    y1 = y0 + h

    img = Image.new("RGBA", (W, H), (0, 0, 0, 0))

    # Тінь (м'якша для Tahoe)
    if shadow_blur > 0 and shadow_opacity > 0:
        shadow = Image.new("RGBA", (W, H), (0, 0, 0, 0))
        sd = ImageDraw.Draw(shadow)
        sd.rounded_rectangle(
            [x0 + shadow_offset[0], y0 + shadow_offset[1], x1 + shadow_offset[0], y1 + shadow_offset[1]],
            radius=radius,
            fill=(0, 0, 0, shadow_opacity // 2),  # Легша тінь (40% opacity)
        )
        shadow = shadow.filter(ImageFilter.GaussianBlur(shadow_blur + 2))  # +2px для fluidity
        img = Image.alpha_composite(img, shadow)

    # База: Фон для glass (розширено для Tahoe)
    if mode in ["glass", "tahoe_glass"] and background_image_path is not None:
        bg = Image.open(background_image_path).convert("RGB")
        bg_resized = bg.resize((W, H), Image.LANCZOS)

        if mode == "tahoe_glass":
            # Адаптивний тінт: Обчисліть середній колір фону
            if adaptive_tint:
                avg_color = np.mean(bg_resized, axis=(0, 1)).astype(int)  # Середній RGB
                # Тепліший/холодніший тінт: +20% синього для "скла", opacity 15–25%
                tint_r, tint_g, tint_b = int(avg_color[0] * 0.8 + 255 * 0.2), int(avg_color[1] * 0.8 + 255 * 0.2), int(
                    avg_color[2] * 0.7 + 255 * 0.3)
                glass_tint = (tint_r, tint_g, tint_b, int(255 * 0.18))  # ~18% opacity для Tahoe

            # Розмиття (frosted)
            bg_resized = bg_resized.filter(ImageFilter.GaussianBlur(glass_blur + 5))  # Збільште до 15px

            # Рефракція: Проста симуляція — зсув пікселів (wave distortion)
            if enable_refraction:
                # Створіть маску для distortion (синусоїда для "рідинності")
                dist_mask = Image.new("L", (W, H), 255)
                for y in range(H):
                    for x in range(W):
                        dist = int(2 * np.sin(x * 0.01) * np.sin(y * 0.01))  # Легкий зсув
                        dist_mask.putpixel((x, y), max(0, 255 - abs(dist)))
                bg_resized = Image.composite(bg_resized, bg_resized.offset(1, 1), dist_mask)  # Offset composite

            # Базовий шар
            base = bg_resized.convert("RGBA")
            tint_layer = Image.new("RGBA", (W, H), glass_tint)
            base = Image.alpha_composite(base, tint_layer)

            # Specular highlights: Градієнт блиску на краях (Tahoe-стиль)
            highlight = Image.new("RGBA", (W, H), (0, 0, 0, 0))
            hd = ImageDraw.Draw(highlight)
            # Білий градієнт на верх/краях (squircle для Tahoe)
            for edge in [0, W - 1]:  # Лівий/правий
                for yy in range(H):
                    alpha = int(255 * specular_strength * (1 - (yy / H) ** 2))  # Кулішоподібний градієнт
                    highlight.putpixel((edge, yy), (255, 255, 255, alpha // 3))
            base = Image.alpha_composite(base, highlight.filter(ImageFilter.GaussianBlur(2)))

        else:  # Базовий glass
            base = bg_resized.convert("RGBA").filter(ImageFilter.GaussianBlur(glass_blur))
            tint_layer = Image.new("RGBA", (W, H), glass_tint)
            base = Image.alpha_composite(base, tint_layer)
    else:
        base = Image.new("RGBA", (W, H), (0, 0, 0, 0))
        bd = ImageDraw.Draw(base)
        bd.rounded_rectangle([x0, y0, x1, y1], radius=radius, fill=fill)
        # Градієнт альфи (для fluidity)
        if enable_alpha_gradient and fill[3] > 0:
            grad = Image.new("L", (1, h), 0)
            for i in range(h):
                t = 1.0 - (i / max(1, h - 1))
                val = int(255 * alpha_grad_strength * t * 1.2)  # Посилити для Tahoe
                grad.putpixel((0, i), val)
            grad = grad.resize((w, h))
            grad_rgba = Image.new("RGBA", (w, h), (255, 255, 255, 0))
            grad_rgba.putalpha(grad)
            base.alpha_composite(grad_rgba, dest=(x0, y0))

    img = Image.alpha_composite(img, base)

    # Бордер (тонший, з блиском для Tahoe)
    if border_width > 0:
        stroke = Image.new("RGBA", (W, H), (0, 0, 0, 0))
        sd = ImageDraw.Draw(stroke)
        sd.rounded_rectangle([x0, y0, x1, y1], radius=radius, outline=border_color, width=border_width - 1)  # Тонший
        # Додайте блиск на бордер
        if mode == "tahoe_glass":
            stroke = ImageEnhance.Brightness(stroke).enhance(1.2)  # +20% яскравості
        img = Image.alpha_composite(img, stroke)

    # Inner stroke (як у Tahoe)
    if enable_inner_stroke and inner_stroke_width > 0:
        inner = Image.new("RGBA", (W, H), (0, 0, 0, 0))
        idr = ImageDraw.Draw(inner)
        inset = max(1, border_width) + 1
        idr.rounded_rectangle(
            [x0 + inset, y0 + inset, x1 - inset, y1 - inset],
            radius=max(1, radius - inset),
            outline=inner_stroke_color,
            width=inner_stroke_width,
        )
        img = Image.alpha_composite(img, inner)

    arr = np.array(img)
    rgb = arr[..., :3]
    alpha = arr[..., 3] / 255.0

    clip_rgb = ImageClip(rgb)
    clip_mask = ImageClip(alpha, ismask=True)
    return clip_rgb.set_mask(clip_mask)

def effective_text_bbox(clip, threshold: float = 0.01) -> Tuple[int, int, int, int]:
    if clip.mask is None:
        clip = clip.add_mask()
    frame = clip.mask.get_frame(0)
    cols = np.where(frame.max(axis=0) > threshold)[0]
    rows = np.where(frame.max(axis=1) > threshold)[0]
    if cols.size == 0 or rows.size == 0:
        return 0, 0, 0, 0
    x0, x1 = int(cols[0]), int(cols[-1])
    y0, y1 = int(rows[0]), int(rows[-1])  # ← ДОДАЄМО ПАДІНГ ЗНИЗУ
    w = x1 - x0 + 1
    h = y1 - y0 + 1
    return x0, y0, w, h


def make_layered_textclip(
    text: str,
    *,
    fontsize: int,
    font: str = "Arial-Bold",
    fill_color: str = "white",
    stroke_color: str = "#111111",
    stroke_width: int = 2,
    depth: int = 6,
    depth_opacity: float = 0.10,
    depth_offset: Tuple[int, int] = (1, 1),
    shadow_offset: Tuple[int, int] = (3, 3),
    shadow_opacity: float = 0.25,
    align: str = "West",
    method: str = "caption",
    size: Optional[Tuple[int, int]] = None,
    kerning: Optional[int] = None,
    interline: int = 5,
):
    if TextClip is None or CompositeVideoClip is None:
        raise RuntimeError("MoviePy is required for make_layered_textclip")

    layers = []
    for i in range(1, depth + 1):
        dx = depth_offset[0] * i
        dy = depth_offset[1] * i
        extr = TextClip(
            text,
            fontsize=fontsize,
            color="#000000",
            font=font,
            stroke_color=None,
            stroke_width=0,
            align=align,
            method=method,
            size=size,
            kerning=(kerning if kerning is not None else 0),
            interline=interline
        ).set_opacity(depth_opacity * (1 - (i - 1) / max(1, depth))).set_position((dx, dy))
        layers.append(extr)

    shadow = TextClip(
        text,
        fontsize=fontsize,
        color="#000000",
        font=font,
        stroke_color=None,
        stroke_width=0,
        align=align,
        method=method,
        size=size,
        kerning=(kerning if kerning is not None else 0),
        interline=interline
    ).set_opacity(shadow_opacity).set_position(shadow_offset)
    layers.append(shadow)

    base = TextClip(
        text,
        fontsize=fontsize,
        color=fill_color,
        font=font,
        stroke_color=stroke_color,
        stroke_width=stroke_width,
        align=align,
        method=method,
        size=size,
        kerning=(kerning if kerning is not None else 0),
        interline=interline
    )
    layers.append(base)

    return CompositeVideoClip(layers)


# =============================================================================================
#                                        АГЕНТ
# =============================================================================================

class VideoAgent:
    """Агент/сервіс із наборами інструментів для побудови коротких відео."""

    def __init__(
        self,
        *,
        typography: Typography | None = None,
        layout: Layout | None = None,
        plates: BackgroundPlates | None = None,
        anim: Anim | None = None,
        out_dir: str = "video_out",
    ) -> None:
        self.typography = typography or Typography()
        self.layout = layout or Layout()
        self.plates = plates or BackgroundPlates()
        self.anim = anim or Anim()

        self.out = Path(out_dir)
        self.out.mkdir(parents=True, exist_ok=True)
        # Безпечна перевірка:
        resolved = self.out.resolve()
        if resolved == Path("/") or resolved == Path.home():
            raise ValueError(f"Refusing to clean unsafe directory: {resolved}")

   
    @staticmethod
    def _write_empty_srt(path: str) -> None:
        """Створює порожній .srt файл (без реплік)."""
        with open(path, "w", encoding="utf-8") as f:
            f.write("")

    # ----------------------------- ПУБЛІЧНІ ІНСТРУМЕНТИ -----------------------------

    def generate_subtitles(
        self,
        full_text: str,
        total_duration: float,
        cfg: Subtitles,
    ) -> List[Tuple[float, float, str]]:
        return build_srt_segments_compact(
            full_text=full_text,
            total_duration=total_duration,
            max_line_chars=cfg.max_line_chars,
            max_lines=cfg.max_lines,
            cps=cfg.cps,
            min_dur=cfg.min_dur,
            max_dur=cfg.max_dur,
        )

    def create_into_video(
        self,
        *,
        image_path: str,
        text_blocks: List[str],  # [title, subtitle]
        audio_path: str,
        output_path: str = "output_video.mp4",
        subtitles: Optional[Subtitles] = None,
    ) -> str:
        """
        Головний конвеєр: збирає відео 1280×720 із фоном, заголовком, підзаголовком та аудіо.
        Повертає шлях до створеного відео.
        """
        if ImageClip is None or AudioFileClip is None or CompositeVideoClip is None:
            raise RuntimeError("MoviePy must be installed to create video")

        if len(text_blocks) != 2:
            raise ValueError("text_blocks повинен містити рівно 2 елементи: [title, subtitle]")
        if not os.path.exists(image_path):
            raise FileNotFoundError(f"Зображення не знайдено: {image_path}")
        if not os.path.exists(audio_path):
            raise FileNotFoundError(f"Аудіо не знайдено: {audio_path}")

        # ---- Аудіо ----
        audio = AudioFileClip(audio_path)
        duration = audio.duration

        # ---- SRT: побудова та запис ----
        srt_file_path: Optional[str] = None
        if subtitles and subtitles.text:
            segs = self.generate_subtitles(subtitles.text, total_duration=duration, cfg=subtitles)
            srt_file_path = subtitles.srt_path or os.path.splitext(output_path)[0] + ".srt"
            srt_file_path = self.out / srt_file_path
            _write_srt(segs, srt_file_path)
            print(f"Субтитри збережено: {srt_file_path}")

        # ---- Картинка (crop-to-fit) ----
        video_w, video_h = self.layout.size
        img = ImageClip(image_path).set_duration(duration)
        img_ratio = img.w / img.h
        target_ratio = video_w / video_h
        if img_ratio > target_ratio:
            img = img.resize(height=video_h).crop(x_center=img.w // 2, width=video_w)
        else:
            img = img.resize(width=video_w).crop(y_center=img.h // 2, height=video_h)
        img = img.set_position("center")

        # ---- Текстові кліпи ----
        t = self.typography
        lay = self.layout
        plates = self.plates

        title_clip = make_layered_textclip(
            text_blocks[0],
            fontsize=t.title_size,
            font=t.font,
            fill_color=t.title_color,
            stroke_color=t.stroke_color,
            stroke_width=t.stroke_width,
            depth=6,
            depth_opacity=0.10,
            depth_offset=(1, 1),
            shadow_offset=(3, 4),
            shadow_opacity=0.22,
            align="West",
            method=t.title_method,
        ).set_duration(duration)

        title_x = int(video_w * lay.title_left_margin)
        title_y = int((video_h - title_clip.h) / 2)
        title_clip = title_clip.set_position((title_x, title_y))

        max_width = int(video_w * lay.subtitle_width_ratio)
        subtitle_clip = make_layered_textclip(
            text_blocks[1],
            fontsize=t.subtitle_size,
            font=t.font,
            fill_color=t.subtitle_color,
            stroke_color=t.stroke_color,
            stroke_width=max(2, t.stroke_width - 1),
            depth=5,
            depth_opacity=0.10,
            depth_offset=(1, 1),
            shadow_offset=(3, 4),
            shadow_opacity=0.22,
            align="West",
            method=t.subtitle_method,
            size=(max_width, None),
            kerning=-1,
        ).set_duration(duration)

        subtitle_x = int(video_w * lay.subtitle_left_margin)
        subtitle_y = int(video_h * (1 - lay.subtitle_bottom_margin) - subtitle_clip.h)
        subtitle_clip = subtitle_clip.set_position((subtitle_x, subtitle_y))

        # ---- Плашки під текст ----
        clips = [img]
        if plates.enabled:
            tx0, ty0, tw_eff, th_eff = effective_text_bbox(title_clip)
            sx0, sy0, sw_eff, sh_eff = effective_text_bbox(subtitle_clip)

            title_bg = rounded_box_clip_3d(
                box_size=(tw_eff + plates.padding_left + plates.padding_right,
                          th_eff + plates.padding_top + plates.padding_bottom),
                radius=plates.title_radius,
                fill=(255, 255, 255, int(255 * plates.opacity)),
                border_color=(255, 255, 255, 255),
                border_width=2,
                shadow_offset=(3, 4),
                shadow_blur=8,
                shadow_opacity=80,
                enable_inner_stroke=True,
                alpha_grad_strength=0.16,
                # mode="paper3D",
            ).set_duration(duration).set_position((title_x - plates.padding_left + tx0,
                                                   title_y - plates.padding_top + ty0))

            subtitle_bg = rounded_box_clip_3d(
                box_size=(sw_eff + plates.padding_left + plates.padding_right,
                          sh_eff + plates.padding_top + plates.padding_bottom),
                radius=plates.subtitle_radius,
                fill=(255, 255, 255, int(255 * plates.opacity)),
                border_color=(255, 255, 255, 255),
                border_width=2,
                shadow_offset=(3, 4),
                shadow_blur=8,
                shadow_opacity=80,
                enable_inner_stroke=True,
                alpha_grad_strength=0.12,
                # mode="paper3D",
            ).set_duration(duration).set_position((subtitle_x - plates.padding_left + sx0,
                                                   subtitle_y - plates.padding_top + sy0))

            clips.extend([title_bg, subtitle_bg])

        # ---- Анімація ----
        an = self.anim
        def _fadeout_alpha(clip, dur: float):
            if clip.mask is None:
                clip = clip.add_mask()
            return clip.set_mask(clip.mask.fx(fadeout, dur))

        # Почерговий intro
        title_clip = title_clip.set_start(0).crossfadein(an.intro_fade)
        subtitle_clip = subtitle_clip.set_start(an.intro_fade).crossfadein(an.intro_fade)
        if plates.enabled:
            clips[-2] = clips[-2].set_start(0).crossfadein(an.intro_fade)  # title_bg
            clips[-1] = clips[-1].set_start(an.intro_fade).crossfadein(an.intro_fade)  # subtitle_bg

        outro_start_time = max(0.0, duration + an.outro_start)
        final_end_time = outro_start_time + an.outro_fade

        title_clip = title_clip.set_end(final_end_time)
        subtitle_clip = subtitle_clip.set_end(final_end_time)
        if plates.enabled:
            clips[-2] = clips[-2].set_end(final_end_time)
            clips[-1] = clips[-1].set_end(final_end_time)

        title_clip = _fadeout_alpha(title_clip, an.outro_fade)
        subtitle_clip = _fadeout_alpha(subtitle_clip, an.outro_fade)
        if plates.enabled:
            clips[-2] = _fadeout_alpha(clips[-2], an.outro_fade)
            clips[-1] = _fadeout_alpha(clips[-1], an.outro_fade)

        clips.extend([title_clip, subtitle_clip])

        # ---- Subtitles overlay (опційно) ----
        if subtitles and subtitles.text and subtitles.burn_in:
            try:
                from moviepy.video.tools.subtitles import SubtitlesClip

                def _make(tc: str):
                    return TextClip(tc, fontsize=40, color="white", font=t.font, method="caption")

                segs = self.generate_subtitles(subtitles.text, duration, subtitles)
                overlay = SubtitlesClip([(st, en, tx) for st, en, tx in segs], make_textclip=_make)
                overlay = overlay.set_position(("center", "bottom")).set_duration(duration)
                clips.append(overlay)
            except Exception:
                # Фолбек: ігноруємо burn-in, але .srt вже записано вище
                pass

        video = CompositeVideoClip(clips).set_audio(audio)
        output_path = os.path.join(self.out, output_path)
        print("Починаємо рендеринг відео...")
        video.write_videofile(
            output_path,
            fps=24,
            codec="libx264",
            audio_codec="aac",
            threads=4,
            preset="medium",
            verbose=True,
            logger="bar",
        )
        print(f"Відео успішно збережено: {output_path}")
        return output_path


    def create_outro_video(
        self,
        *,
        image_path: str,
        text_blocks: List[str],  # [title, subtitle]
        output_path: str = "output_video.mp4",
        duration: float = 5.0,
        create_empty_srt: bool = True,
        srt_path: Optional[str] = None,
    ) -> str:
        """
        Рендер «заключного» відео без звуку та без субтитрів.
        - Відео триває рівно `duration` секунд.
        - За потреби створюється порожній .srt файл.
        - Застосовуються ті самі плашки та анімації, що й у звичайному режимі.
        """
        if ImageClip is None or CompositeVideoClip is None:
            raise RuntimeError("MoviePy must be installed to create video")
        if len(text_blocks) != 2:
            raise ValueError("text_blocks повинен містити рівно 2 елементи: [title, subtitle]")
        if not os.path.exists(image_path):
            raise FileNotFoundError(f"Зображення не знайдено: {image_path}")

        # Картинка (crop-to-fit)
        video_w, video_h = self.layout.size
        img = ImageClip(image_path).set_duration(duration)
        img_ratio = img.w / img.h
        target_ratio = video_w / video_h
        if img_ratio > target_ratio:
            img = img.resize(height=video_h).crop(x_center=img.w // 2, width=video_w)
        else:
            img = img.resize(width=video_w).crop(y_center=img.h // 2, height=video_h)
        img = img.set_position("center")

        # Тексти
        t = self.typography
        lay = self.layout
        plates = self.plates

        title_clip = make_layered_textclip(
            text_blocks[0],
            fontsize=t.title_size,
            font=t.font,
            fill_color=t.title_color,
            stroke_color=t.stroke_color,
            stroke_width=t.stroke_width,
            depth=6,
            depth_opacity=0.10,
            depth_offset=(1, 1),
            shadow_offset=(3, 4),
            shadow_opacity=0.22,
            align="West",
            method=t.title_method,
        ).set_duration(duration)

        title_x = int(video_w * lay.title_left_margin)
        title_y = int((video_h - title_clip.h) / 2)
        title_clip = title_clip.set_position((title_x, title_y))

        max_width = int(video_w * lay.subtitle_width_ratio)
        subtitle_clip = make_layered_textclip(
            text_blocks[1],
            fontsize=t.subtitle_size,
            font=t.font,
            fill_color=t.subtitle_color,
            stroke_color=t.stroke_color,
            stroke_width=max(2, t.stroke_width - 1),
            depth=5,
            depth_opacity=0.10,
            depth_offset=(1, 1),
            shadow_offset=(3, 4),
            shadow_opacity=0.22,
            align="West",
            method=t.subtitle_method,
            size=(max_width, None),
            kerning=-1,
        ).set_duration(duration)

        subtitle_x = int(video_w * lay.subtitle_left_margin)
        subtitle_y = int(video_h * (1 - lay.subtitle_bottom_margin) - subtitle_clip.h)
        subtitle_clip = subtitle_clip.set_position((subtitle_x, subtitle_y))

        # Плашки (опційно)
        clips = [img]
        if plates.enabled:
            tx0, ty0, tw_eff, th_eff = effective_text_bbox(title_clip)
            sx0, sy0, sw_eff, sh_eff = effective_text_bbox(subtitle_clip)
            title_bg = rounded_box_clip_3d(
                box_size=(tw_eff + plates.padding_left + plates.padding_right,
                          th_eff + plates.padding_top + plates.padding_bottom),
                radius=plates.title_radius,
                fill=(255, 255, 255, int(255 * plates.opacity)),
                border_color=(255, 255, 255, 255),
                border_width=2,
                shadow_offset=(3, 4),
                shadow_blur=8,
                shadow_opacity=80,
                enable_inner_stroke=True,
                alpha_grad_strength=0.16,
                # mode="paper3D",
            ).set_duration(duration).set_position((title_x - plates.padding_left + tx0,
                                                   title_y - plates.padding_top + ty0))
            subtitle_bg = rounded_box_clip_3d(
                box_size=(sw_eff + plates.padding_left + plates.padding_right,
                          sh_eff + plates.padding_top + plates.padding_bottom),
                radius=plates.subtitle_radius,
                fill=(255, 255, 255, int(255 * plates.opacity)),
                border_color=(255, 255, 255, 255),
                border_width=2,
                shadow_offset=(3, 4),
                shadow_blur=8,
                shadow_opacity=80,
                enable_inner_stroke=True,
                alpha_grad_strength=0.12,
                # mode="paper3D",
            ).set_duration(duration).set_position((subtitle_x - plates.padding_left + sx0,
                                                   subtitle_y - plates.padding_top + sy0))
            clips.extend([title_bg, subtitle_bg])

        # Анімації (intro/outro)
        an = self.anim
        def _fadeout_alpha(clip, dur: float):
            if clip.mask is None:
                clip = clip.add_mask()
            return clip.set_mask(clip.mask.fx(fadeout, dur))
        title_clip = title_clip.set_start(0).crossfadein(an.intro_fade)
        subtitle_clip = subtitle_clip.set_start(an.intro_fade).crossfadein(an.intro_fade)
        if plates.enabled:
            clips[-2] = clips[-2].set_start(0).crossfadein(an.intro_fade)
            clips[-1] = clips[-1].set_start(an.intro_fade).crossfadein(an.intro_fade)
        outro_start_time = max(0.0, duration + an.outro_start)
        final_end_time = outro_start_time + an.outro_fade
        title_clip = title_clip.set_end(final_end_time)
        subtitle_clip = subtitle_clip.set_end(final_end_time)
        if plates.enabled:
            clips[-2] = clips[-2].set_end(final_end_time)
            clips[-1] = clips[-1].set_end(final_end_time)
        title_clip = _fadeout_alpha(title_clip, an.outro_fade)
        subtitle_clip = _fadeout_alpha(subtitle_clip, an.outro_fade)
        if plates.enabled:
            clips[-2] = _fadeout_alpha(clips[-2], an.outro_fade)
            clips[-1] = _fadeout_alpha(clips[-1], an.outro_fade)
        clips.extend([title_clip, subtitle_clip])

        # Порожній SRT
        if create_empty_srt:
            srt_out = srt_path or os.path.splitext(output_path)[0] + ".srt"
            srt_out = self.out / srt_out
            self._write_empty_srt(srt_out)   # <-- було _write_empty_srt(...)
            print(f"Порожній SRT збережено: {srt_out}")

        # Рендер без аудіо
        video = CompositeVideoClip(clips)
        output_path = os.path.join(self.out, output_path)
        print("Починаємо рендеринг відео (без звуку)...")
        video.write_videofile(
            output_path,
            fps=24,
            codec="libx264",
            audio=False,
            threads=4,
            preset="medium",
            verbose=True,
            logger="bar",
        )
        print(f"Відео успішно збережено: {output_path}")
        return output_path


    # ADD: у class VideoShortAgent
    def create_video_bulleted_sequence(
        self,
        *,
        image_path: str,
        title_text: str,
        title_audio_path: str,
        title_subs_text: str | None,
        bullets: list[dict],   # кожен: {"text": str, "audio": str, "subs": Optional[str]}
        output_path: str = "output_bullets.mp4",
        burn_subtitles: bool = False,
        srt_path: str | None = None,
        title_top_margin: int = 32,
    ) -> str:

        """
        Розширений режим:
          1) Title угорі ліворуч (із плашкою).
          2) Усі bullets на ОДНІЙ плашці, з’являються по черзі та залишаються видимими.
          3) Для Title і кожного bullet — окремий mp3 + текст для SRT.
        """
        # --- валідації та імпорти ---
        if ImageClip is None or CompositeVideoClip is None or AudioFileClip is None:
            raise RuntimeError("MoviePy must be installed to create video")
        if not os.path.exists(image_path):
            raise FileNotFoundError(f"Зображення не знайдено: {image_path}")
        if not os.path.exists(title_audio_path):
            raise FileNotFoundError(f"Title аудіо не знайдено: {title_audio_path}")
        for i, b in enumerate(bullets):
            if "text" not in b:
                raise ValueError(f"Немає 'text' для bullet #{i+1}")
            if "audio" not in b or not os.path.exists(b["audio"]):
                raise FileNotFoundError(f"Аудіо для bullet #{i+1} не знайдено: {b.get('audio')}")
        from moviepy.editor import concatenate_audioclips
    
        # --- аудіо: title + bullets у спільну доріжку ---
        audio_clips = [AudioFileClip(title_audio_path)] + [AudioFileClip(b["audio"]) for b in bullets]
        full_audio = concatenate_audioclips(audio_clips)
        total_duration = float(sum(ac.duration for ac in audio_clips))
    
        # --- фон (crop-to-fit) ---
        vw, vh = self.layout.size
        img = ImageClip(image_path).set_duration(total_duration)
        if img.w / img.h > vw / vh:
            img = img.resize(height=vh).crop(x_center=img.w // 2, width=vw)
        else:
            img = img.resize(width=vw).crop(y_center=img.h // 2, height=vh)
        img = img.set_position("center")
        clips = [img]
    
        # утиліта для fadeout по альфі
        def _fadeout_alpha(clip, dur: float):
            if clip.mask is None:
                clip = clip.add_mask()
            return clip.set_mask(clip.mask.fx(fadeout, dur))
    
        t, lay, plates, an = self.typography, self.layout, self.plates, self.anim
    
        # --- TITLE угорі ліворуч + плашка ---
        title_clip = make_layered_textclip(
            title_text,
            fontsize=t.title_size, font=t.font,
            fill_color=t.title_color,
            stroke_color=t.stroke_color, stroke_width=t.stroke_width,
            depth=6, depth_opacity=0.10, depth_offset=(1, 1),
            shadow_offset=(3, 4), shadow_opacity=0.22,
            align="West", method=t.title_method,
        ).set_duration(total_duration)
        title_x = int(vw * lay.title_left_margin)
        title_y = int(title_top_margin)
        title_clip = title_clip.set_position((title_x, title_y))
    
        if plates.enabled:
            tx0, ty0, tw_eff, th_eff = effective_text_bbox(title_clip)
            title_bg = rounded_box_clip_3d(
                box_size=(tw_eff + plates.padding_left + plates.padding_right,
                          th_eff + plates.padding_top + plates.padding_bottom),
                radius=plates.title_radius,
                fill=(255, 255, 255, int(255 * plates.opacity)),
                border_color=(255, 255, 255, 255),
                border_width=2,
                shadow_offset=(3, 4),
                shadow_blur=8,
                shadow_opacity=80,
                enable_inner_stroke=True,
                alpha_grad_strength=0.16,
                # mode="paper3D",
            ).set_duration(total_duration).set_position((
                title_x - plates.padding_left + tx0,
                title_y - plates.padding_top + ty0
            ))
            # анімація плашки Title
            title_bg = title_bg.set_start(0).crossfadein(an.intro_fade)
            title_bg = title_bg.set_end(total_duration + an.outro_start + an.outro_fade)
            title_bg = _fadeout_alpha(title_bg, an.outro_fade)
            clips.append(title_bg)
    
        # анімація Title (цілий ролик)
        title_clip = title_clip.set_start(0).crossfadein(an.intro_fade) \
                               .set_end(total_duration + an.outro_start + an.outro_fade)
        title_clip = _fadeout_alpha(title_clip, an.outro_fade)
        clips.append(title_clip)
    
        # --- SRT: ініціалізація + Title (офсет = 0) ---
        srt_segments: list[tuple[float, float, str]] = []
        if title_subs_text:
            segs = build_srt_segments_compact(
                full_text=title_subs_text, total_duration=audio_clips[0].duration,
                max_line_chars=38, max_lines=2, cps=16.0, min_dur=0.8, max_dur=3.2
            )
            srt_segments.extend(segs)
    
        # --- BULLETS (одна плашка + по черзі, стеком вниз) ---
        cur_t = audio_clips[0].duration  # коли закінчиться Title-аудіо
        max_width = int(vw * lay.subtitle_width_ratio)
        spacing = 20  # відступ між рядками
    
        # 1) Попередньо створюємо кліпи, щоб поміряти висоти/ширину
        bp_clips: list = []
        bp_durations: list[float] = []
        for b, ac in zip(bullets, audio_clips[1:]):
            txt = f"• {b['text']}"
            bp = make_layered_textclip(
                txt,
                fontsize=t.subtitle_size, font=t.font,
                fill_color=t.subtitle_color,
                stroke_color=t.stroke_color, stroke_width=max(2, t.stroke_width - 1),
                depth=5, depth_opacity=0.10, depth_offset=(1, 1),
                shadow_offset=(3, 4), shadow_opacity=0.22,
                align="West", method=t.subtitle_method,
                size=(max_width, None), kerning=-1,
            ).set_duration(ac.duration)
            bp_clips.append(bp)
            bp_durations.append(ac.duration)
    
        # 2) Порахувати розміри стеку
        heights, widths = [], []
        for bp in bp_clips:
            _, _, w_eff, h_eff = effective_text_bbox(bp)
            widths.append(w_eff)
            heights.append(h_eff)
        total_h = sum(heights) + (len(heights) - 1) * spacing if heights else 0
        plate_w_eff = max(widths) if widths else 0
    
        # 3) Позиція спільної плашки (вирівнюємо стек так, щоб низ лишався як раніше)
        # sub_x = int(vw * lay.subtitle_left_margin)
        # stack_top_y = int(vh * (1 - lay.subtitle_bottom_margin) - total_h)
        
        # === НОВІ ПАРАМЕТРИ ВИРІВНЮВАННЯ (можеш винести в аргументи функції) ===
        anchor_h = "left"    # "left" | "center" | "right"
        anchor_v = "center"     # "top"  | "center" | "bottom"
        
        # ДОДАТКОВІ ВІДСТУПИ (за бажанням)
        top_margin_px    =  int(vh * 0.12)                  # коли anchor_v == "top"
        bottom_margin_px =  int(vh * self.layout.subtitle_bottom_margin)  # коли "bottom"
        side_margin_px   =  int(vw * self.layout.subtitle_left_margin)    # для "left"/"right"
        
        # plate_w_eff і total_h вже пораховані вище (ширина/висота єдиної плашки під bullets)
        
        # === HORIZONTAL X ===
        if anchor_h == "left":
            sub_x = side_margin_px
        elif anchor_h == "center":
            sub_x = int((vw - plate_w_eff) / 2)
        elif anchor_h == "right":
            sub_x = int(vw - side_margin_px - plate_w_eff)
        else:
            raise ValueError("anchor_h must be left|center|right")
        
        # === VERTICAL Y (stack_top_y — верхній край «стеку» булетів) ===
        if anchor_v == "top":
            stack_top_y = top_margin_px
        elif anchor_v == "center":
            stack_top_y = int((vh - total_h) / 2)
        elif anchor_v == "bottom":
            stack_top_y = int(vh - bottom_margin_px - total_h)
        else:
            raise ValueError("anchor_v must be top|center|bottom")
    
        # 4) Одна спільна плашка під усі bullets (з fade in/out)
        if plates.enabled and total_h > 0 and plate_w_eff > 0:
            bullet_plate = rounded_box_clip_3d(
                box_size=(plate_w_eff + plates.padding_left + plates.padding_right,
                          total_h + plates.padding_top + plates.padding_bottom),
                radius=plates.subtitle_radius,
                fill=(255, 255, 255, int(255 * plates.opacity)),
                border_color=(255, 255, 255, 255),
                border_width=2,
                shadow_offset=(3, 4),
                shadow_blur=8,
                shadow_opacity=80,
                enable_inner_stroke=True,
                alpha_grad_strength=0.12,
                # mode="paper3D",
            ).set_duration(total_duration).set_position(
                (sub_x - plates.padding_left, stack_top_y - plates.padding_top)
            )
            # Плашка з’являється з першим bullet
            bullet_plate = bullet_plate.set_start(cur_t).crossfadein(an.intro_fade)
            bullet_plate = bullet_plate.set_end(total_duration + an.outro_start + an.outro_fade)
            bullet_plate = _fadeout_alpha(bullet_plate, an.outro_fade)
            clips.append(bullet_plate)
    
        # 5) Розставляємо bullets вертикально та показуємо по черзі
        y_cursor = stack_top_y
        for b, bp, dur in zip(bullets, bp_clips, bp_durations):
            # позиція рядка на спільній плашці
            _, _, _, h_eff = effective_text_bbox(bp)
            bp = bp.set_position((sub_x, y_cursor))
    
            # start_t = cur_t
            # end_t = total_duration + an.outro_start + an.outro_fade  # лишається до фіналу
    
            # bp = bp.set_start(start_t).crossfadein(an.intro_fade).set_end(end_t)
            # clips.append(bp)

            start_t = cur_t
            end_t = total_duration + an.outro_start + an.outro_fade  # лишається до фіналу
            
            # РОЗТЯГУЄМО САМ КЛІП до глобального фіналу, щоб fadeout рахувався по новій duration
            bp = bp.set_start(start_t).crossfadein(an.intro_fade).set_duration(end_t - start_t)
            bp = _fadeout_alpha(bp, an.outro_fade)
            clips.append(bp)            
    
            # SRT для цього bullet (з офсетом)
            if b.get("subs"):
                segs = build_srt_segments_compact(
                    full_text=b["subs"], total_duration=dur,
                    max_line_chars=38, max_lines=2, cps=16.0, min_dur=0.8, max_dur=3.2
                )
                srt_segments.extend([(st + start_t, en + start_t, tx) for st, en, tx in segs])
    
            y_cursor += h_eff + spacing
            cur_t += dur
    
        # --- SRT запис + (опційно) burn-in ---
        if srt_segments:
            srt_out = srt_path or os.path.splitext(output_path)[0] + ".srt"
            srt_out = self.out / srt_out
            _write_srt(srt_segments, srt_out)
            print(f"SRT збережено: {srt_out}")
            if burn_subtitles:
                try:
                    from moviepy.video.tools.subtitles import SubtitlesClip
                    def _make(tc: str):
                        return TextClip(tc, fontsize=40, color="white", font=t.font, method="caption")
                    overlay = SubtitlesClip([(st, en, tx) for st, en, tx in srt_segments], make_textclip=_make)
                    overlay = overlay.set_position(("center", "bottom")).set_duration(total_duration)
                    clips.append(overlay)
                except Exception:
                    pass
    
        # --- фінальна композиція + аудіо ---
        final = CompositeVideoClip(clips).set_audio(full_audio)
        print("Починаємо рендеринг відео з булетами...")
        output_path = os.path.join(self.out, output_path)
        final.write_videofile(
            output_path,
            fps=24,
            codec="libx264",
            audio_codec="aac",
            threads=4,
            preset="medium",
            verbose=True,
            logger="bar",
        )
        print(f"Відео успішно збережено: {output_path}")
        return output_path

    def create_final_video(
            self,
            output_dir: str,
            *,
            video_name: str = "final_video.mp4",
            srt_name: str = "final_video.srt",
            preserve_original_settings: bool = True,
    ) -> Tuple[str, str]:
        """
        Scans ``self.out`` for:
            intro.mp4 + intro.srt
            slide_2.mp4 + slide_2.srt
            ...
            outro.mp4 + outro.srt

        Concatenates videos and merges subtitles → saves to ``output_dir``.
        """
        # --- ЛОКАЛЬНІ ІМПОРТИ (щоб уникнути NameError) ---
        from moviepy.editor import VideoFileClip, concatenate_videoclips
        import pysrt
        from datetime import timedelta

        if VideoFileClip is None:
            raise RuntimeError("MoviePy must be installed and importable.")

        out_dir = Path(output_dir)
        out_dir.mkdir(parents=True, exist_ok=True)

        # ----------------------------------------------------------------- #
        # 1. Gather ordered list of (mp4_path, srt_path) from self.out
        # ----------------------------------------------------------------- #
        all_files = {p.name: p for p in self.out.iterdir() if p.is_file()}

        def _get_pair(base_name: str):
            mp4 = all_files.get(f"{base_name}.mp4")
            srt = all_files.get(f"{base_name}.srt")
            return mp4, srt

        # Build ordered list
        ordered_bases = ["intro"]
        slide_re = re.compile(r"slide_(\d+)\.mp4")
        slides = []
        for fname in all_files:
            m = slide_re.match(fname)
            if m:
                num = int(m.group(1))
                base = fname.replace(".mp4", "")
                slides.append((num, base))
        slides.sort()
        ordered_bases += [base for _, base in slides]
        ordered_bases.append("outro")

        pairs: List[Tuple[Path, Optional[Path]]] = []
        for base in ordered_bases:
            mp4, srt = _get_pair(base)
            if mp4:
                pairs.append((mp4, srt))
            else:
                print(f"Warning: Missing video: {base}.mp4 – skipping.")

        if not pairs:
            raise FileNotFoundError(f"No video segments found in {self.out}")

        # ----------------------------------------------------------------- #
        # 2. Load clips + collect durations
        # ----------------------------------------------------------------- #
        clips: List[VideoFileClip] = []
        durations: List[float] = []
        first_clip: Optional[VideoFileClip] = None

        for mp4_path, _ in pairs:
            clip = VideoFileClip(str(mp4_path))
            clips.append(clip)
            durations.append(clip.duration)
            if first_clip is None:
                first_clip = clip

        # ----------------------------------------------------------------- #
        # 3. Concatenate videos
        # ----------------------------------------------------------------- #
        final_clip = concatenate_videoclips(clips, method="compose")

        # ----------------------------------------------------------------- #
        # 4. Write video to output_dir
        # ----------------------------------------------------------------- #
        final_video_path = out_dir / video_name
        write_kwargs: dict = {
            "codec": "libx264",
            "audio_codec": "aac",
            "threads": 4,
            "preset": "medium",
            "logger": "bar",
        }

        if preserve_original_settings and first_clip:
            write_kwargs["fps"] = first_clip.fps

            vinfo = first_clip.reader.infos
            if "video_codec" in vinfo:
                write_kwargs["codec"] = vinfo["video_codec"]
            if "video_bitrate" in vinfo:
                write_kwargs["bitrate"] = str(vinfo["video_bitrate"])

            if first_clip.audio:
                ainfo = first_clip.audio.reader.infos
                if "audio_codec" in ainfo:
                    write_kwargs["audio_codec"] = ainfo["audio_codec"]
                if "audio_bitrate" in ainfo:
                    write_kwargs["audio_bitrate"] = str(ainfo["audio_bitrate"])

            # Try stream copy if all clips match
            if all(c.reader.infos.get("video_codec") == write_kwargs["codec"] for c in clips):
                write_kwargs["codec"] = "copy"
            if "audio_codec" in write_kwargs and all(
                    c.audio and c.audio.reader.infos.get("audio_codec") == write_kwargs["audio_codec"]
                    for c in clips
            ):
                write_kwargs["audio_codec"] = "copy"

        print(f"Rendering final video → {final_video_path}")
        final_clip.write_videofile(str(final_video_path), **write_kwargs)

        # ----------------------------------------------------------------- #
        # 5. Merge SRT files + SAVE WITH UTF-8 BOM
        # ----------------------------------------------------------------- #
        combined_subs: List[pysrt.SubRipItem] = []
        offset_ms = 0

        for (mp4_path, srt_path), dur in zip(pairs, durations):
            if srt_path and srt_path.exists():
                subs = pysrt.open(str(srt_path), encoding="utf-8")  # примусово UTF-8
                for sub in subs:
                    start_ms = sub.start.ordinal + offset_ms
                    end_ms = sub.end.ordinal + offset_ms
                    new_sub = pysrt.SubRipItem(
                        index=len(combined_subs) + 1,
                        start=pysrt.SubRipTime.from_ordinal(start_ms),
                        end=pysrt.SubRipTime.from_ordinal(end_ms),
                        text=sub.text,
                    )
                    combined_subs.append(new_sub)
            offset_ms += int(dur * 1000)

        # === ЗАПИС З BOM ===
        final_srt_path = out_dir / srt_name
        with final_srt_path.open("w", encoding="utf-8-sig", newline="\n") as f:
            for i, sub in enumerate(combined_subs, 1):
                f.write(f"{i}\n")
                f.write(f"{sub.start} --> {sub.end}\n")
                f.write(f"{sub.text}\n\n")

        print(f"Subtitles saved (UTF-8 with BOM) → {final_srt_path}")
        print(f"Subtitles saved → {final_srt_path}")

        # Cleanup
        for c in clips:
            c.close()
        final_clip.close()

        return str(final_video_path), str(final_srt_path)

# =============================================================================================
#                                         CLI
# =============================================================================================
if __name__ == "__main__":
    import typer

    app = typer.Typer(add_completion=False, help="VideoShortAgent — CLI для рендеру коротких відео")

    @app.command("create")
    def cli_create(
        image: str = typer.Option(..., help="Шлях до фонового зображення"),
        audio: str = typer.Option(..., help="Шлях до аудіо (mp3/wav)"),
        title: str = typer.Option(..., help="Заголовок"),
        subtitle: str = typer.Option(..., help="Підзаголовок"),
        transcript: Optional[str] = typer.Option(None, help="Повний текст для субтитрів"),
        output: str = typer.Option("output_video.mp4", help="Куди зберегти відео"),
        burn_subs: bool = typer.Option(False, help="Підпалити субтитри в відео (burn-in)"),
    ):
        agent = VideoAgent()
        subs = Subtitles(text=transcript, burn_in=burn_subs)
        agent.create_intro_video(
            image_path=image,
            text_blocks=[title, subtitle],
            audio_path=audio,
            output_path=output,
            subtitles=subs,
        )

    app()
