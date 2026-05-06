from __future__ import annotations

import math
import struct
import zlib
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
ASSETS_DIR = ROOT / "assets"
PNG_PATH = ASSETS_DIR / "sound_manager.png"
ICO_PATH = ASSETS_DIR / "sound_manager.ico"


def clamp(value: float) -> int:
    return max(0, min(255, round(value)))


def blend(top: tuple[int, int, int, int], bottom: tuple[int, int, int, int]) -> tuple[int, int, int, int]:
    alpha = top[3] / 255
    inv = 1 - alpha
    return (
        clamp(top[0] * alpha + bottom[0] * inv),
        clamp(top[1] * alpha + bottom[1] * inv),
        clamp(top[2] * alpha + bottom[2] * inv),
        255,
    )


def rounded_rect_alpha(x: float, y: float, size: int, radius: float) -> float:
    cx = min(max(x, radius), size - radius)
    cy = min(max(y, radius), size - radius)
    distance = math.hypot(x - cx, y - cy)
    edge = radius - distance
    if edge >= 1:
        return 1
    if edge <= -1:
        return 0
    return (edge + 1) / 2


def draw_icon(size: int) -> bytes:
    pixels: list[tuple[int, int, int, int]] = []
    radius = size * 0.22

    for y in range(size):
        for x in range(size):
            alpha = rounded_rect_alpha(x + 0.5, y + 0.5, size, radius)
            if alpha <= 0:
                pixels.append((0, 0, 0, 0))
                continue

            gx = x / max(1, size - 1)
            gy = y / max(1, size - 1)
            base = (
                clamp(14 + 20 * gx),
                clamp(18 + 44 * gy),
                clamp(26 + 38 * gx),
                clamp(255 * alpha),
            )

            wave = math.sin((gx * 3.2 + gy * 1.4) * math.pi)
            glow = max(0, wave) * 54
            color = (
                clamp(base[0] + glow * 0.2),
                clamp(base[1] + glow * 1.8),
                clamp(base[2] + glow * 1.5),
                base[3],
            )
            pixels.append(color)

    def set_px(x: int, y: int, color: tuple[int, int, int, int]) -> None:
        if 0 <= x < size and 0 <= y < size:
            index = y * size + x
            pixels[index] = blend(color, pixels[index])

    def line(x1: float, y1: float, x2: float, y2: float, width: float, color: tuple[int, int, int, int]) -> None:
        min_x = max(0, int(min(x1, x2) - width))
        max_x = min(size - 1, int(max(x1, x2) + width))
        min_y = max(0, int(min(y1, y2) - width))
        max_y = min(size - 1, int(max(y1, y2) + width))
        dx = x2 - x1
        dy = y2 - y1
        length_sq = dx * dx + dy * dy
        for yy in range(min_y, max_y + 1):
            for xx in range(min_x, max_x + 1):
                if length_sq == 0:
                    distance = math.hypot(xx - x1, yy - y1)
                else:
                    t = max(0, min(1, ((xx - x1) * dx + (yy - y1) * dy) / length_sq))
                    px = x1 + t * dx
                    py = y1 + t * dy
                    distance = math.hypot(xx - px, yy - py)
                if distance <= width:
                    fade = max(0, min(1, width - distance))
                    set_px(xx, yy, (color[0], color[1], color[2], clamp(color[3] * fade)))

    def circle(cx: float, cy: float, radius: float, color: tuple[int, int, int, int]) -> None:
        for yy in range(int(cy - radius - 2), int(cy + radius + 3)):
            for xx in range(int(cx - radius - 2), int(cx + radius + 3)):
                distance = math.hypot(xx + 0.5 - cx, yy + 0.5 - cy)
                if distance <= radius:
                    fade = max(0, min(1, radius - distance + 0.5))
                    set_px(xx, yy, (color[0], color[1], color[2], clamp(color[3] * fade)))

    white = (242, 250, 252, 240)
    aqua = (101, 228, 209, 255)
    coral = (255, 138, 120, 245)
    line(size * 0.25, size * 0.34, size * 0.75, size * 0.34, size * 0.026, white)
    line(size * 0.25, size * 0.50, size * 0.75, size * 0.50, size * 0.026, white)
    line(size * 0.25, size * 0.66, size * 0.75, size * 0.66, size * 0.026, white)
    circle(size * 0.42, size * 0.34, size * 0.058, aqua)
    circle(size * 0.61, size * 0.50, size * 0.058, coral)
    circle(size * 0.36, size * 0.66, size * 0.058, aqua)

    return write_png(size, size, pixels)


def write_png(width: int, height: int, pixels: list[tuple[int, int, int, int]]) -> bytes:
    raw = bytearray()
    for y in range(height):
        raw.append(0)
        for x in range(width):
            raw.extend(pixels[y * width + x])

    def chunk(name: bytes, data: bytes) -> bytes:
        return (
            struct.pack(">I", len(data))
            + name
            + data
            + struct.pack(">I", zlib.crc32(name + data) & 0xFFFFFFFF)
        )

    return (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 6, 0, 0, 0))
        + chunk(b"IDAT", zlib.compress(bytes(raw), 9))
        + chunk(b"IEND", b"")
    )


def write_ico(images: list[tuple[int, bytes]]) -> bytes:
    header = struct.pack("<HHH", 0, 1, len(images))
    entries = bytearray()
    offset = 6 + 16 * len(images)
    payload = bytearray()
    for size, data in images:
        width = 0 if size == 256 else size
        entries.extend(struct.pack("<BBBBHHII", width, width, 0, 0, 1, 32, len(data), offset))
        payload.extend(data)
        offset += len(data)
    return header + bytes(entries) + bytes(payload)


def main() -> None:
    ASSETS_DIR.mkdir(exist_ok=True)
    png_256 = draw_icon(256)
    PNG_PATH.write_bytes(png_256)
    images = [(16, draw_icon(16)), (32, draw_icon(32)), (48, draw_icon(48)), (64, draw_icon(64)), (128, draw_icon(128)), (256, png_256)]
    ICO_PATH.write_bytes(write_ico(images))
    print(f"Wrote {ICO_PATH}")


if __name__ == "__main__":
    main()
