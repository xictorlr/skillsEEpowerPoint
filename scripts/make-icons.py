"""Genera iconos PNG sólidos para el manifest usando solo la stdlib."""
import pathlib
import struct
import zlib

COLOR = (0x1F, 0x6F, 0xEB)  # azul acento del panel
SIZES = (16, 32, 64, 80)


def chunk(tag: bytes, data: bytes) -> bytes:
    return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)


def png(size: int, rgb: tuple[int, int, int]) -> bytes:
    row = b"\x00" + bytes(rgb) * size
    raw = row * size
    ihdr = struct.pack(">IIBBBBB", size, size, 8, 2, 0, 0, 0)
    return b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", ihdr) + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b"")


def main() -> None:
    out = pathlib.Path(__file__).resolve().parent.parent / "public" / "assets"
    out.mkdir(parents=True, exist_ok=True)
    for size in SIZES:
        (out / f"icon-{size}.png").write_bytes(png(size, COLOR))
        print(f"icon-{size}.png")


if __name__ == "__main__":
    main()
