"""Convert the canonical repository logo to desktop icons without smoothing.

Run on macOS with Python 3, Pillow, and the system iconutil utility installed.
"""

from pathlib import Path
import subprocess
import tempfile

from PIL import Image


destination = Path(__file__).resolve().parent
source = destination.parents[3] / "logo.png"
logo = Image.open(source).convert("RGBA")


def resize(size):
    return logo.resize((size, size), Image.Resampling.NEAREST)


resize(1024).save(destination / "icon.png")

# Supply every frame explicitly so Pillow never applies its default smoothing.
sizes = [16, 24, 32, 48, 64, 128, 256]
resize(256).save(
    destination / "icon.ico",
    format="ICO",
    sizes=[(size, size) for size in sizes],
    append_images=[resize(size) for size in sizes],
)

with tempfile.TemporaryDirectory(prefix="injoffice-icons-") as temporary:
    iconset = Path(temporary) / "InjOffice.iconset"
    iconset.mkdir()
    for size in [16, 32, 128, 256, 512]:
        resize(size).save(iconset / f"icon_{size}x{size}.png")
        resize(size * 2).save(iconset / f"icon_{size}x{size}@2x.png")
    subprocess.run(
        ["iconutil", "--convert", "icns", "--output", str(destination / "icon.icns"), str(iconset)],
        check=True,
    )

print(f"Generated desktop icons from {source}")
