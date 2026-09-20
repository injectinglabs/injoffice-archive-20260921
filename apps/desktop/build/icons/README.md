# InjOffice desktop icons

These icons are format conversions of the canonical repository-root `logo.png`
(32 × 32 pixel artwork). Nearest-neighbor resizing preserves its colors, alpha,
and pixel edges; no artwork, padding, or background has been added.

- `icon.png`: 1024 × 1024, for Linux packaging.
- `icon.icns`: macOS 16–1024 pixel representations.
- `icon.ico`: Windows 16, 24, 32, 48, 64, 128, and 256 pixel representations.

Regenerate on macOS with Python 3, Pillow, and Apple's `iconutil`:

```sh
python3 apps/desktop/build/icons/generate.py
```

The script reads the original logo without modifying it.
