"""
iPod Music Extractor - Flat Version
-------------------------------------
Copies all songs from a connected iPod into a single flat folder.
Files are named: Artist - Album - Song.ext

Metadata and cover art already embedded in each file are preserved as-is.

Requirements:
    pip install pymobiledevice3 mutagen

Usage:
    1. Plug in iPod via USB
    2. Open iTunes and make sure the iPod shows up
    3. Run: python ipod_extractor_flat.py
"""

import asyncio
import os
import re
import shutil
import sys
import tempfile
from pathlib import Path

try:
    from mutagen import File as MutagenFile
except ImportError:
    print("ERROR: mutagen is not installed.")
    print("Run: pip install mutagen")
    sys.exit(1)

# ── Configuration ──────────────────────────────────────────────────────────────
OUTPUT_FOLDER = os.path.join(os.path.expanduser("~"), "Desktop", "iPod Music")
# ──────────────────────────────────────────────────────────────────────────────


def sanitize(name: str) -> str:
    if not name or not name.strip():
        return "Unknown"
    return re.sub(r'[\\/*?:"<>|]', "_", name).strip()


def get_tag(audio, *keys):
    for key in keys:
        v = audio.get(key)
        if v:
            val = v[0] if isinstance(v, list) else v
            return str(val).strip()
    return ""


async def main():
    try:
        from pymobiledevice3.lockdown import create_using_usbmux
        from pymobiledevice3.services.afc import AfcService
    except ImportError:
        print("ERROR: pymobiledevice3 is not installed.")
        print("Run: pip install pymobiledevice3")
        return

    print("iPod Music Extractor (Flat)")
    print("=" * 60)

    print("\nConnecting to device...")
    lockdown = await create_using_usbmux()
    print(f"Connected: iOS {lockdown.product_version}\n")

    afc = AfcService(lockdown)

    print("Scanning device for audio files...")
    music_files = []
    base = "/iTunes_Control/Music"
    for folder in await afc.listdir(base):
        if not folder.startswith("F"):
            continue
        path = f"{base}/{folder}"
        for fname in await afc.listdir(path):
            if fname.lower().endswith((".mp3", ".m4a", ".m4p", ".aac", ".flac")):
                music_files.append(f"{path}/{fname}")

    total = len(music_files)
    print(f"Found {total} audio files on device.\n")

    if total == 0:
        print("No tracks found. Exiting.")
        return

    print(f"Copying to: {OUTPUT_FOLDER}\n")
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    ok = 0
    failed = []  # (device_path, reason)

    for i, device_path in enumerate(music_files, 1):
        ext = os.path.splitext(device_path)[1]
        with tempfile.NamedTemporaryFile(suffix=ext, delete=False) as tmp:
            tmp_path = tmp.name
        try:
            data = await afc.get_file_contents(device_path)
            with open(tmp_path, "wb") as dst:
                dst.write(data)

            audio = MutagenFile(tmp_path, easy=True)
            if audio is None:
                reason = "unreadable by mutagen"
                print(f"[{i}/{total}] SKIP: {device_path} — {reason}")
                failed.append((device_path, reason))
                continue

            artist = sanitize(get_tag(audio, "artist", "TPE1") or "Unknown Artist")
            album  = sanitize(get_tag(audio, "album",  "TALB") or "Unknown Album")
            title  = sanitize(get_tag(audio, "title",  "TIT2") or Path(device_path).stem)

            base_name = f"{artist} - {album} - {title}"
            dest_path = os.path.join(OUTPUT_FOLDER, f"{base_name}{ext}")
            counter = 1
            while os.path.isfile(dest_path):
                dest_path = os.path.join(OUTPUT_FOLDER, f"{base_name} ({counter}){ext}")
                counter += 1
            filename = os.path.basename(dest_path)

            shutil.move(tmp_path, dest_path)
            print(f"[{i}/{total}] OK: {filename}")
            ok += 1
        except Exception as e:
            print(f"[{i}/{total}] FAIL: {device_path} — {e}")
            failed.append((device_path, str(e)))
        finally:
            if os.path.exists(tmp_path):
                os.unlink(tmp_path)

    print("\n" + "=" * 60)
    print("Done!")
    print(f"  Copied:  {ok}")
    print(f"  Failed:  {len(failed)}")

    if failed:
        print("\n── Failed ───────────────────────────────────────────────")
        for name, reason in failed:
            print(f"  {name}\n    reason: {reason}")

    print(f"\nFiles saved to:\n  {OUTPUT_FOLDER}")


if __name__ == "__main__":
    asyncio.run(main())
