"""
Pulls all music off iPod Touch via AFC protocol (no Apple ID needed).
Reads embedded tags from each file to rename/organize them.
"""
import asyncio
import os, re, shutil, tempfile
from pathlib import Path
from mutagen import File as MutagenFile

OUTPUT_FOLDER = os.path.join(os.path.expanduser("~"), "Desktop", "iPod Music")

def sanitize(name):
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
        print("Run: pip install pymobiledevice3")
        return

    print("Connecting to device...")
    lockdown = await create_using_usbmux()
    print(f"Connected: {lockdown.product_type}, iOS {lockdown.product_version}")

    afc = AfcService(lockdown)

    # Collect all music files from iTunes_Control/Music/F*/
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
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    ok, skip, fail = 0, 0, 0
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
                print(f"[{i}/{total}] SKIP (unreadable): {device_path}")
                fail += 1
                continue

            artist = sanitize(get_tag(audio, "artist", "TPE1") or "Unknown Artist")
            album  = sanitize(get_tag(audio, "album",  "TALB") or "Unknown Album")
            title  = sanitize(get_tag(audio, "title",  "TIT2") or Path(device_path).stem)

            dest_dir = os.path.join(OUTPUT_FOLDER, artist, album)
            os.makedirs(dest_dir, exist_ok=True)
            dest_path = os.path.join(dest_dir, f"{title}{ext}")

            if os.path.isfile(dest_path):
                print(f"[{i}/{total}] SKIP (exists): {artist} / {album} / {title}{ext}")
                skip += 1
                continue

            shutil.move(tmp_path, dest_path)
            print(f"[{i}/{total}] OK: {artist} / {album} / {title}{ext}")
            ok += 1
        except Exception as e:
            print(f"[{i}/{total}] FAIL: {device_path} — {e}")
            fail += 1
        finally:
            if os.path.exists(tmp_path):
                os.unlink(tmp_path)

    print(f"\nDone. Copied: {ok}  Skipped: {skip}  Failed: {fail}")
    print(f"Files saved to: {OUTPUT_FOLDER}")

if __name__ == "__main__":
    asyncio.run(main())
