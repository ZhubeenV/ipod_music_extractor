"""
iPod Music Extractor (with full metadata)
------------------------------------------
Copies all songs from a connected iPod to your computer
organized as: OUTPUT_FOLDER / Artist / Album / Song

Metadata written to each file:
    Title, Artist, Album, Album Artist, Track Number,
    Disc Number, Year, Genre, Composer, Comment, Cover Art

Requirements:
    pip install pywin32 mutagen

Usage:
    1. Plug in iPod via USB
    2. Open iTunes and make sure the iPod shows up
    3. Run: python ipod_extractor.py
"""

import os
import re
import shutil
import sys

try:
    import win32com.client
except ImportError:
    print("ERROR: pywin32 is not installed.")
    print("Run: pip install pywin32")
    sys.exit(1)

try:
    from mutagen.mp3 import MP3
    from mutagen.id3 import (
        ID3, ID3NoHeaderError,
        TIT2, TPE1, TALB, TPE2, TRCK, TPOS,
        TDRC, TCON, TCOM, COMM, APIC
    )
    from mutagen.mp4 import MP4, MP4Cover
    from mutagen.flac import FLAC, Picture
except ImportError:
    print("ERROR: mutagen is not installed.")
    print("Run: pip install mutagen")
    sys.exit(1)

# ── Configuration ──────────────────────────────────────────────────────────────
OUTPUT_FOLDER = os.path.join(os.path.expanduser("~"), "Desktop", "iPod Music")
# ──────────────────────────────────────────────────────────────────────────────


def sanitize(name: str) -> str:
    """Remove characters that are illegal in Windows file/folder names."""
    if not name or not name.strip():
        return "Unknown"
    name = re.sub(r'[\\/*?:"<>|]', "_", name)
    return name.strip()


def safe_get(track, attr, default=""):
    """Safely read an attribute from an iTunes track COM object."""
    try:
        val = getattr(track, attr)
        return val if val is not None else default
    except Exception:
        return default


def get_artwork_bytes(track):
    """Try to extract cover art bytes from iTunes track. Returns (bytes, format) or (None, None)."""
    try:
        artworks = track.Artwork
        if artworks.Count > 0:
            artwork = artworks.Item(1)
            import tempfile
            tmp = tempfile.NamedTemporaryFile(suffix=".jpg", delete=False)
            tmp.close()
            artwork.SaveArtworkToFile(tmp.name)
            with open(tmp.name, "rb") as f:
                data = f.read()
            os.unlink(tmp.name)
            fmt = "png" if data[:8] == b'\x89PNG\r\n\x1a\n' else "jpeg"
            return data, fmt
    except Exception:
        pass
    return None, None


def write_metadata_mp3(path: str, track, artwork_bytes, artwork_fmt):
    """Write metadata into an MP3 file using ID3 tags."""
    try:
        audio = ID3(path)
    except ID3NoHeaderError:
        audio = ID3()

    def s(attr):
        return str(safe_get(track, attr, ""))

    audio["TIT2"] = TIT2(encoding=3, text=s("Name"))
    audio["TPE1"] = TPE1(encoding=3, text=s("Artist"))
    audio["TALB"] = TALB(encoding=3, text=s("Album"))
    audio["TPE2"] = TPE2(encoding=3, text=s("AlbumArtist") or s("Artist"))
    audio["TCON"] = TCON(encoding=3, text=s("Genre"))
    audio["TCOM"] = TCOM(encoding=3, text=s("Composer"))

    year = safe_get(track, "Year", 0)
    if year:
        audio["TDRC"] = TDRC(encoding=3, text=str(year))

    track_num = safe_get(track, "TrackNumber", 0)
    track_cnt = safe_get(track, "TrackCount", 0)
    if track_num:
        audio["TRCK"] = TRCK(encoding=3, text=f"{track_num}/{track_cnt}" if track_cnt else str(track_num))

    disc_num = safe_get(track, "DiscNumber", 0)
    disc_cnt = safe_get(track, "DiscCount", 0)
    if disc_num:
        audio["TPOS"] = TPOS(encoding=3, text=f"{disc_num}/{disc_cnt}" if disc_cnt else str(disc_num))

    comment = safe_get(track, "Comment", "")
    if comment:
        audio["COMM"] = COMM(encoding=3, lang="eng", desc="", text=comment)

    if artwork_bytes:
        mime = "image/png" if artwork_fmt == "png" else "image/jpeg"
        audio["APIC"] = APIC(encoding=3, mime=mime, type=3, desc="Cover", data=artwork_bytes)

    audio.save(path)


def write_metadata_m4a(path: str, track, artwork_bytes, artwork_fmt):
    """Write metadata into an M4A/M4P/AAC file using MP4 tags."""
    try:
        audio = MP4(path)
    except Exception:
        return

    def s(attr):
        return str(safe_get(track, attr, ""))

    audio["\xa9nam"] = [s("Name")]
    audio["\xa9ART"] = [s("Artist")]
    audio["\xa9alb"] = [s("Album")]
    audio["aART"]    = [s("AlbumArtist") or s("Artist")]
    audio["\xa9gen"] = [s("Genre")]
    audio["\xa9wrt"] = [s("Composer")]

    year = safe_get(track, "Year", 0)
    if year:
        audio["\xa9day"] = [str(year)]

    track_num = safe_get(track, "TrackNumber", 0)
    track_cnt = safe_get(track, "TrackCount", 0)
    if track_num:
        audio["trkn"] = [(int(track_num), int(track_cnt) if track_cnt else 0)]

    disc_num = safe_get(track, "DiscNumber", 0)
    disc_cnt = safe_get(track, "DiscCount", 0)
    if disc_num:
        audio["disk"] = [(int(disc_num), int(disc_cnt) if disc_cnt else 0)]

    comment = safe_get(track, "Comment", "")
    if comment:
        audio["\xa9cmt"] = [comment]

    if artwork_bytes:
        fmt = MP4Cover.FORMAT_PNG if artwork_fmt == "png" else MP4Cover.FORMAT_JPEG
        audio["covr"] = [MP4Cover(artwork_bytes, imageformat=fmt)]

    audio.save()


def write_metadata_flac(path: str, track, artwork_bytes, artwork_fmt):
    """Write metadata into a FLAC file."""
    try:
        audio = FLAC(path)
    except Exception:
        return

    def s(attr):
        return str(safe_get(track, attr, ""))

    audio["title"]       = s("Name")
    audio["artist"]      = s("Artist")
    audio["album"]       = s("Album")
    audio["albumartist"] = s("AlbumArtist") or s("Artist")
    audio["genre"]       = s("Genre")
    audio["composer"]    = s("Composer")

    year = safe_get(track, "Year", 0)
    if year:
        audio["date"] = str(year)

    track_num = safe_get(track, "TrackNumber", 0)
    if track_num:
        audio["tracknumber"] = str(track_num)

    disc_num = safe_get(track, "DiscNumber", 0)
    if disc_num:
        audio["discnumber"] = str(disc_num)

    comment = safe_get(track, "Comment", "")
    if comment:
        audio["comment"] = comment

    if artwork_bytes:
        pic = Picture()
        pic.type = 3  # Front cover
        pic.mime = "image/png" if artwork_fmt == "png" else "image/jpeg"
        pic.data = artwork_bytes
        audio.clear_pictures()
        audio.add_picture(pic)

    audio.save()


def write_metadata(path: str, track):
    """Dispatch metadata writing based on file extension."""
    ext = os.path.splitext(path)[1].lower()
    artwork_bytes, artwork_fmt = get_artwork_bytes(track)

    if ext == ".mp3":
        write_metadata_mp3(path, track, artwork_bytes, artwork_fmt)
    elif ext in (".m4a", ".m4p", ".aac"):
        write_metadata_m4a(path, track, artwork_bytes, artwork_fmt)
    elif ext == ".flac":
        write_metadata_flac(path, track, artwork_bytes, artwork_fmt)
    # Other formats (wav, aiff) are copied as-is with no tag rewriting


def find_ipod_source(itunes):
    """Return the first iPod source found in iTunes, or None."""
    sources = itunes.Sources
    for i in range(1, sources.Count + 1):
        source = sources.Item(i)
        if source.Kind == 2:  # 2 = iPod
            return source
    return None


def get_all_tracks(ipod_source):
    """Collect all tracks from all playlists on the iPod, deduplicated by persistent ID."""
    seen = set()
    tracks = []

    libraries = ipod_source.Playlists
    for i in range(1, libraries.Count + 1):
        playlist = libraries.Item(i)
        try:
            playlist_tracks = playlist.Tracks
        except Exception:
            continue

        for j in range(1, playlist_tracks.Count + 1):
            try:
                track = playlist_tracks.Item(j)
                pid = track.PersistentID
                if pid not in seen:
                    seen.add(pid)
                    tracks.append(track)
            except Exception:
                continue

    return tracks


def copy_track(track, output_root: str) -> str:
    """
    Copy a single track to output_root/Artist/Album/Song.ext
    and write full metadata into the destination file.
    """
    try:
        src_path = track.Location
    except Exception:
        return "[SKIP] Could not get file location"

    if not src_path or not os.path.isfile(src_path):
        return f"[SKIP] File not found on disk: {src_path}"

    artist   = sanitize(safe_get(track, "Artist") or "Unknown Artist")
    album    = sanitize(safe_get(track, "Album")  or "Unknown Album")
    name     = sanitize(safe_get(track, "Name")   or "Unknown Track")
    ext      = os.path.splitext(src_path)[1]
    filename = f"{name}{ext}"

    dest_dir  = os.path.join(output_root, artist, album)
    os.makedirs(dest_dir, exist_ok=True)
    dest_path = os.path.join(dest_dir, filename)

    if os.path.isfile(dest_path):
        return f"[SKIP] Already exists: {artist} / {album} / {filename}"

    shutil.copy2(src_path, dest_path)

    try:
        write_metadata(dest_path, track)
        meta_status = "metadata OK"
    except Exception as e:
        meta_status = f"metadata failed ({e})"

    return f"[OK] {artist} / {album} / {filename} ({meta_status})"


def main():
    print("iPod Music Extractor")
    print("=" * 60)

    print("\nConnecting to iTunes...")
    try:
        itunes = win32com.client.Dispatch("iTunes.Application")
    except Exception as e:
        print(f"ERROR: Could not connect to iTunes: {e}")
        print("Make sure iTunes is installed and open.")
        sys.exit(1)

    print("Looking for iPod...")
    ipod = find_ipod_source(itunes)
    if ipod is None:
        print("ERROR: No iPod found in iTunes.")
        print("Make sure the iPod is plugged in and shows up in iTunes.")
        sys.exit(1)

    print(f"Found iPod: {ipod.Name}")

    print("Reading track list...")
    tracks = get_all_tracks(ipod)
    total  = len(tracks)
    print(f"Found {total} unique tracks.\n")

    if total == 0:
        print("No tracks found. Exiting.")
        sys.exit(0)

    print(f"Copying to: {OUTPUT_FOLDER}\n")
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    ok_count   = 0
    skip_count = 0
    fail_count = 0

    for idx, track in enumerate(tracks, 1):
        track_name = safe_get(track, "Name") or "Unknown"
        print(f"[{idx}/{total}] {track_name}")
        result = copy_track(track, OUTPUT_FOLDER)
        print(f"  {result}")

        if result.startswith("[OK]"):
            ok_count += 1
        elif result.startswith("[SKIP]"):
            skip_count += 1
        else:
            fail_count += 1

    print("\n" + "=" * 60)
    print("Done!")
    print(f"  Copied:  {ok_count}")
    print(f"  Skipped: {skip_count}")
    print(f"  Failed:  {fail_count}")
    print(f"\nFiles saved to:\n  {OUTPUT_FOLDER}")


if __name__ == "__main__":
    main()
