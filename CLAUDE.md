# iPod Music Extractor

## Project Overview
This project extracts music files from an old iPod Touch (4th gen, iOS 6) connected via USB to a Windows machine. It uses the iTunes COM API to access the iPod as a source, copies the audio files to a local folder, and rewrites full metadata into each file using mutagen.

## Scripts
- **ipod_extractor.py** — copies files into an `Artist / Album / Song` folder structure
- **ipod_extractor_flat.py** — dumps all files into a single flat folder, named `Artist - Album - Song.ext`

Both scripts output to `~/Desktop/iPod Music/` by default (configurable via `OUTPUT_FOLDER` at the top of each file).

## Dependencies
```
pip install pywin32 mutagen
```

## How It Works
1. Connects to a running iTunes instance via `win32com.client`
2. Finds the iPod as a source (source kind = 2)
3. Iterates all playlists on the iPod, deduplicating tracks by persistent ID
4. Copies each audio file with `shutil.copy2`
5. Rewrites metadata into the copied file using mutagen:
   - `.mp3` → ID3 tags
   - `.m4a` / `.m4p` → MP4 tags
   - `.flac` → FLAC Vorbis comments
   - Cover art is extracted from iTunes artwork cache and embedded

## Important Notes
- iTunes must be open and the iPod must be visible in the iTunes sidebar before running
- Purchased tracks (`.m4p`) will copy fine but are DRM-locked — they'll only play in iTunes on an authorized machine
- The scripts are safe to re-run; already-copied files are skipped
- Tracks with missing artist/album metadata fall back to `Unknown Artist` / `Unknown Album`
