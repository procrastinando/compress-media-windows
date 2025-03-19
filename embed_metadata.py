#!/usr/bin/env python
import os
import re
import json
import time
import subprocess
import shutil
import ctypes
from ctypes import wintypes
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
import sys
import platform

# ------------- User Editable File Extensions -------------
FILE_EXTENSIONS = (".jpg", ".jpeg", ".png", ".mp4")
# Define which extensions are considered video files.
VIDEO_EXTENSIONS = (".mp4",)

# ---------------------------
# Tool Locator for ExifTool
# ---------------------------
def get_tool_path():
    """
    Locate ExifTool executable relative to the script (or bundled executable). 
    On Windows, if the binaries are bundled in "exiftool", use the .exe files.
    On Linux/macOS, if the folder exists, use the local binaries; otherwise, assume they are installed.
    """
    base_dir = sys._MEIPASS if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
    exiftool_dir = os.path.join(base_dir, "exiftool")
    if platform.system() == "Windows":
        exiftool_bin = "exiftool.exe"
    else:
        exiftool_bin = "exiftool"
    exiftool_path_local = os.path.join(exiftool_dir, exiftool_bin) if os.path.isdir(exiftool_dir) else exiftool_bin
    return exiftool_path_local

exiftool_path = get_tool_path()

# ---------------------------
# Windows Creation Time Setter
# ---------------------------
def set_file_times(file_path, timestamp):
    """
    Set the creation, modification, and access times of a file to the given timestamp.
    Works on Windows using ctypes.
    """
    # Convert Unix timestamp to Windows FILETIME (100-nanosecond intervals since Jan 1, 1601)
    windows_timestamp = int((timestamp + 11644473600) * 10000000)
    low = windows_timestamp & 0xFFFFFFFF
    high = windows_timestamp >> 32
    filetime = wintypes.FILETIME(low, high)
    
    # Constants for CreateFileW.
    GENERIC_WRITE = 0x40000000
    FILE_SHARE_READ = 0x00000001
    FILE_SHARE_WRITE = 0x00000002
    OPEN_EXISTING = 3
    FILE_FLAG_BACKUP_SEMANTICS = 0x02000000
    
    # Open file handle.
    handle = ctypes.windll.kernel32.CreateFileW(
        file_path,
        GENERIC_WRITE,
        FILE_SHARE_READ | FILE_SHARE_WRITE,
        None,
        OPEN_EXISTING,
        FILE_FLAG_BACKUP_SEMANTICS,
        None
    )
    if handle in (0, -1):
        return
    # Set creation, last access, and last write times.
    ctypes.windll.kernel32.SetFileTime(handle, ctypes.byref(filetime), ctypes.byref(filetime), ctypes.byref(filetime))
    ctypes.windll.kernel32.CloseHandle(handle)

# ---------------------------
# Embed Metadata with exiftool
# ---------------------------
def embed_metadata(media_file, metadata):
    """
    Embed metadata into the media file using exiftool.
    The metadata dictionary may include:
      - "timestamp": Unix timestamp (int)
      - "geo": dictionary with keys "latitude", "longitude", and optionally "altitude"
      - "camera": dictionary with key "description"
    """
    args = [exiftool_path, "-overwrite_original"]

    # Embed timestamp if available.
    if "timestamp" in metadata:
        # Format: YYYY:MM:DD HH:MM:SS (UTC)
        time_str = time.strftime("%Y:%m:%d %H:%M:%S", time.gmtime(metadata["timestamp"]))
        args.append(f"-DateTimeOriginal={time_str}")
        args.append(f"-CreateDate={time_str}")
        args.append(f"-ModifyDate={time_str}")

    # Embed geodata if available.
    if "geo" in metadata:
        geo = metadata["geo"]
        # Only embed if latitude and longitude are not both zero.
        if not (abs(geo.get("latitude", 0.0)) < 1e-6 and abs(geo.get("longitude", 0.0)) < 1e-6):
            args.append(f"-GPSLatitude={geo.get('latitude')}")
            args.append(f"-GPSLongitude={geo.get('longitude')}")
            if "altitude" in geo and abs(geo.get("altitude", 0.0)) > 1e-6:
                args.append(f"-GPSAltitude={geo.get('altitude')}")
                
    # Embed camera details if available.
    if "camera" in metadata:
        cam = metadata["camera"]
        if "description" in cam and cam["description"]:
            args.append(f"-ImageDescription={cam['description']}")

    args.append(media_file)
    subprocess.run(args, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

# ---------------------------
# Find Metadata File
# ---------------------------
def find_metadata_file(media_path):
    """
    Given a media file, try to locate its metadata JSON file.
    The file must be in the same directory and its name starts with either:
      - The full media file name (e.g. image.jpg.supplemental-metadata.json)
      - The base name of the media file (e.g. image.json)
    """
    dir_path = os.path.dirname(media_path)
    media_filename = os.path.basename(media_path)
    
    # Candidate 1: <media_filename>.supplemental-metadata.json
    candidate = os.path.join(dir_path, media_filename + ".supplemental-metadata.json")
    if os.path.isfile(candidate):
        return candidate

    # Candidate 2: Replace media extension with .json (e.g., image.jpg -> image.json)
    base, _ = os.path.splitext(media_filename)
    candidate = os.path.join(dir_path, base + ".json")
    if os.path.isfile(candidate):
        return candidate

    # Candidate 3: Any file in the same directory that starts with the full media filename and ends with .json
    for f in os.listdir(dir_path):
        if f != media_filename and f.startswith(media_filename) and f.endswith(".json"):
            candidate = os.path.join(dir_path, f)
            if os.path.isfile(candidate):
                return candidate

    # Candidate 4: Any file that starts with the base name and ends with .json
    for f in os.listdir(dir_path):
        if f != media_filename and f.startswith(base) and f.endswith(".json"):
            candidate = os.path.join(dir_path, f)
            if os.path.isfile(candidate):
                return candidate

    return None

# ---------------------------
# Extract "Date Taken" from File (for images)
# ---------------------------
def extract_date_taken(media_path):
    """
    Use exiftool to extract the DateTimeOriginal metadata from the file.
    Returns the Unix timestamp if found, or None.
    """
    try:
        result = subprocess.run([exiftool_path, "-DateTimeOriginal", "-s3", media_path],
                                capture_output=True, text=True)
        date_str = result.stdout.strip()
        if date_str:
            ts = time.mktime(time.strptime(date_str, "%Y:%m:%d %H:%M:%S"))
            return ts
    except Exception:
        pass
    return None

# ---------------------------
# Extract "Media created" from File (for videos)
# ---------------------------
def extract_media_created(media_path):
    """
    Use exiftool to extract the MediaCreateDate metadata from the file.
    Returns the Unix timestamp if found, or None.
    """
    try:
        result = subprocess.run([exiftool_path, "-MediaCreateDate", "-s3", media_path],
                                capture_output=True, text=True)
        date_str = result.stdout.strip()
        if date_str:
            ts = time.mktime(time.strptime(date_str, "%Y:%m:%d %H:%M:%S"))
            return ts
    except Exception:
        pass
    return None

# ---------------------------
# Extract Date and Time from Filename
# ---------------------------
def extract_datetime_from_filename(media_path):
    """
    Attempt to extract date and time from the filename.
    Looks for a pattern with an 8-digit date followed by an underscore and a 6-digit time.
    Returns Unix timestamp if successful, or None.
    """
    basename = os.path.basename(media_path)
    match = re.search(r"(\d{8})_(\d{6})", basename)
    if match:
        date_part = match.group(1)
        time_part = match.group(2)
        try:
            t = time.strptime(date_part + time_part, "%Y%m%d%H%M%S")
            return time.mktime(t)
        except Exception:
            return None
    return None

# ---------------------------
# Get Default Date from Directory Name
# ---------------------------
def default_date_from_directory(media_path, base_dir, dir_prefix):
    """
    Check the file's relative path for a directory component that starts with dir_prefix.
    If found, extract the first 4-digit year and return a Unix timestamp for July 1st of that year.
    Otherwise, return None.
    """
    relative = os.path.relpath(media_path, base_dir)
    components = relative.split(os.sep)
    for comp in components:
        if comp.startswith(dir_prefix):
            m = re.search(r'(\d{4})', comp)
            if m:
                year = m.group(1)
                try:
                    default_date = time.mktime(time.strptime(f"{year}-07-01 00:00:00", "%Y-%m-%d %H:%M:%S"))
                    return default_date
                except Exception:
                    pass
    return None

# ---------------------------
# Process a Single Media File
# ---------------------------
def process_media_file(media_path, base_dir, dir_prefix):
    """
    Process a single media file:
      - First, try to locate an associated metadata JSON file.
      - If no JSON file is found:
            • For video files, try extracting "Media created" metadata.
            • For images, try extracting "Date Taken" metadata.
            • If not available, attempt to extract date and time from the filename.
          If a valid date is found, update the file system times.
          Otherwise, check if any parent directory name starts with dir_prefix and extract a default date.
          If none yield a date, move the file to no_metadata/<relative_path>.
      - If a JSON metadata file is found, parse it to extract the oldest valid timestamp (from creationTime and photoTakenTime),
        and optionally geodata and camera details.
          If the timestamp remains invalid, attempt to extract it from the filename.
      - Embed the metadata via exiftool and update the file system timestamps.
    Returns a status string.
    """
    relative_path = os.path.relpath(media_path, base_dir)
    metadata_file = find_metadata_file(media_path)
    
    if not metadata_file:
        # No JSON metadata found.
        if media_path.lower().endswith(VIDEO_EXTENSIONS):
            date_extracted = extract_media_created(media_path)
            extraction_tag = "Media created extracted"
        else:
            date_extracted = extract_date_taken(media_path)
            extraction_tag = "Date taken extracted"
            
        if not date_extracted:
            # Attempt to extract date from filename.
            date_from_filename = extract_datetime_from_filename(media_path)
            if date_from_filename:
                date_extracted = date_from_filename
                extraction_tag = "Date extracted from filename"
            
        if date_extracted:
            os.utime(media_path, (date_extracted, date_extracted))
            if os.name == "nt":
                try:
                    set_file_times(media_path, date_extracted)
                except Exception:
                    pass
            return f"{relative_path}: OK ({extraction_tag})"
        else:
            # Try to get a default date from the directory name.
            default_date = default_date_from_directory(media_path, base_dir, dir_prefix)
            if default_date:
                os.utime(media_path, (default_date, default_date))
                if os.name == "nt":
                    try:
                        set_file_times(media_path, default_date)
                    except Exception:
                        pass
                return f"{relative_path}: OK (Default date from directory)"
            else:
                # Move file to no_metadata folder.
                dest_path = os.path.join(base_dir, "no_metadata", relative_path)
                os.makedirs(os.path.dirname(dest_path), exist_ok=True)
                shutil.move(media_path, dest_path)
                return f"{relative_path}: no metadata"

    # Process JSON metadata.
    try:
        with open(metadata_file, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        dest_path = os.path.join(base_dir, "no_metadata", relative_path)
        os.makedirs(os.path.dirname(dest_path), exist_ok=True)
        shutil.move(media_path, dest_path)
        return f"{relative_path}: no metadata"

    # Extract timestamps from JSON.
    try:
        ct = int(data.get("creationTime", {}).get("timestamp", "0"))
    except Exception:
        ct = 0
    try:
        pt = int(data.get("photoTakenTime", {}).get("timestamp", "0"))
    except Exception:
        pt = 0

    if ct and pt:
        timestamp = min(ct, pt)
    elif ct:
        timestamp = ct
    elif pt:
        timestamp = pt
    else:
        timestamp = None

    # Fallback: if JSON did not yield a timestamp, try extracting from filename.
    if not timestamp:
        filename_date = extract_datetime_from_filename(media_path)
        if filename_date:
            timestamp = filename_date

    # Extract and validate geodata.
    geo = None
    try:
        g = data.get("geoData", {})
        if g and "latitude" in g and "longitude" in g:
            lat = float(g.get("latitude", 0.0))
            lon = float(g.get("longitude", 0.0))
            alt = float(g.get("altitude", 0.0)) if "altitude" in g else None
            if not (abs(lat) < 1e-6 and abs(lon) < 1e-6):
                geo = {"latitude": lat, "longitude": lon}
                if alt is not None:
                    geo["altitude"] = alt
    except Exception:
        geo = None

    # Extract camera details.
    camera = None
    desc = data.get("description", "").strip()
    if desc:
        camera = {"description": desc}

    # Build metadata to embed.
    metadata_to_embed = {}
    if timestamp:
        metadata_to_embed["timestamp"] = timestamp
    if geo:
        metadata_to_embed["geo"] = geo
    if camera:
        metadata_to_embed["camera"] = camera

    # Embed metadata and update file times.
    embed_metadata(media_path, metadata_to_embed)
    if timestamp:
        os.utime(media_path, (timestamp, timestamp))
        if os.name == "nt":
            try:
                set_file_times(media_path, timestamp)
            except Exception:
                pass

    return f"{relative_path}: OK"

# ---------------------------
# Gather All Media Files
# ---------------------------
def get_all_media_files(base_dir):
    """
    Traverse the base directory recursively and return a list of media files (by extension).
    """
    media_files = []
    for root, dirs, files in os.walk(base_dir):
        if os.path.relpath(root, base_dir).startswith("no_metadata"):
            continue
        for file in files:
            if file.lower().endswith(FILE_EXTENSIONS):
                media_files.append(os.path.join(root, file))
    return media_files

# ---------------------------
# Main Function with Batch Processing and Summary
# ---------------------------
def main():
    base_dir = os.getcwd()
    media_files = get_all_media_files(base_dir)
    total_files = len(media_files)
    if total_files == 0:
        print("No media files found.")
        return

    start_time = time.time()

    try:
        batch_input = input("Enter batch size (default 8): ").strip()
        batch_size = int(batch_input) if batch_input else 8
    except Exception:
        batch_size = 8

    # Prompt for directory prefix for files without metadata.
    dir_prefix = input("Enter directory prefix for files without metadata (default 'Photos from 20'): ").strip()
    if not dir_prefix:
        dir_prefix = "Photos from 20"

    processed_count = 0
    print_lock = threading.Lock()

    with ThreadPoolExecutor(max_workers=batch_size) as executor:
        futures = [executor.submit(process_media_file, media_file, base_dir, dir_prefix) for media_file in media_files]
        for future in as_completed(futures):
            result = future.result()
            with print_lock:
                processed_count += 1
                percentage = (processed_count / total_files) * 100
                print(f"{percentage:05.2f}% {result}")

    elapsed_time = int(time.time() - start_time)
    summary = {ext: sum(1 for f in media_files if f.lower().endswith(ext)) for ext in FILE_EXTENSIONS}

    print("\nSummary:")
    for ext, count in summary.items():
        print(f"{ext}: {count}")
    print(f"Total time: {elapsed_time} seconds")

if __name__ == "__main__":
    try:
        # Check for exiftool availability.
        subprocess.run([exiftool_path, "-ver"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=True)
        main()
    except (subprocess.CalledProcessError, FileNotFoundError):
        print("ffmpeg and exiftool are missing. Please install or include them in the current directory.\n"
              "Download: https://github.com/procrastinando/compress-media-windows")
