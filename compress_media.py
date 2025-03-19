import os
import sys
import re
import subprocess
import shutil
import platform
import ctypes
import time
from ctypes import wintypes
from concurrent.futures import ThreadPoolExecutor, as_completed

def is_intel_cpu():
    """
    Attempt to determine if the CPU is from Intel.
    """
    proc = platform.processor().lower()
    if "intel" in proc:
        return True
    if os.path.exists("/proc/cpuinfo"):
        try:
            with open("/proc/cpuinfo", "r") as f:
                content = f.read().lower()
                return "intel" in content or "genuineintel" in content
        except Exception:
            pass
    return False

def is_arm_cpu():
    """
    Determine if the machine is ARM-based.
    """
    machine = platform.machine().lower()
    return "arm" in machine or "aarch64" in machine

def preserve_file_timestamps(src, dst):
    """
    Copy file system timestamps from src to dst.
    Uses shutil.copystat for access and modification times.
    On Windows, also attempts to copy the file creation time.
    """
    try:
        shutil.copystat(src, dst)
    except Exception as e:
        print(f"Error copying basic timestamps: {e}")
    if platform.system() == "Windows":
        try:
            kernel32 = ctypes.windll.kernel32
            GENERIC_READ = 0x80000000
            FILE_SHARE_READ = 0x00000001
            OPEN_EXISTING = 3
            FILE_FLAG_BACKUP_SEMANTICS = 0x02000000

            src_handle = kernel32.CreateFileW(src, GENERIC_READ, FILE_SHARE_READ, None, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, None)
            if src_handle in (0, -1):
                return
            ctime = wintypes.FILETIME()
            atime = wintypes.FILETIME()
            mtime = wintypes.FILETIME()
            if not kernel32.GetFileTime(src_handle, ctypes.byref(ctime), ctypes.byref(atime), ctypes.byref(mtime)):
                kernel32.CloseHandle(src_handle)
                return
            kernel32.CloseHandle(src_handle)

            FILE_WRITE_ATTRIBUTES = 0x0100
            dst_handle = kernel32.CreateFileW(dst, FILE_WRITE_ATTRIBUTES, 0, None, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, None)
            if dst_handle in (0, -1):
                return
            kernel32.SetFileTime(dst_handle, ctypes.byref(ctime), None, None)
            kernel32.CloseHandle(dst_handle)
        except Exception as e:
            print(f"Error preserving creation time on Windows: {e}")

def copy_metadata(input_file, output_file):
    """
    Use ExifTool to copy all metadata from the original file to the converted file.
    """
    try:
        subprocess.run([
            exiftool_path, "-TagsFromFile", input_file,
            "-all:all>all:all", output_file,
            "-overwrite_original"
        ], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except subprocess.CalledProcessError as e:
        print(f"Error copying metadata from {input_file} to {output_file}: {e}")

def get_video_bitrate(file_path):
    """
    Extract the video bitrate (in kbps) using ffprobe.
    ffprobe returns the bitrate in bits per second.
    """
    try:
        result = subprocess.run(
            [ffprobe_path, "-v", "error", "-select_streams", "v:0",
             "-show_entries", "stream=bit_rate", "-of", "default=noprint_wrappers=1:nokey=1", file_path],
            stdout=subprocess.PIPE,
            stderr=subprocess.DEVNULL,
            text=True
        )
        bitrate_str = result.stdout.strip()
        if bitrate_str.isdigit():
            # Convert to kbps (divide by 1000)
            return int(bitrate_str) // 1000
        else:
            return 0
    except Exception as e:
        print(f"Error getting bitrate for {file_path}: {e}")
        return 0

def compress_video(input_file, temp_output, video_bitrate, audio_bitrate, acceleration):
    """
    Compress the video using FFmpeg while copying global metadata.
    Chooses codec based on hardware acceleration:
      - "nvidia": uses h264_nvenc
      - "intel": uses h264_qsv
      - "cpu": uses libx265 (with x265 logging suppressed)
    """
    if acceleration == "nvidia":
        codec = "h264_nvenc"
        tag = "avc1"
    elif acceleration == "intel":
        codec = "h264_qsv"
        tag = "avc1"
    else:
        codec = "libx265"
        tag = "hvc1"
    cmd = [
        ffmpeg_path, "-loglevel", "quiet", "-y", "-i", input_file,
        "-map_metadata", "0",
        "-movflags", "use_metadata_tags",
        "-c:v", codec, "-b:v", f"{video_bitrate}k",
        "-c:a", "aac", "-b:a", f"{audio_bitrate}k",
    ]
    if codec == "libx265":
        # Suppress x265 logging
        cmd.extend(["-x265-params", "log-level=0"])
    cmd.extend(["-tag:v", tag, temp_output])
    subprocess.run(cmd, check=True)

def compress_and_preserve_metadata(input_file, temp_output, video_bitrate, audio_bitrate, acceleration):
    """
    Compress the video and then copy metadata from the original file.
    """
    compress_video(input_file, temp_output, video_bitrate, audio_bitrate, acceleration)
    copy_metadata(input_file, temp_output)

def compress_image(input_file, temp_output, quality):
    """
    Compress the image using FFmpeg and then copy metadata.
    The command uses '-frames:v 1' so that FFmpeg writes a single image.
    """
    subprocess.run([
        ffmpeg_path, "-loglevel", "quiet", "-y", "-i", input_file,
        "-q:v", str(quality), "-frames:v", "1", temp_output
    ], check=True)
    copy_metadata(input_file, temp_output)

def process_file(input_file, output_dir, replace_original, video_bitrate, audio_bitrate, jpg_quality, acceleration):
    """
    Process a single media file:
      - For videos: if bitrate is above the threshold, compress and preserve metadata.
      - For images: compress and preserve metadata.
    After conversion, file timestamps are preserved.
    Returns a tuple: (input_file, action, extension) where action is "OK" or "SKIPPED".
    """
    filename = os.path.basename(input_file)
    temp_output = f"{os.path.splitext(input_file)[0]}_temp{os.path.splitext(input_file)[1]}"
    output_file = input_file if replace_original else os.path.join(output_dir, filename)
    action = "SKIPPED"
    try:
        if input_file.lower().endswith(".mp4"):
            current_bitrate = get_video_bitrate(input_file)
            if current_bitrate > video_bitrate:
                compress_and_preserve_metadata(input_file, temp_output, video_bitrate, audio_bitrate, acceleration)
                if replace_original:
                    os.replace(temp_output, input_file)
                else:
                    os.rename(temp_output, output_file)
                preserve_file_timestamps(input_file, output_file)
                action = "OK"
            else:
                action = "SKIPPED"
        elif input_file.lower().endswith((".jpg", ".jpeg")):
            compress_image(input_file, temp_output, jpg_quality)
            if replace_original:
                os.replace(temp_output, input_file)
            else:
                os.rename(temp_output, output_file)
            preserve_file_timestamps(input_file, output_file)
            action = "OK"
        elif input_file.lower().endswith(".png"):
            # For .png files, processing can be added if needed.
            action = "SKIPPED"
        else:
            action = "SKIPPED"
    except subprocess.CalledProcessError as e:
        print(f"Error processing {input_file}: {e}")
    return (input_file, action, os.path.splitext(input_file)[1].lower())

def get_media_files(root_dir):
    """
    Recursively collect media files (MP4, JPG/JPEG, PNG) from the root directory.
    Excludes directories containing 'ffmpeg' or 'exiftool' to avoid processing tool files.
    """
    media_files = []
    for dirpath, dirnames, filenames in os.walk(root_dir):
        if any(excluded in dirpath.lower() for excluded in ["ffmpeg", "exiftool"]):
            continue
        for filename in filenames:
            if filename.lower().endswith((".mp4", ".jpg", ".jpeg", ".png")):
                media_files.append(os.path.join(dirpath, filename))
    return media_files

def get_tool_paths():
    """
    Locate FFmpeg, ffprobe, and ExifTool executables relative to the script (or bundled executable).
    On Windows, if the binaries are bundled in "ffmpeg/bin" or "exiftool", use the .exe files.
    On Linux/macOS, if the folders exist, use the local binaries; otherwise, assume they are installed.
    """
    base_dir = sys._MEIPASS if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
    
    ffmpeg_dir = os.path.join(base_dir, "ffmpeg", "bin")
    exiftool_dir = os.path.join(base_dir, "exiftool")
    
    if platform.system() == "Windows":
        ffmpeg_bin = "ffmpeg.exe"
        ffprobe_bin = "ffprobe.exe"
        exiftool_bin = "exiftool.exe"
    else:
        ffmpeg_bin = "ffmpeg"
        ffprobe_bin = "ffprobe"
        exiftool_bin = "exiftool"
    
    if os.path.isdir(ffmpeg_dir):
        ffmpeg_path_local = os.path.join(ffmpeg_dir, ffmpeg_bin)
        ffprobe_path_local = os.path.join(ffmpeg_dir, ffprobe_bin)
    else:
        ffmpeg_path_local = ffmpeg_bin
        ffprobe_path_local = ffprobe_bin

    exiftool_path_local = os.path.join(exiftool_dir, exiftool_bin) if os.path.isdir(exiftool_dir) else exiftool_bin
    
    return ffmpeg_path_local, ffprobe_path_local, exiftool_path_local

ffmpeg_path, ffprobe_path, exiftool_path = get_tool_paths()

if __name__ == "__main__":
    try:
        # Check for tool availability.
        subprocess.run([ffmpeg_path, "-version"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=True)
        subprocess.run([ffprobe_path, "-version"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=True)
        subprocess.run([exiftool_path, "-ver"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=True)

        start_time = time.time()
        try:
            root_directory = input(f"Root directory (default: {os.getcwd()}): ").strip() or os.getcwd()
            video_bitrate = int(input("Enter video bitrate threshold in kbps (default 3000): ") or 3000)
            audio_bitrate = int(input("Enter audio bitrate threshold in kbps (default 192): ") or 192)
            jpg_quality = int(input("Enter JPG quality (default 7): ") or 7)
            replace_original_input = input("Replace original files? (yes or no, default yes): ").strip().lower() or "yes"
            replace_original = (replace_original_input == "yes")
        except ValueError:
            print("Invalid input, using default values.")
            root_directory = os.getcwd()
            video_bitrate = 3000
            audio_bitrate = 192
            jpg_quality = 7
            replace_original = True

        # Determine hardware acceleration:
        use_nvidia_input = input("Use NVIDIA NVENC for conversion? (yes or no, default yes): ").strip().lower() or "yes"
        if use_nvidia_input == "yes":
            acceleration = "nvidia"
        else:
            use_intel_input = input("Use Intel QuickSync for conversion? (yes or no, default yes): ").strip().lower() or "yes"
            if use_intel_input == "yes":
                acceleration = "intel"
            else:
                acceleration = "cpu"

        # Check for potential hardware mismatches:
        if acceleration == "intel" and not is_intel_cpu():
            print("Warning: Intel QuickSync was selected, but your CPU does not appear to be an Intel processor.\n"
                  "Hardware acceleration may fail. Reverting to CPU mode.")
            acceleration = "cpu"
        if is_arm_cpu() and acceleration in ("nvidia", "intel"):
            print("Warning: Hardware acceleration (NVIDIA/Intel) is not supported on ARM processors.\n"
                  "Reverting to CPU mode.")
            acceleration = "cpu"

        # Ask for batch size for parallel conversion regardless of acceleration type.
        try:
            batch_input = input("Enter batch size for parallel conversion (default 2): ").strip()
            batch_size = int(batch_input) if batch_input else 2
            print(f"Using batch size: {batch_size}")
        except ValueError:
            batch_size = 2
            print("Invalid input for batch size, defaulting to 2.")

        output_dir = os.path.join(root_directory, "Compressed")
        if not replace_original:
            os.makedirs(output_dir, exist_ok=True)

        media_files = get_media_files(root_directory)
        total_files = len(media_files)
        if total_files == 0:
            print("No media files found to process.")
            sys.exit(0)

        # Summary counts for extensions.
        summary = {'.jpg': 0, '.jpeg': 0, '.png': 0, '.mp4': 0}
        results = [None] * total_files  # placeholder for ordered results

        completed_count = 0
        # Process files in parallel and print progress as each file finishes.
        with ThreadPoolExecutor(max_workers=batch_size) as executor:
            future_to_index = {}
            for idx, file in enumerate(media_files):
                future = executor.submit(process_file, file, output_dir, replace_original,
                                           video_bitrate, audio_bitrate, jpg_quality, acceleration)
                future_to_index[future] = idx
            for future in as_completed(future_to_index):
                idx = future_to_index[future]
                result = future.result()
                results[idx] = result
                completed_count += 1
                summary[result[2]] = summary.get(result[2], 0) + 1
                rel_path = os.path.relpath(result[0], root_directory)
                percent = (completed_count / total_files) * 100
                print(f"{percent:05.2f}% {rel_path}: {result[1]}")

        elapsed_time = int(time.time() - start_time)
        print("\n>>> Summary <<<")
        for ext in ['.jpg', '.jpeg', '.png', '.mp4']:
            print(f"{ext}: {summary.get(ext, 0)}")
        print(f"Total time: {elapsed_time} seconds")

    except (subprocess.CalledProcessError, FileNotFoundError):
        print("ffmpeg, ffprobe, and/or exiftool are missing. Please install or include them in the current directory.\n"
              "Download: https://github.com/procrastinando/compress-media-windows")
