import os
import sys
import re
import subprocess
import shutil
import platform
import ctypes
from ctypes import wintypes
from concurrent.futures import ThreadPoolExecutor, as_completed

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

            # Open source file for reading timestamps.
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

            # Open destination file for writing attributes.
            FILE_WRITE_ATTRIBUTES = 0x0100
            dst_handle = kernel32.CreateFileW(dst, FILE_WRITE_ATTRIBUTES, 0, None, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, None)
            if dst_handle in (0, -1):
                return
            # Set destination file creation time to that of the source.
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
        ], check=True)
    except subprocess.CalledProcessError as e:
        print(f"Error copying metadata from {input_file} to {output_file}: {e}")

def get_video_bitrate(file_path):
    """
    Extract the video bitrate (in kbps) using FFmpeg.
    """
    try:
        result = subprocess.run(
            [ffmpeg_path, "-i", file_path],
            stderr=subprocess.PIPE,
            stdout=subprocess.DEVNULL,
            text=True
        )
        match = re.search(r"bitrate:\s(\d+)\s", result.stderr)
        bitrate = int(match.group(1)) if match else 0
        print(f"The bitrate of the video {file_path} is: {bitrate} kbps")
        return bitrate
    except Exception as e:
        print(f"Error getting bitrate for {file_path}: {e}")
        return 0

def compress_video(input_file, temp_output, video_bitrate, audio_bitrate, use_nvidia):
    """
    Compress the video using FFmpeg while copying global metadata.
    Chooses between NVIDIA accelerated (h264_nvenc) or CPU-based (libx265) encoding.
    """
    codec = "h264_nvenc" if use_nvidia else "libx265"
    tag = "avc1" if codec == "h264_nvenc" else "hvc1"
    subprocess.run([
        ffmpeg_path, "-loglevel", "quiet",  "-y", "-i", input_file,
        "-map_metadata", "0",
        "-movflags", "use_metadata_tags",
        "-c:v", codec, "-b:v", f"{video_bitrate}k",
        "-c:a", "aac", "-b:a", f"{audio_bitrate}k",
        "-tag:v", tag,
        temp_output
    ], check=True)

def compress_and_preserve_metadata(input_file, temp_output, video_bitrate, audio_bitrate, use_nvidia):
    """
    Compress the video and then copy metadata from the original file.
    """
    compress_video(input_file, temp_output, video_bitrate, audio_bitrate, use_nvidia)
    copy_metadata(input_file, temp_output)

def compress_image(input_file, temp_output, quality):
    """
    Compress the image using FFmpeg and then copy metadata.
    The command uses '-frames:v 1' so that FFmpeg writes a single image.
    """
    subprocess.run([
        ffmpeg_path, "-loglevel", "quiet",  "-y", "-i", input_file,
        "-q:v", str(quality), "-frames:v", "1", temp_output
    ], check=True)
    copy_metadata(input_file, temp_output)

def process_file(input_file, output_dir, replace_original, video_bitrate, audio_bitrate, jpg_quality, use_nvidia):
    """
    Process a single media file:
      - For videos: if bitrate is above the threshold, compress and preserve metadata.
      - For images: compress and preserve metadata.
    After conversion, file timestamps are preserved.
    """
    filename = os.path.basename(input_file)
    temp_output = f"{os.path.splitext(input_file)[0]}_temp{os.path.splitext(input_file)[1]}"
    output_file = input_file if replace_original else os.path.join(output_dir, filename)
    try:
        if input_file.lower().endswith(".mp4"):
            current_bitrate = get_video_bitrate(input_file)
            if current_bitrate > video_bitrate:
                print(f"Compressing video: {input_file}")
                compress_and_preserve_metadata(input_file, temp_output, video_bitrate, audio_bitrate, use_nvidia)
                if replace_original:
                    os.replace(temp_output, input_file)
                else:
                    os.rename(temp_output, output_file)
                preserve_file_timestamps(input_file, output_file)
            else:
                print(f"Skipping video (bitrate below threshold): {input_file}")
        elif input_file.lower().endswith((".jpg", ".jpeg")):
            print(f"Compressing image: {input_file}")
            compress_image(input_file, temp_output, jpg_quality)
            if replace_original:
                os.replace(temp_output, input_file)
            else:
                os.rename(temp_output, output_file)
            preserve_file_timestamps(input_file, output_file)
        else:
            print(f"Skipping unsupported file type: {input_file}")
    except subprocess.CalledProcessError as e:
        print(f"Error processing {input_file}: {e}")

def get_media_files(root_dir):
    """
    Recursively collect media files (MP4 and JPG/JPEG) from the root directory.
    Excludes directories containing 'ffmpeg' or 'exiftool' to avoid processing tool files.
    """
    media_files = []
    for dirpath, dirnames, filenames in os.walk(root_dir):
        if any(excluded in dirpath.lower() for excluded in ["ffmpeg", "exiftool"]):
            continue
        for filename in filenames:
            if filename.lower().endswith((".mp4", ".jpg", ".jpeg")):
                media_files.append(os.path.join(dirpath, filename))
    return media_files

def get_tool_paths():
    """
    Locate FFmpeg and ExifTool executables relative to the script (or bundled executable).
    """
    base_dir = sys._MEIPASS if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
    
    ffmpeg_dir = os.path.join(base_dir, "ffmpeg", "bin")
    exiftool_dir = os.path.join(base_dir, "exiftool")
    
    ffmpeg_path = os.path.join(ffmpeg_dir, "ffmpeg.exe") if os.path.isdir(ffmpeg_dir) else "ffmpeg"
    exiftool_path = os.path.join(exiftool_dir, "exiftool.exe") if os.path.isdir(exiftool_dir) else "exiftool"
    
    return ffmpeg_path, exiftool_path

ffmpeg_path, exiftool_path = get_tool_paths()

if __name__ == "__main__":
    try:
        root_directory = input(f"Root directory (default: {os.getcwd()}): ").strip() or os.getcwd()
        video_bitrate = int(input("Enter video bitrate threshold in kbps (default 3000): ") or 3000)
        audio_bitrate = int(input("Enter audio bitrate threshold in kbps (default 192): ") or 192)
        jpg_quality = int(input("Enter JPG quality (default 7): ") or 7)
        replace_original_input = input("Replace original files? (yes or no, default yes): ").strip().lower() or "yes"
        use_nvidia_input = input("Use NVIDIA NVENC for conversion? (yes or no, default yes): ").strip().lower() or "yes"
        replace_original = (replace_original_input == "yes")
        use_nvidia = (use_nvidia_input == "yes")
    except ValueError:
        print("Invalid input, using default values.")
        root_directory = os.getcwd()
        video_bitrate = 3000
        audio_bitrate = 192
        jpg_quality = 7
        replace_original = True
        use_nvidia = True

    batch_size = 1
    if not use_nvidia:
        try:
            batch_input = input("Enter CPU batch size for parallel conversion (default 2): ").strip()
            batch_size = int(batch_input) if batch_input else 2
            print(f"Using batch size: {batch_size}")
        except ValueError:
            batch_size = 2
            print("Invalid input for batch size, defaulting to 2.")

    output_dir = os.path.join(root_directory, "Compressed")
    if not replace_original:
        os.makedirs(output_dir, exist_ok=True)

    media_files = get_media_files(root_directory)
    if not media_files:
        print("No media files found to process.")
        sys.exit(0)

    print(f"Found {len(media_files)} media files to process.")

    if use_nvidia:
        # GPU mode: Process files sequentially.
        for file in media_files:
            process_file(file, output_dir, replace_original, video_bitrate, audio_bitrate, jpg_quality, use_nvidia)
    else:
        # CPU mode: Process files in parallel using ThreadPoolExecutor.
        with ThreadPoolExecutor(max_workers=batch_size) as executor:
            futures = {executor.submit(process_file, file, output_dir, replace_original,
                                       video_bitrate, audio_bitrate, jpg_quality, use_nvidia): file
                       for file in media_files}
            for future in as_completed(futures):
                file = futures[future]
                try:
                    future.result()
                except Exception as exc:
                    print(f"Error processing {file}: {exc}")

    print("Compression process completed!")
