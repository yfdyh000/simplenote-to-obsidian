#!/usr/bin/env python3

import json
import os
import re
import sys
from datetime import datetime
import pytz
from subprocess import call
from pathvalidate import sanitize_filename
import win32file
import win32con
import pywintypes



# Path to the JSON file we'll read in:
INPUT_FILE = "./notes.json"

# Path to the directory where we'll save the converted notes:
OUTPUT_DIRECTORY = "./Simplenote_converted/"

# Should the creation time of the created files be set to the creation
# time of the original notes?
KEEP_ORIGINAL_CREATION_TIME = True

# Should the last-modified time of the created files be set to the
# last-modified time of the original notes?
KEEP_ORIGINAL_MODIFIED_TIME = True

KEEP_METAINFO_INTO_YAML = True

# Should convert dates to local timezone?
CONVERT_TO_LOCAL_TIMEZONE = True

def convert_to_local_time(utc_time_str):
    """Convert UTC time string to local timezone ISO format."""
    if not CONVERT_TO_LOCAL_TIMEZONE:
        return utc_time_str
        
    utc_time = datetime.strptime(utc_time_str, "%Y-%m-%dT%H:%M:%S.%fZ")
    utc_time = utc_time.replace(tzinfo=pytz.UTC)
    local_time = utc_time.astimezone()
    return local_time.isoformat()

def main():

    ###################################################################
    # 1. Set-up and checking.

    if not os.path.exists(INPUT_FILE):
        sys.exit(f"There is no file at {INPUT_FILE}")

    if not os.path.isfile(INPUT_FILE):
        sys.exit(f"{INPUT_FILE} is not a file")

    tag_position = input("\nWhere should tags be put? Either 'start' or 'end' (default is 'start'):")

    if tag_position == "":
        tag_position = "start"

    if tag_position not in ["start", "end"]:
        sys.exit("Enter either 'start' or 'end'.")

    if not os.path.isdir(OUTPUT_DIRECTORY):
        os.mkdir(OUTPUT_DIRECTORY)

    ###################################################################
    # 2. Loop through all the notes and create new ones

    # Empty line before next output
    print("")

    # The keys will be filenames, the values will be an integer -
    # the number of times that filename was used.
    filenames = {}

    with open(INPUT_FILE, encoding="UTF-8") as json_file:
        # Load the JSON data into a dict:
        try:
            data = json.load(json_file)
        except json.decoder.JSONDecodeError as e:
            sys.exit(f"Could not parse {INPUT_FILE}. Are you sure it's a JSON file?")

        if not isinstance(data, dict):
            sys.exit(f"The data from {INPUT_FILE} is not a dict, so it can't be used.")

        if "activeNotes" not in data:
            sys.exit(f"There is no 'activeNotes' element in the data found in {INPUT_FILE}")

        for note in data["activeNotes"]:
            # Get all the note's lines into a list:
            lines = note["content"].splitlines()

            if len(lines) == 0:
                # We'll skip any empty notes
                print(f"Skipping empty note with ID of {note['id']}")
            else:
                # 保存原始内容用于文件名
                original_content = note["content"].splitlines()
                filename_start = original_content[0] if original_content else ""
                
                # 处理内容行
                content_lines = note["content"].splitlines()
                
                # 准备YAML front matter
                yaml_front_matter = []
                if KEEP_METAINFO_INTO_YAML:
                    creation_date = convert_to_local_time(note['creationDate'])
                    modified_date = convert_to_local_time(note['lastModified'])
                    
                    yaml_front_matter = [
                        "---",
                        f"Simplenote_id: {note['id']}",
                        f"Simplenote_creationDate: {creation_date}",
                        f"Simplenote_lastModified: {modified_date}",
                        "---",
                        ""
                    ]
                    
                    # 将YAML添加到内容前面
                    content_lines = yaml_front_matter + content_lines

                # 处理标签
                if "tags" in note:
                    tags = note["tags"]
                    tags = [re.sub(r'\W+', '_', tag) for tag in tags]
                    tags = ["#"+tag for tag in tags]
                    tag_text = " ".join(tags)

                    if tag_position == "start":
                        if KEEP_METAINFO_INTO_YAML:
                            # 在YAML之后插入标签
                            insert_pos = len(yaml_front_matter)
                            content_lines.insert(insert_pos, "")
                            content_lines.insert(insert_pos + 1, tag_text)
                        else:
                            content_lines.insert(1, "")
                            content_lines.insert(2, tag_text)
                    else:
                        content_lines.append("")
                        content_lines.append(tag_text)

                # 使用原始内容的第一行创建文件名
                if len(filename_start) > 100:
                    base_filename = filename_start[0:100]
                else:
                    base_filename = filename_start

                # 使用pathvalidate库规范化文件名，指定替换字符为下划线
                filename = sanitize_filename(
                    base_filename, 
                    platform="auto",
                    replacement_text="_"  # 使用下划线替换所有非法字符
                ) + ".md"

                filepath = os.path.join(OUTPUT_DIRECTORY, filename)

                # 处理文件名重复
                if filename in filenames:
                    filenames[filename] += 1
                else:
                    filenames[filename] = 1

                if os.path.exists(filepath):
                    sanitized_base = sanitize_filename(
                        f"{filename[:-3]} {filenames[filename]}", 
                        platform="auto",
                        replacement_text="_"  # 保持一致的替换字符
                    )
                    filename = f"{sanitized_base}.md"
                    filepath = os.path.join(OUTPUT_DIRECTORY, filename)

                # 写入文件
                with open(filepath, "w", encoding="UTF-8") as outfile:
                    outfile.write("\n".join(content_lines))

                if KEEP_ORIGINAL_CREATION_TIME:
                    try:
                        # 将UTC时间转换为Windows文件时间格式
                        creation_time = datetime.strptime(
                            note["creationDate"], "%Y-%m-%dT%H:%M:%S.%fZ"
                        )
                        win_time = pywintypes.Time(creation_time)
                        
                        # 打开文件句柄
                        handle = win32file.CreateFile(
                            filepath,
                            win32con.GENERIC_WRITE,
                            0,
                            None,
                            win32con.OPEN_EXISTING,
                            win32con.FILE_ATTRIBUTE_NORMAL,
                            None
                        )
                        
                        # 设置文件时间
                        win32file.SetFileTime(
                            handle,
                            win_time,  # 创建时间
                            None,      # 访问时间（保持不变）
                            None       # 修改时间（保持不变）
                        )
                        
                        # 关闭文件句柄
                        handle.Close()
                    except Exception as e:
                        print(f"Warning: Could not set creation time for {filepath}: {e}")

                if KEEP_ORIGINAL_MODIFIED_TIME:
                    # Set the file access and modified times:
                    modified_time = datetime.strptime(
                        note["lastModified"], "%Y-%m-%dT%H:%M:%S.%fZ"
                    )
                    modified_time = modified_time.timestamp()
                    os.utime(filepath, (modified_time, modified_time))

    num_files = sum(filenames.values())
    files_were = "file was" if num_files == 1 else "files were"
    print(f"\n{num_files} .md {files_were} created in {OUTPUT_DIRECTORY}")


if __name__ == "__main__":
    main()
