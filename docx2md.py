# !/usr/bin/env python3
# -*- coding: utf-8 -*-

from typing import Union
from hashlib import md5
from pathlib import Path, PurePosixPath
from zipfile import ZipFile
from PIL import Image
from io import BytesIO
from subprocess import Popen

__all__ = ['docx2md']


def sava_image(file_path_name: str, archive: ZipFile, save_dir: Path, sub_dir: PurePosixPath):

    img_data = archive.read(file_path_name)
    fh = BytesIO(img_data)

    md5sum = md5(img_data).hexdigest()
    output_file_name = md5sum + '.png'
    output_file_path = sub_dir.joinpath(output_file_name)
    output_file_full_path = save_dir.joinpath(output_file_path)

    img = Image.open(fh)
    output_file_full_path.parent.mkdir(parents=True, exist_ok=True)
    img.save(output_file_full_path, "PNG")

    return PurePosixPath(file_path_name).name, output_file_path


def docx2md(src_file: Union[Path, str]):
    if isinstance(src_file, str):
        src_file = Path(src_file)
    if not src_file.exists():
        raise ValueError("{} is not exist".format(src_file))
    elif not src_file.is_file():
        raise ValueError("{} is not a file".format(src_file))
    elif src_file.suffix != '.docx':
        raise ValueError("{} is not a docx file".format(src_file))

    src_dir = src_file.parent
    doc_filename = src_file.stem
    # slug = slugify(src_file.stem)

    output_rst_file_name = doc_filename + ".md"
    output_rst_file_path = src_dir.joinpath(output_rst_file_name)

    img_save_dir = PurePosixPath(".").joinpath("img")

    archive = ZipFile(src_file, 'r')
    img_list = []

    for x in archive.namelist():
        if x.startswith("word/media/") and (not x.endswith('/')):
            try:
                origin_filename, save_path = sava_image(
                    file_path_name=x,
                    archive=archive,
                    save_dir=src_dir,
                    sub_dir=img_save_dir)
                img_list.append((origin_filename, save_path.__str__()))
            except:
                pass

    cmd = ["pandoc",
           "-f docx",
           "\"{}\"".format(src_file),
           "-t markdown",
           "-o \"{}\"".format(output_rst_file_path)]
    Popen(" ".join(cmd)).wait()

    with open(output_rst_file_path, 'rt', encoding='utf-8') as file:
        file_data = file.read()
    # print(img_list)
    for origin_filename, save_path in img_list:
        before = "media/" + origin_filename
        after = save_path

        file_data = file_data.replace(before, after)

    with open(output_rst_file_path, 'wt', encoding='utf-8') as file:
        file.write(file_data)
    return output_rst_file_path


if __name__ == "__main__":
    from pathlib import Path
    from tkinter import filedialog, Tk

    root = Tk()
    root.withdraw()
    folder_selected = filedialog.askdirectory()
    for file in Path(folder_selected).glob('*.docx'):
        docx2md(file)
