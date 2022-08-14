# !/usr/bin/env python3
# -*- coding: utf-8 -*-

from typing import Union
from slugify import slugify
from pathlib import Path, PurePosixPath
from zipfile import ZipFile
from PIL import Image
from io import BytesIO
from docx import Document
from subprocess import Popen
from datetime import datetime
from jieba.analyse import extract_tags

__all__ = ['docx2rst']

head_template = """{title}
{sharp}

:Title: {title}
:Date: {date}
:tags: {tags}
:Slug: {slug}

"""


def sava_image(file_path_name: str, archive: ZipFile, save_dir: Path, sub_dir: PurePosixPath, name_prefix: str):
    origin_filename = PurePosixPath(file_path_name).name

    file_name = name_prefix + '-' + PurePosixPath(file_path_name).stem + ".png"
    save_path = sub_dir.joinpath(file_name)
    final_save_path = save_dir.joinpath(save_path)

    img_data = archive.read(file_path_name)
    fh = BytesIO(img_data)
    img = Image.open(fh)
    final_save_path.parent.mkdir(parents=True, exist_ok=True)
    img.save(final_save_path, "PNG")

    return origin_filename, save_path


def docx2rst(src_file: Union[Path, str]):
    if isinstance(src_file, str):
        src_file = Path(src_file)
    if not src_file.exists():
        raise ValueError("{} is not exist".format(src_file))
    elif not src_file.is_file():
        raise ValueError("{} is not a file".format(src_file))
    elif src_file.suffix != '.docx':
        raise ValueError("{} is not a docx file".format(src_file))

    document = Document(src_file)
    src_dir = src_file.parent
    slug = slugify(src_file.stem)
    doc_filename = src_file.stem

    output_rst_file_name = slug + ".rst"
    output_rst_file_path = src_dir.joinpath(output_rst_file_name)

    try:
        created = document.core_properties.created
    except:
        created = datetime.now()

    media_save_dir = PurePosixPath("."). \
        joinpath("media"). \
        joinpath(created.strftime("%Y-%m-%d"))

    archive = ZipFile(src_file, 'r')
    img_list = []

    for x in archive.namelist():
        if x.startswith("word/media/") and (not x.endswith('/')):
            origin_filename, save_path = sava_image(
                file_path_name=x,
                archive=archive,
                save_dir=src_dir,
                sub_dir=media_save_dir,
                name_prefix=slug)
            img_list.append((origin_filename, save_path.__str__()))

    cmd = ["pandoc",
           "-f docx",
           "\"{}\"".format(src_file),
           "-t rst",
           "-o \"{}\"".format(output_rst_file_path)]
    Popen(" ".join(cmd)).wait()

    with open(output_rst_file_path, 'rt', encoding='utf-8') as file:
        file_data = file.read()

    for origin_filename, save_path in img_list:
        before = "image:: media/" + origin_filename
        after = "image:: " + save_path

        file_data = file_data.replace(before, after)

    tags = extract_tags(file_data, topK=4, allowPOS=('ns', 'n', 'vn', 'v', 'nz', 'vn', 'f', 's', 'ns',))
    file_head = head_template.format(
        title=doc_filename,
        sharp="#" * 2 * len(doc_filename),
        date=document.core_properties.created.strftime("%Y-%m-%d %H:%M"),
        slug=slug,
        tags=", ".join(tags)

    )
    file_data = file_head + file_data
    with open(output_rst_file_path, 'wt', encoding='utf-8') as file:
        file.write(file_data)


if __name__ == "__main__":
    from pathlib import Path
    from tkinter import filedialog, Tk

    root = Tk()
    root.withdraw()
    folder_selected = filedialog.askdirectory()
    for file in Path(folder_selected).glob('*.docx'):
        docx2rst(file)
