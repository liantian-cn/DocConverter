# !/usr/bin/env python3
# -*- coding: utf-8 -*-
from typing import Union
from pathlib import Path
from jinja2 import Template
from markdown import markdown

__all__ = ['md2html']

with open('template.md2html.html') as f:
    template = Template(f.read())


def md2html(src_file: Union[Path, str]):
    if isinstance(src_file, str):
        src_file = Path(src_file)
    if not src_file.exists():
        raise ValueError("{} is not exist".format(src_file))
    elif not src_file.is_file():
        raise ValueError("{} is not a file".format(src_file))
    elif src_file.suffix != '.md':
        raise ValueError("{} is not a md file".format(src_file))

    title = src_file.stem
    src_dir = src_file.parent
    output_file_name = title + ".html"
    output_file_path = src_dir.joinpath(output_file_name)

    with open(src_file, encoding='utf-8') as f:
        source = f.read()

    md_content = markdown(source, extensions=['extra', 'nl2br', 'toc', 'meta', 'codehilite'])
    page = template.render(title=title, content=md_content)

    with open(output_file_path, 'w', encoding='utf-8') as f:
        f.write(page)
    return output_file_path


if __name__ == "__main__":
    from pathlib import Path
    from tkinter import filedialog, Tk

    root = Tk()
    root.withdraw()
    folder_selected = filedialog.askdirectory()
    for file in Path(folder_selected).glob('*.md'):
        md2html(file)
