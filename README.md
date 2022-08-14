# docx2rst
将Word文档（docx）转换为rst的工具

需要先安装 pandoc
https://pandoc.org/

包含步骤

1. 用`pandoc`将docx转换为rst。
2. 用`python-docx`提取创建日期，作为`Date`属性。
3. 用`jieba`提取热词，作为`tags`。
4. 用`slugify`处理文件名，作为`slug`。
5. 用`zipfile`提取图片，用`pillow`将输出图片转为`png`，并替换`pandoc`的输出结果。

# md2html
自己拼凑的markdown转html工具

将使用github风格的css。


# docx2md
将Word文档（docx）转换为md的工具
类似docx2rst，不过没有处理tags