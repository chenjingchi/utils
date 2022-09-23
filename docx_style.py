import docx
import os
import re

from docx.shared import Pt

INPUT_ROOT = "path/to/input/txt/directory"
OUTPUT_ROOT = "path/to/output/docx/directory"

TIBETAN_FONT_SIZE = 20
TIBETAN_FONT_NAME = 'Microsoft Himalaya'

CHINESE_FONT_SIZE = 10.5  # 小四
CHINESE_FONT_NAME = '仿宋'


for root, dirs, files in os.walk(INPUT_ROOT):
    for filename in files:
        print(filename)
        if filename.endswith(".txt"):
            fullpath = os.path.join(root, filename)
            print(fullpath)
            with open(fullpath, "rb") as file:
                txt = file.read().decode('utf-8')

                # Adjust output font size and style.
                doc = docx.Document()
                for idx, para_txt in enumerate(txt.split("\n")):
                    # Skip digit line.
                    if re.match(r'[\.0-9]+', para_txt):
                        continue
                    para = doc.add_paragraph()
                    font = para.add_run(para_txt).font
                    # is Tibetan
                    if re.findall(r'[\u0f00-\u0fff]+', para_txt):
                        font.size = Pt(TIBETAN_FONT_SIZE)
                        font.name = TIBETAN_FONT_NAME
                    else:
                        font.size = Pt(CHINESE_FONT_SIZE)
                        font.name = CHINESE_FONT_NAME

                output_filename = os.path.splitext(filename)[0]
                output_dir = os.path.join(OUTPUT_ROOT, os.path.basename(root))
                if not os.path.exists(output_dir):
                    os.makedirs(output_dir)
                output_fullpath = os.path.join(
                    output_dir, output_filename) + ".docx"
                doc.save(output_fullpath)
