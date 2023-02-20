import requests
import re
from docx2python import docx2python
from bs4 import BeautifulSoup


filename = "Test.docx"

FONT_SIZE = 14
LANG_IN = "uk"
LANG_OUT = "en"

# open the document
docx_content = docx2python(filename, html=True)

# save images
docx_content.save_images('media')

# todo: try to use async functions to separate translation process and
#  html tags editing process

# parse through the content
soup = BeautifulSoup(docx_content.text, "html.parser")

# break text into paragraphs and discards empty ones
paragraphs = [p for p in soup.get_text().splitlines() if p]

# clean the paragraphs from image placeholders
text_blocks = set([])
for p in paragraphs:
    if "----media/" not in p:
        text_blocks.add(p)
    else:
        text_blocks.update(
            [block for block in p.split("----")
             if "media/image" not in block and not block.isascii()]
        )

# translate the text by sending its parts to the Google Translator API and
# create translation table
trans_table = []
for block in text_blocks:
    url = f"https://translate.googleapis.com/translate_a/single?client=gtx" \
          f"&sl={LANG_IN}&tl={LANG_OUT}&dt=t&q={block}&ie=UTF-8&oe=UTF-8"
    r = requests.get(url)
    if r.json()[0]:
        response = ""
        for item in r.json()[0]:
            if item[0]:
                response += item[0]
        trans_table.append((block, f" {response} "))

# Prepare the text for editing
text = docx_content.text.replace("\n", "").replace("\t", "")
for trans_tuple in trans_table:
    text = text.replace(trans_tuple[0], trans_tuple[1])

# Edit <span> html-tag to <p> in order to fix the problem with tags inside
# a text parsed by docx2python module. This module places paragraphs
# inside a <span> tag, that is not very great.
text = text.replace("<span", "<p").replace("span>", "p>")


# Find occurrences of a pattern like "----media/image1.png----", which are
# image placeholders
#   Match characters "----media/" and then "----" literally
#   Group (image\d\..{3}):
#     "image" matches literally
#     "\d" matches digits
#       "\d+" matches digits, one or more occurrences
#     "\." matches "."
#     ".{3}" matches exactly 3 any characters
regex_for_image = re.compile(r'----media/(image\d+\..{3})----')

# Edit image placeholders
for img_pl in regex_for_image.findall(text):
    # First, search for images inside text blocks, usually it's the wmf images,
    # which is used as formulas in MS Word
    if img_pl.endswith(".wmf") and \
            f'</p>----media/{img_pl}----<p style="font-size:28pt">' in text:
        text = text.replace(
            f'</p>----media/{img_pl}----<p style="font-size:28pt">',
            f'<span><img src="media/{img_pl}" alt="{img_pl}"></span>'
        )
    # Second, for images with an extension other than .wmf. These images are
    # usually figures that are placed in a separate paragraph
    elif not img_pl.endswith(".wmf"):
        text = text.replace(
            f'----media/{img_pl}----',
            f'<p><img src="media/{img_pl}" alt="{img_pl}"></p>'
        )
    # Third, wmf images that are placed outside of tags for some reason
    elif f'</p>----media/{img_pl}----' in text:
        text = text.replace(
            f'</p>----media/{img_pl}----',
            f'<span><img src="media/{img_pl}" alt="{img_pl}"></span></p>'
        )
    # Fourth, some other images that were not covered by previous clauses
    else:
        text = text.replace(
            f'----media/{img_pl}----',
            f'<span><img src="media/{img_pl}" alt="{img_pl}"></span>'
        )

# todo: check if docx2python always writes font size as 28pt, otherwise create
#  a mechanism to change the arbitrary font size
# By default, docx2python writes the font size as 28pt (for some reason)
if FONT_SIZE != 28:
    text = text.replace(
        'font-size:28pt',
        f'font-size:{FONT_SIZE}pt'
    )


# write a html file
filename = filename.rsplit(".", 1)[0] + ".html"
with open(filename, "w") as html_file:
    html_file.write('<!DOCTYPE html><html lang="en"><head><meta charset='
                    '"UTF-8"><title>Translation</title></head><body>')
    html_file.write(text)
    html_file.write('</body></html>')

# close the document
docx_content.close()
