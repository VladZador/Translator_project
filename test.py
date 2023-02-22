import json
import time
import requests
import re
from docx2python import docx2python
from bs4 import BeautifulSoup

start1 = time.perf_counter()

filename = "Test.docx"

LANG_IN = "uk"
LANG_OUT = "en"

# open the document, save images and write the text into a variable
with docx2python(filename, image_folder="media", html=True) as docx_content:
    docx_text = docx_content.text

start2 = time.perf_counter()
print("Opening the document, saving images:", start2 - start1)

# parse through the content
soup = BeautifulSoup(docx_text, "html.parser")

# break text into paragraphs and discards empty ones
paragraphs = [p for p in soup.get_text().splitlines() if p]

# clean the paragraphs from image placeholders. isascii() method is used in
# order to discard empty strings or those containing punctuation marks or
# words in English.
text_blocks = set([])
for p in paragraphs:
    if f"----media/" not in p and not p.isascii():
        text_blocks.add(p)
    else:
        text_blocks.update(
            [block for block in p.split("----")
             if f"media/image" not in block and not block.isascii()]
        )

# Now, when there's no image placeholders, break text into sentences
sentences = set([])
for block in text_blocks:
    sentences.update(
        [s for s in block.split(".") if not s.isascii()]
    )

start3 = time.perf_counter()
print("Breaking the text into segments:", start3 - start2)


def translate_with_google(string):
    url = f"https://translate.googleapis.com/translate_a/single?client=gtx" \
          f"&sl={LANG_IN}&tl={LANG_OUT}&dt=t&q={string}&ie=UTF-8&oe=UTF-8"
    r = requests.get(url)
    response = ""
    if r.json()[0]:
        for item in r.json()[0]:
            if item[0]:
                response += item[0]
    return response


# translate the text by sending its parts to the Google Translator API and
# create translation table
trans_table = []
for s in sorted(sentences, key=len, reverse=True):
    translated_str = translate_with_google(s)
    trans_table.append((s, f" {translated_str} "))

# todo: remove this and other debugging "prints"
with open("translation_table.json", "w") as json_file:
    json_file.write(json.dumps(dict(trans_table)))

# Prepare the text for editing; replace strings using the translation table
text = docx_text.replace("\n", "")
for trans_tuple in trans_table:
    text = text.replace(trans_tuple[0], trans_tuple[1])

start4 = time.perf_counter()
print("Translation:", start4 - start3)

# Edit <span> html-tag to <p> in order to fix the problem with tags inside
# a text parsed by docx2python module. This module places paragraphs
# inside a <span> tag, that is not very great.
text = text.replace("<span", "<p").replace("span>", "p>")


# Find occurrences of a pattern like "----media/image1.png----",
# "----media/image6.jpeg----", which are image placeholders.
#       Match characters "----media/" and then "----" literally
#       Group (image\d+\.(.{3}|.{4})):
#           "image" matches literally
#           "\d" matches digits
#               "\d+" matches digits, one or more occurrences
#           "\." matches "."
#           Group (.{3}|.{4}) matches exactly 3 or 4 any characters:
#               ".{3}" matches exactly 3 any characters
#               ".{4}" matches exactly 4 any characters
regex_for_image = re.compile(r'----media/(image\d+\.(.{3}|.{4}))----')

# Find occurrences of a pattern like "font-size:28pt"
#       "\d+" matches digits, one or more occurrences
regex_for_font = re.compile(r'font-size:(\d+)pt')
font_sizes = set(regex_for_font.findall(text))

# Edit image placeholders
for img_pl in regex_for_image.findall(text):
    img_pl = img_pl[0]
    for fs in font_sizes:
        # First, search for images inside text blocks, usually it's the wmf
        # images, which is used as formulas in MS Word
        if img_pl.endswith(".wmf") and \
                f'</p>----media/{img_pl}----<p style="font-size:{fs}pt">' \
                in text:
            text = text.replace(
                f'</p>----media/{img_pl}----<p style="font-size:{fs}pt">',
                f'<span><img src="media/{img_pl}" alt="{img_pl}"></span>'
            )
        # Second, for images with an extension other than wmf. These images
        # are usually figures that are placed in a separate paragraph
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

start5 = time.perf_counter()
print("Editing image placeholders:", start5 - start4)

# By default, docx2python writes the font size as double the original size.
# Need to edit these fonts
for fs in font_sizes:
    text = text.replace(
        f'font-size:{fs}pt',
        f'font-size:{int(fs)//2}pt'
    )

start6 = time.perf_counter()
print("Editing font size:", start6 - start5)

# write a html file
filename = filename.rsplit(".", 1)[0] + ".html"
with open(filename, "w") as html_file:
    html_file.write('<!DOCTYPE html><html lang="en"><head><meta charset='
                    '"UTF-8"><title>Translation</title></head><body>')
    html_file.write(text)
    html_file.write('</body></html>')

end = time.perf_counter()
print("Writing the html file:", end - start6)
