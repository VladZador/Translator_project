import requests
import re
from docx2python import docx2python
from bs4 import BeautifulSoup


filename = "Test.docx"

# todo: add font size, languages as variables
FONT_SIZE = 14

# open the document
docx_content = docx2python(filename, html=True)

# save images
docx_content.save_images('media')

# parse through the content
soup = BeautifulSoup(docx_content.text, "html.parser")

# break text into paragraphs and discards empty ones
paragraphs = [p for p in soup.get_text().splitlines() if p]

# discard the image placeholders in text
text_blocks = set([])
for p in paragraphs:
    if "----media/" not in p:
        text_blocks.add(p)
    else:
        text_blocks.update([block for block in p.split("----") if "media/image"
                            not in block and not block.isascii()])

# translate the text by sending its parts to the Google Translator API and
# create translation table
trans_table = []
for block in text_blocks:
    url = f"https://translate.googleapis.com/translate_a/single?client=gtx" \
          f"&sl=uk&tl=en&dt=t&q={block}&ie=UTF-8&oe=UTF-8"
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


# Find occurrences of a pattern like "----media/image1.png----
#     Match characters "----media/" and then "----" literally
#     Group (image\d\..{3}):
#         "image" matches literally
#         "\d" matches digits
#             "\d+" matches digits, one or more occurrences
#         "\." matches "."
#         ".{3}" matches exactly 3 any characters
regex_for_image = re.compile(r'----media/(image\d+\..{3})----')
image_placeholders = set(regex_for_image.findall(text))

# Edit image placeholders
for img_pl in image_placeholders:
    # First, search for images inside text blocks (with incorrect <span> tags)
    if f'</span>----media/{img_pl}----<span style="font-size:28pt">' in text:
        text = text.replace(
            f'</span>----media/{img_pl}----<span style="font-size:28pt">',
            f'<foo_bar_baz><img src="media/{img_pl}" alt="{img_pl}">'
            f'</foo_bar_baz>')
    # Second, for images with an extension other than .wmf, which is used for
    # formulas in MS Word. These images are usually figures that are placed in
    # a separate paragraph
    elif not img_pl.endswith(".wmf"):
        text = text.replace(
            f'----media/{img_pl}----',
            f'<p><img src="media/{img_pl}" alt="{img_pl}"></p>')
    # Third, other .wmf images
    else:
        text = text.replace(
            f'----media/{img_pl}----',
            f'<foo_bar_baz><img src="media/{img_pl}" alt="{img_pl}">'
            f'</foo_bar_baz>')

# Edit <span> html-tag to <p>
text = text.replace("<span", "<p").replace("span>", "p>")

# Edit dummy <foo_bar_baz> tags to <span>
text = text.replace("foo_bar_baz>", "span>")


# write a html file
filename = filename.rsplit(".", 1)[0] + ".html"
with open(filename, "w") as html_file:
    html_file.write('<!DOCTYPE html><html lang="en"><head><meta charset='
                    '"UTF-8"><title>Translation</title></head><body>')
    html_file.write(text)
    html_file.write('</body></html>')

# close the document
docx_content.close()
