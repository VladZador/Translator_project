import requests
import time
from docx2python import docx2python

# open the document
docx_content = docx2python("Test.docx", html=True)

# save images
docx_content.save_images('media')


# translate the text by sending its parts to the Google Translator API
translated_content = ""
for content_part in docx_content.body[0][0][0]:
    if content_part.isascii():
        translated_content += content_part
    else:
        url = f"https://translate.googleapis.com/translate_a/single?client=gtx&sl=uk&tl=en&dt=t&q={content_part}&ie=UTF-8&oe=UTF-8"
        r = requests.get(url, timeout=(5, 5))
        if r.json()[0]:
            for item in r.json()[0]:
                if item[0]:
                    translated_content += item[0]
        time.sleep(0.1)


# Correct extra whitespaces inside of tags
content = translated_content.replace("< /", "</").replace("< s", "<s").replace("< p", "<p").replace("< b", "<b")

# Edit <span> html-tag to <p>
content = content.replace("<span", "<p").replace("span>", "p>")


# Edit text in order to include <img> tag where images should be.
content_with_images = ""
for elem in content.split("----"):
    if "media/image" in elem:
        elem = elem.replace("media/image", "")
        elem = f'<span><img src="media/image{elem}" alt="{elem}"></span>'
    content_with_images += elem


# write a html file
with open("Test.html", "w") as html_file:
    html_file.write('<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">'
                    '<title>Translation</title></head><body>')
    html_file.write(content_with_images)
    html_file.write('</body></html>')

# close the document
docx_content.close()
