from docx2python import docx2python

# open the document
docx_content = docx2python("Test.docx", html=True)

# save images
docx_content.save_images('media')

# Edit <span> html-tag to <p>
content = docx_content.text
content.replace("<span", "<p").replace("span>", "p>")

# Edit text in order to include <img> tag where images should be.
new_content = ""
for elem in content.split("----"):
    if "media/image" in elem:
        elem = elem.replace("media/image", "")
        elem = f'<img src="media/image{elem}" alt="{elem}">'
    new_content += elem


# write a html file
with open("Test.html", "w") as html_file:
    html_file.write(new_content)

# close the document
docx_content.close()
