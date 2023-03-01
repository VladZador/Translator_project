import json
import time
import requests
import re
from docx2python import docx2python
from bs4 import BeautifulSoup
import os
from dotenv import load_dotenv

load_dotenv()


def _open_doc(file_name: str, image_folder: str = None, html=False) -> str:
    """
    Open the document, save images and write the text into a variable
    :param file_name:
    :param image_folder:
    :param html:
    :return:
    """
    with docx2python(
            os.environ.get("UPLOAD_FOLDER") + file_name,
            image_folder=image_folder,
            html=html
    ) as docx_content:
        doc_text = docx_content.text
    return doc_text


def _parse_through_html(html_content: str) -> str:
    """
    Parse through a given html content
    :param html_content: string of a html content
    :return: parsed text
    """
    soup = BeautifulSoup(html_content, "html.parser")
    return soup.get_text()


def _break_into_paragraphs(text: str) -> list:
    """
    Break the text into paragraphs and discards empty ones
    :param text:
    :return:
    """
    paragraphs = [p for p in text.splitlines() if p]
    return paragraphs


def _clean_from_image_placeholders(par_list: list) -> set:
    """
    Clean the paragraphs from image placeholders. isascii() method is used in
    order to discard empty strings or those containing punctuation marks or
    words in English.
    :param par_list:
    :return:
    """
    text_blocks = set([])
    for p in par_list:
        if f"----media/" not in p and not p.isascii():
            text_blocks.add(p)
        else:
            text_blocks.update(
                [block for block in p.split("----")
                 if f"media/image" not in block and not block.isascii()]
            )
    return text_blocks


def _break_into_sentences(text_iterable) -> set:
    """
    Break text into sentences
    :param text_iterable: list or set containing strings
    :return:
    """
    sentences = set([])
    for block in text_iterable:
        sentences.update(
            [s for s in block.split(".") if not s.isascii()]
        )
    return sentences


def _create_an_html_text(text: str) -> str:
    return "".join(f"<p>{p}</p>" for p in text.splitlines() if p)


def _translate_block_with_google(string) -> str:
    """
    Translate the text by sending it to the Google Translator API
    :param string:
    :return:
    """
    url = f"https://translate.googleapis.com/translate_a/single?client=gtx" \
          f"&sl={LANG_IN}&tl={LANG_OUT}&dt=t&q={string}&ie=UTF-8&oe=UTF-8"
    req = requests.get(url)
    response = ""
    if req.json()[0]:
        for item in req.json()[0]:
            if item[0]:
                response += item[0]
    return response


def _make_trans_table_with_google(text_blocks) -> list:
    """
    Translates the text and creates translation table
    :param text_blocks:
    :return:
    """
    translation_table = []
    for s in sorted(text_blocks, key=len, reverse=True):
        translated_str = _translate_block_with_google(s)
        translation_table.append((s, f" {translated_str} "))
    return translation_table


def _translate_text(text: str, translation_table: list) -> str:
    """
    Replace strings using the translation table
    :param text:
    :param translation_table:
    :return:
    """
    for trans_tuple in translation_table:
        text = text.replace(trans_tuple[0], trans_tuple[1])
    return text


def _change_span_and_p_tags(given_text: str) -> str:
    """
    Edit <span> html-tag to <p> in order to fix the problem with tags inside
    a text parsed by docx2python module. This module places paragraphs
    inside a <span> tag, that is not very great.
    :param given_text:
    :return:
    """
    return given_text.replace("<span", "<p").replace("span>", "p>")


def _edit_image_placeholders(html_str: str, image: str, size: str) -> str:
    # First, search for images inside text blocks, usually it's the wmf
    # images, which is used as formulas in MS Word
    if image.endswith(".wmf") and \
            f'</p>----media/{image}----<p style="font-size:{size}pt">' \
            in html_str:
        html_str = html_str.replace(
            f'</p>----media/{image}----<p style="font-size:{size}pt">',
            f'<span><img src="static/media/{image}" alt="{image}"></span>'
        )
    # Second, for images with an extension other than wmf. These images
    # are usually figures that are placed in a separate paragraph
    elif not image.endswith(".wmf"):
        html_str = html_str.replace(
            f'----media/{image}----',
            f'<p><img src="static/media/{image}" alt="{image}"></p>'
        )
    # Third, wmf images that are placed outside of tags for some reason
    elif f'</p>----media/{image}----' in html_str:
        html_str = html_str.replace(
            f'</p>----media/{image}----',
            f'<span><img src="static/media/{image}" alt="{image}"></span></p>'
        )
    # Fourth, some other images that were not covered by previous clauses
    else:
        html_str = html_str.replace(
            f'----media/{image}----',
            f'<span><img src="static/media/{image}" alt="{image}"></span>'
        )
    return html_str


def _edit_images_and_fonts(html_content: str) -> str:
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
    font_sizes = set(regex_for_font.findall(html_content))

    # Edit image placeholders
    for img_pl in regex_for_image.findall(html_content):
        img_pl = img_pl[0]
        for fs in font_sizes:
            html_content = _edit_image_placeholders(html_content, img_pl, fs)

    # By default, docx2python writes the font size as double the original size.
    # Need to edit these fonts
    for fs in font_sizes:
        html_content = html_content.replace(
            f'font-size:{fs}pt',
            f'font-size:{int(fs)//2}pt'
        )
    return html_content


def _write_html_file(file_name, text, simple=True):
    """
    Write a html file
    :param file_name:
    :param text:
    :return:
    """
    suffix = "" if simple else "_complex"
    file_name = "templates/" + file_name.rsplit(".", 1)[0] + suffix + ".html"

    with open(file_name, "w") as html_file:
        html_file.write('<!DOCTYPE html><html lang="en"><head><meta charset='
                        '"UTF-8"><title>Translation</title></head><body>')
        html_file.write(text)
        html_file.write('</body></html>')

    return html_file


def _translate_as_html(file_name):
    start1 = time.perf_counter()

    docx_text = _open_doc(file_name, image_folder="static/media", html=True)

    start2 = time.perf_counter()
    print("Opening the document, saving images:", start2 - start1)

    parsed_text = _parse_through_html(docx_text)
    paragraph_list = _break_into_paragraphs(parsed_text)
    text_set = _clean_from_image_placeholders(paragraph_list)
    text_set = _break_into_sentences(text_set)

    start3 = time.perf_counter()
    print("Breaking the text into segments:", start3 - start2)

    trans_table = _make_trans_table_with_google(text_set)

    with open("translation_table.json", "w") as json_file:
        json_file.write(json.dumps(dict(trans_table)))

    # Prepare the text for editing
    docx_text = docx_text.replace("\n", "")

    trans_text = _translate_text(docx_text, trans_table)

    start4 = time.perf_counter()
    print("Translation:", start4 - start3)
    print(f"There were {len(text_set)} phrases to translate, average time "
          f"is {(start4 - start3)/len(text_set)} for each phrase")

    html_text = _change_span_and_p_tags(trans_text)
    html_text = _edit_images_and_fonts(html_text)

    start5 = time.perf_counter()
    print("Editing images and font size:", start5 - start4)

    file = _write_html_file(file_name, html_text, simple=False)

    return file


def _translate_as_text(file_name):
    start1 = time.perf_counter()

    docx_text = _open_doc(file_name)

    start2 = time.perf_counter()
    print("Opening the document", start2 - start1)

    paragraph_list = _break_into_paragraphs(docx_text)

    start3 = time.perf_counter()
    print("Breaking the text into segments:", start3 - start2)

    trans_table = _make_trans_table_with_google(paragraph_list)

    with open("translation_table.json", "w") as json_file:
        json_file.write(json.dumps(dict(trans_table)))

    # Prepare the text for editing
    # docx_text = _create_an_html_text(docx_text)
    trans_text = _translate_text(docx_text, trans_table)

    start4 = time.perf_counter()
    print("Translation:", start4 - start3)
    print(f"There were {len(paragraph_list)} phrases to translate, average "
          f"time is {(start4 - start3)/len(paragraph_list)} for each phrase")

    file = _write_html_file(file_name, trans_text)

    return file


def translate(file_name, simple=True):
    if simple:
        return _translate_as_text(file_name)
    return _translate_as_html(file_name)


LANG_IN = "uk"
LANG_OUT = "en"


if __name__ == "__main__":
    test_filename = "Test.docx"
    translated_text = translate(test_filename)
    # docx_text = _open_doc(
    #     test_filename, image_folder="static/media", html=True
    # )
    # _write_html_file(test_filename, docx_text)


# todo: remove writing a json file and "prints", used for debugging
