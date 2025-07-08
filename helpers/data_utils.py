import sys  # nopep8
sys.path.insert(0, './lib')  # nopep8

import re
from bs4 import BeautifulSoup


def fix_incomplete_links(text):
    # أي رابط مش مسبوق بـ http نحوله
    words = text.split()
    fixed_words = []
    for word in words:
        if (
            re.match(r'^(www\.|[a-zA-Z0-9-]+\.[a-z]{2,})', word) and
            not word.startswith(('http://', 'https://'))
        ):
            fixed_words.append('http://' + word)
        else:
            fixed_words.append(word)
    return ' '.join(fixed_words)


def extract_text_and_image(description: str):
    """Extract cleaned text and first image, fixing raw links."""
    soup = BeautifulSoup(description, 'html.parser')

    lines = []
    for div in soup.find_all('div'):
        if not div.text.strip() and div.find('br'):
            lines.append('')
        else:
            line_parts = []
            for elem in div.descendants:
                if isinstance(elem, str):
                    text = elem.replace('\u200b', '').replace(
                        '\xa0', ' ').strip()
                    if text:
                        line_parts.append(text)
                elif elem.name == 'a' and elem.has_attr('href'):
                    href = elem['href'].strip()
                    visible = elem.get_text(strip=True)
                    if visible and visible != href:
                        line_parts.append(visible)
                    else:
                        line_parts.append(href)

            # إزالة التكرار داخل السطر
            line = ' '.join(dict.fromkeys(line_parts))
            line = fix_incomplete_links(line)
            if line:
                lines.append(line.strip())

    cleaned_text = '\n'.join(lines)

    # الصورة
    img_tag = soup.find('img')
    image = img_tag['src'].strip(
    ) if img_tag and img_tag.has_attr('src') else ''

    return cleaned_text, image
