import re
from collections import defaultdict
from heapq import nlargest
from json import loads

from lxml import etree
from requests import get

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.85 Safari/537.36 Edg/90.0.818.49'
}
title: str
author_name: str
author_dict = {}
tag_dict = defaultdict(int)


def get_data(book_name):

    global title, author_name
    author_dict.clear()
    tag_dict.clear()

    # Get possible book list.
    json_str = get('https://book.douban.com/j/subject_suggest?q=' +
                   book_name, headers=headers).content

    # Store data in dictionary.
    for dic in loads(json_str):

        # Skip if the dictionary has no author name or title.
        if 'author_name' not in dic.keys() or 'title' not in dic.keys():
            continue

        # Get title.
        title = dic['title']

        # Get author name.
        author_name = dic['author_name']
        if author_name == '':
            continue
        author_name = re.findall(
            r'(?:[【（(\[][^\x00-\xff]+[）】)\]])?\s*(.*)', author_name)[0]
        author_name = re.sub('著', '', author_name).strip()

        # Get tags.
        get_tags(dic['url'])

        # Store.
        author_dict[author_name] = title

    tag_list = nlargest(5, tag_dict, key=lambda k: tag_dict[k])
    return author_dict, tag_list


def get_tags(url):
    detail_html = get(url=url, headers=headers).text
    detail_tree = etree.HTML(detail_html)
    span_list = detail_tree.xpath('//*[@id="db-tags-section"]/div/span')
    for span in span_list:
        tag = span.xpath('./a/text()')[0]
        if tag == title or tag == author_name:
            continue
        tag_dict[tag] += 1


if __name__ == '__main__':
    while True:
        book = input('>> ').strip()
        print(get_data(book))
