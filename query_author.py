from requests import get
from json import loads
import re


def get_author(book_name):

    # Get the content of the search request.
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.85 Safari/537.36 Edg/90.0.818.49'
    }
    json_str = get('https://book.douban.com/j/subject_suggest?q=' +
                   book_name, headers=headers).content

    # Store possible authors in a list.
    recommended_dict = {}
    for dic in loads(json_str):

        # Skip if the dictionary has no author name or title.
        if 'author_name' not in dic.keys() or 'title' not in dic.keys():
            continue

        title = dic['title']

        # Clean author name.
        author_name = re.findall(
            r'(?:[【（(\[][^\x00-\xff]+[）】)\]])?\s*(.*)', dic['author_name'])[0]
        author_name = re.sub('著', '', author_name).strip()

        # Store.
        if author_name != '':
            recommended_dict[author_name] = title

    return recommended_dict


if __name__ == '__main__':
    while True:
        book = input('>> ').strip()
        print(get_author(book))
