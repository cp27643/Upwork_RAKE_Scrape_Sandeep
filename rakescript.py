from rake_nltk import Rake
from pprint import pprint
import xlwings as xw
import os, string
from bs4 import BeautifulSoup
import re, csv


def CountOccurrences(my_text):
    word_dict = {}
    end = False
    i = 0
    exclude = set(string.punctuation)
    exclude.add("\'")
    while not end:
        nextspace = str.find(my_text, ' ', i + 1)
        if nextspace < 0:
            end = True
            return word_dict
        cur_word = my_text[i:nextspace]
        cur_word = ''.join(ch for ch in cur_word if ch not in exclude)
        cur_word = str.lower(cur_word)
        i = nextspace + 1
        if cur_word in word_dict.keys():
            cur_count = word_dict[cur_word]
            cur_count += 1
            word_dict[cur_word] = cur_count  # Increase word occurence by 1
        else:  # Word hasnt been put in dictionary yet
            word_dict[cur_word] = 1




def RankWords(mytext):
    """

    :param mytext: should be a string with the text wanting to be analyized
    :return: going to return two dictionaries, one with the word and it's frequency, the other with a word and its RAKE score.
    """
    r = Rake()
    r.extract_keywords_from_text(text=mytext)
    ranked_phrases = r.get_ranked_phrases_with_scores()
    return ranked_phrases


def PrintOutput():
    try:
        wb = xw.Book(os.path.join(os.curdir, r'RakeResults.xlsx'))
    except:
        wb = xw.Book()  # Book doesnt exist yet
    cur_sheet = wb.sheets['Sheet1']  # Selecting current sheet
    cur_sheet.range(r'A1').value = 'Python RAKE Results'
    cur_sheet.range(r'B1').value = 'Frequency Count'
    cur_sheet.range(r'D1').value = 'RAKE Word'
    cur_sheet.range(r'E1').value = 'RAKE Word Rank'

    for i, item in enumerate(frequency_dict.items()):
        print(item)
        xw.Range(r'A' + str(i + 2)).value = [item]

    for i, item in enumerate(rank_dict):
        xw.Range(r'D' + str(i + 2)).value = item[1], item[0]
    wb.save(os.path.join(os.curdir, r'RakeResults.xlsx'))


def visible_texts(soup):
    """ get visible text from a document """
    text = ' '.join([
        s for s in soup.strings
        if s.parent.name not in INVISIBLE_ELEMS
    ])
    # collapse multiple spaces to two spaces.
    return RE_SPACES.sub('  ', text)


frequency_dict = {}
rank_dict = {}
INVISIBLE_ELEMS = ('style', 'script', 'head', 'title')
RE_SPACES = re.compile(r'\s{3,}')

# for dirpath, dirnames, filenames in os.walk(r'C:\Users\c1phill\PycharmProjects\Rake Project\HT\www.hindustantimes.com'):
#     pass
webtext = []
for root, dirs, files in os.walk(r'C:\Users\c1phill\PycharmProjects\Rake Project\HT\www.hindustantimes.com'):
    path = root.split(os.sep)
    #print((len(path) - 1) * '---', os.path.basename(root))
    for file in files:
        if file[-5:] == '.html':
            try:
                soup = BeautifulSoup(open(os.path.join(root,file), mode='rb'), 'html5lib')
            except UnicodeDecodeError:
                print('UnicodeDecodeError')
            # ranktext  = RankWords(visible_texts(soup))
            # count_text = CountOccurrences(visible_texts(soup))
            include = (' ')
            paragraphs = soup.find_all('p')
            paragraphs_text = r''
            for paragraph in paragraphs:
                texttoadd = ''.join(ch for ch in paragraph.text if str.isalpha(ch) or ch in include)
                paragraphs_text += r' ' + texttoadd

            ranktext = RankWords(paragraphs_text)
            count_text = CountOccurrences(paragraphs_text)
            with open(os.path.join(os.curdir, r'RAKEOutput.csv'), mode = 'a', newline = '') as my_csv:
                startpath = path.index(r'www.hindustantimes.com')
                startpath = str('/'.join(path[startpath:])+'/'+file)
                csv.writer(my_csv).writerow([startpath])
                #csv.writer(my_csv).writerow([r'Word', r'Frequency'])
                # for row in count_text:#Count_text is a dictionary
                #     try:
                #         csv.writer(my_csv).writerow([row, count_text[row]])
                #     except:
                #         pass
                # csv.writer(my_csv).writerow([r''])
                csv.writer(my_csv).writerow([r'Word', 'Rake Ranking'])
                for row in ranktext:
                    try:
                        csv.writer(my_csv).writerow([row[1], row[0]])
                    except:
                        pass
                # csv.writer(my_csv).writerow([r''])
            my_csv.close()
            print(len(path) * '---', file)

