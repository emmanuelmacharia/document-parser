import os
import re
from dateutil.parser import parse
from collections.abc import Iterable
import lxml.html
import csv

import docx
from docx2python import docx2python


def flatten(l):
    for el in l:
        if isinstance(el, Iterable) and not isinstance(el, (str, bytes)):
            yield from flatten(el)
        else:
            yield el

def is_date(string, fuzzy=False):
    """
    Return whether the string can be interpreted as a date.

    :param string: str, string to check for date
    :param fuzzy: bool, ignore unknown tokens in string if True
    """
    try: 
        parse(string, fuzzy=fuzzy)
        return True

    except ValueError:
        return False

def absolute_file_paths(directory):
    path = os.path.abspath(directory)
    my_files = [entry.path for entry in os.scandir(path) if entry.is_file() and (entry.name.endswith('docx') or entry.name.endswith('DOCX'))]
    return my_files

def get_doc(fileName, params):
    doc = docx.Document(fileName)
    use_docx2python(fileName, params)
    paragraphs = []
    for para in doc.paragraphs:
        if  para.text != '' and para.text != ' ':
            paragraphs.append(para)
    return paragraphs

def use_docx2python(fileName, params):
    ''' gets our title and the place, if theres any'''
    parsed = docx2python(fileName)
    heading = parsed.body
    body = flatten(heading)
    title = ''
    place = ''
    for i in body:
        if '<a href=' in i:
            title += i

    if title:
        title = lxml.html.fromstring(title).text_content()
    else:
       pass

    if '[' in title:
        place += re.split('\[', title)[-1]


    params["headline"] = title
    params["place"] = place[:-1]



def get_word_from_arguments(file_path) -> str:
    files = absolute_file_paths(file_path)
    params_array = []
    for file in files:
        parameters = {
            'headline': '',
            'place': '',
            'date':'',
            'day': "",
            'words': '',
            'byline': ''
        }
        paragraphs = get_doc(file, parameters)
        everything_found = False
        while not everything_found:
            for para in paragraphs:
                if is_date(para.text.rsplit(' ', 1)[0]):
                    parameters['date'] = para.text.rsplit(' ', 1)[0]
                    parameters['day'] = para.text.split(' ')[-1]
                if 'length:' in para.text.lower():
                    parameters['words'] =  re.split(' ', re.split(':', para.text.replace(u'\xa0', u''))[1])[0]
                if 'byline:' in para.text.lower():
                    parameters['byline'] = re.split(':', para.text.replace(u'\xa0', u''))[1]
                if  is_everything_found(parameters):
                    break
            everything_found = True
        params_array.append(parameters)
    print(params_array)
    generate_output_file(params_array)



def is_everything_found(params: dict) -> bool :
    filled = all(value for value in params.values())
    if filled:
        everything_found = True
        return True
    return False



def generate_output_file(parameters):
    fileName =  csv_file_name = parameters[0]['date'].split(' ')[-1] + '.csv'
    with open(fileName, 'w', newline='\n') as file:
        fieldNames = list(parameters[0].keys())
        writer = csv.DictWriter(file, fieldnames=fieldNames)
        writer.writeheader()
        for param in parameters:
            writer.writerow(param)



# resume = get_doc('F:\duplicates\10 Years On - How 9_11 Changed the World [column].DOCX')

get_word_from_arguments(r'F:\2012')

