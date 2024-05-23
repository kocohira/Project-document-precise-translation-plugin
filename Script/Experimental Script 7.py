from nltk.tokenize import word_tokenize
import re
from openpyxl import load_workbook          ###Excel写入用
import time
import os
import sys
import goto
from dominate.tags import label
from goto import with_goto


@with_goto
def Total() :

    label.begin



    text = '''12:00 - 15:00
    18:00 - 04:00'''


    List_Source = word_tokenize(text, "english")

    print(List_Source)


    print (181//10)

    a = '+853 2888 2866'

    a = a[5:]

    a = a.replace(' ','')

    print (a)

    time.sleep(2)

    goto.begin



Total()
###NT