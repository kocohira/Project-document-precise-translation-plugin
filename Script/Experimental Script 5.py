from nltk.tokenize import word_tokenize
import re
from openpyxl import load_workbook          ###Excel写入用
import time



text = '月~金 22:00 ~ 立2：00'

'''text = text.replace('１', '1')
text = text.replace('２', '2')
text = text.replace('３', '3')
text = text.replace('４', '4')
text = text.replace('５', '5')
text = text.replace('６', '6')
text = text.replace('７', '7')
text = text.replace('８', '8')
text = text.replace('９', '9')
text = text.replace('０', '0')

text = text.replace('L', ' L')

text = text.replace('：', ':')

text = text.replace('時', ':')

text = text.replace('分', '')

text = text.replace('/', ' ')

text = text.replace('／', ' ')

text = text.replace('～', '-')

text = text.replace(' ～ ', '-')

text = text.replace('~', '-')

text = text.replace(' ~ ', '-')

text = text.replace(' - ', '-')

text = text.replace(' ー ', '-')

text = text.replace('ー', '-')

text = re.sub("[\!\%\,\〒\,\，\。\(\)\（\）]", " ", text)

text = re.sub("[\u0800-\u9fa5]+", " ", text)







List_Source = word_tokenize(text,"english")



New_List_Source_0 = ['-', ':17:00-29:00', ':12:00-29:00', ':']

New_List_Source = []

for data in New_List_Source_0:

    if data.isdigit() == True  :

        New_List_Source.append(data)



print(New_List_Source)

'''


#a = ':17:00-29:00'

a = ['-', ':fftghgh', ':uiuiu', ':']

New_List_Source_0 = []

for data in a :

    b = re.compile('[0-9]+')

    c = b.findall(data)

    if len(c) != 0 :

        New_List_Source_0.append(data)


print(New_List_Source_0)

###NT
















