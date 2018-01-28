from xlutils.copy import copy
import sys
import requests
from xlwt import Workbook

try:
    import xlrd
except ImportError:
    print("Please install xlrd package.")
try:
    from xlwt import easyxf
except ImportError:
    print("Please install xlwt package.")

if len(sys.argv) != 2:
    print("Expecting filename. Got: '%s'"
          % ",".join(map(str, sys.argv)))
else:
    fname = sys.argv[1]

workbook = xlrd.open_workbook(fname, encoding_override="cp1252")

sheet = workbook.sheet_by_index(0)
wb = Workbook(encoding='utf-8')
w_sheet = wb.add_sheet('test.xls')
f = open('translation_transcript.txt', 'a')

for i in range(40000, 55975):
    word = sheet.cell(i, 2).value
    try:
        pass
    except Exception as e:
        raise e
    finally:
        pass
    try:
        r = requests.get("https://translate.yandex.net/api/v1.5/tr.json/translate?key=trnsl.1.1.20180123T162008Z.300eb657ff535043.c62444d79cf78f5576dde55e49d0c6fb1473ca3e&text=" + str(word) + "&lang=hr&format=plain")
    except Exception as e:
        r = requests.get("https://translate.yandex.net/api/v1.5/tr.json/translate?key=trnsl.1.1.20180123T162008Z.300eb657ff535043.c62444d79cf78f5576dde55e49d0c6fb1473ca3e&text=" + word.encode('utf-8') + "&lang=hr&format=plain")
    finally:
        pass

    first_bracket = r.content.index("[")+2
    second_bracket = r.content.index("]")-1
    content = r.content

    content = content[first_bracket:second_bracket]

    print content
    w_sheet.write(i, 3, content)
    f.write(content + "\n")
    print "Row:" + str(i) + "; " + str(i*100/55975.0) + "%"

f.close()
wb.save('translated.xls')
