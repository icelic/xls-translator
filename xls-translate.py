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

if len(sys.argv) != 6:
    print("Arguments should be: filename translate_from_col translate_to_col translate_from_row translate_to_row. Got: '%s'"
          % ",".join(map(str, sys.argv)))
    sys.exit()
else:
    fname = sys.argv[1]
    source_col = int(sys.argv[2])
    result_col = int(sys.argv[3])
    start_row = int(sys.argv[4])
    end_row = int(sys.argv[5])

workbook = xlrd.open_workbook(fname, encoding_override="cp1252")

sheet = workbook.sheet_by_index(0)
wb = Workbook(encoding='utf-8')
w_sheet = wb.add_sheet('test.xls')
f = open('translation_transcript.txt', 'a')



for i in range(start_row, end_row):
    word = sheet.cell(i, source_col).value
    try:
        r = requests.get("https://translate.yandex.net/api/v1.5/tr.json/translate?key=trnsl.1.1.20180123T162008Z.300eb657ff535043.c62444d79cf78f5576dde55e49d0c6fb1473ca3e&text=" + str(word) + "&lang=hr&format=plain")
    except Exception as e:
        r = requests.get("https://translate.yandex.net/api/v1.5/tr.json/translate?key=trnsl.1.1.20180123T162008Z.300eb657ff535043.c62444d79cf78f5576dde55e49d0c6fb1473ca3e&text=" + word.encode('utf-8') + "&lang=hr&format=plain")
    finally:
        pass

    # this part is custom logic bnecause of structure that yandex returns
    first_bracket = r.content.find("[") + 2
    second_bracket = r.content.find("]") - 1
    content = r.content

    content = content[first_bracket:second_bracket]
    # end logic

    print content
    # writes to separate file because of encoding error
    # TODO make it work in the same file
    w_sheet.write(i, result_col, content)
    f.write(content + "\n")
    print "Row:" + str(i) + "; " \
        + str((i-start_row)*100/float(end_row-start_row)) + "%"

f.close()
wb.save('translated.xls')
