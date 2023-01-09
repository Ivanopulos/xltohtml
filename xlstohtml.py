#pyinstaller --onefile xltoword.py cd C:\Users\IMatveev\PycharmProjects\wordchanche\

#metka для {}
import re
import os
import pandas  # +openpyxl
import sys
df = pandas.read_excel("1.xlsx", header=None)
df.iloc[0] = df.iloc[0].str.replace(r'\.(?!h)', "", regex=True)
df.iloc[0] = df.iloc[0].str.replace(r'[,/\\\?\!]', "", regex=True)
df.iloc[0] = df.iloc[0].str.lower()
df.columns = df.iloc[0]
print(df.columns.values[1])

colcount = df.count(axis='columns')[0]   # один столбец под разметку один под названия
#pathword = df.iloc[0][tm] #адрес ворда
with open("1.html", 'r', encoding='utf-8') as f:  # save before chenge
    get_all2 = f.readlines()
    f.close()
for fl in range(2, colcount):
    nstolb = df.columns.values[fl]
    print(df.iloc[0][fl])
    mem = 0
    isch = 0
    usch = 0
    found = ""
    get_all = get_all2
    #with open(df.iloc[0][fl], 'w', encoding='utf-8') as f:  # look for { and chenge it
    for i in get_all:         # STARTS THE NUMBERING FROM 1 (by default it begins with 0)
        usch = len(i)-1
        for u in i:
            if get_all[isch][usch] == "}":
                mem = 1
            if mem == 1:
                found = get_all[isch][usch] + found
            if get_all[isch][usch] == "{":
                mem = 0
                print(found)
                print(isch)
                print(usch)
                tx = df[df["{0}"] == found][nstolb].values[0]  # header=False,

                try:
                    float(tx)
                    tx = str(tx).replace(".", ",")
                except:
                    asd = 0
                get_all[isch] = get_all[isch][:usch] + tx + get_all[isch][usch + len(found):]
                found = ""
            usch = usch - 1
        isch = isch + 1
        #f.writelines(get_all)
    my_file = open(df.iloc[0][fl], "w")
    my_file.writelines(get_all)
    my_file.close()
    #with open(df.iloc[0][fl], 'r', encoding='utf-8') as f:  # save before chenge
    #    f.writelines(get_all)


# print("XML chanched")
# try:
#     os.remove(pathwork + "/B.zip")
# except:
#     asd = 1

print("zip saved")
name = str(df.iloc[0, 1])
name = "О разв ПМСП в " + name[82:]
# try:
#     os.remove(pathwork + "/" + name + ".docx")
# except:
#     asd = 1
print("FINISH")
