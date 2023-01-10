import pandas  # +openpyxl
import copy
df = pandas.read_excel("1.xlsx", header=None)
df.iloc[0] = df.iloc[0].str.replace(r'\.(?!h)', "", regex=True)
df.iloc[0] = df.iloc[0].str.replace(r'[,/\\\?\!]', "", regex=True)
df.iloc[0] = df.iloc[0].str.lower()
df.columns = df.iloc[0]
df = df.fillna("-")
print(df.columns.values[1])
colcount = df.count(axis='columns')[0]   # один столбец под разметку один под названия
with open("1.html", 'r', encoding='utf-8') as f:  # save before chenge
    get_all2 = f.readlines()
for fl in range(2, colcount):
    nstolb = df.columns.values[fl]
    print(nstolb)
    mem = 0
    isch = 0
    usch = 0
    found = ""
    get_all = copy.deepcopy(get_all2)
    for i in get_all:         # STARTS THE NUMBERING FROM 1 (by default it begins with 0)
        usch = len(i)-1
        for u in i:
            if get_all[isch][usch] == "}":
                mem = 1
            if mem == 1:
                found = get_all[isch][usch] + found
            if get_all[isch][usch] == "{":
                mem = 0
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
    #/n делит одноименных б
    my_file = open(df.iloc[0][fl], "w",  encoding="utf-8")
    my_file.writelines(get_all)
    my_file.close()
print("FINISH")
