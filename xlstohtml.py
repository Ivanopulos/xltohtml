# для нарезки хтмл-ок по шаблону в эксельке, один столбец-один файл
import pandas  # +openpyxl
import copy
df = pandas.read_excel("1.xlsm", header=None)
df.iloc[0] = df.iloc[0].str.lower()  # хз зачем но очень нада
df.columns = df.iloc[0]
#df = df.fillna("-")
# Функция для замены пустых ячеек на "-"
def fill_empty_cells(row):
    not_empty_count = 0
    for i, cell in enumerate(row):
        if pandas.notna(cell):
            not_empty_count += 1
            if not_empty_count >= 5:
                break
        else:
            not_empty_count = 0
    if not_empty_count >= 5:
        row = row.fillna('-')
    return row
def fill_empty_cells1(row):
    not_empty_count = 0
    for i, cell in enumerate(row):
        if pandas.notna(cell):
            not_empty_count += 1
            if not_empty_count >= 7:
                break
        else:
            not_empty_count = 0
    if not_empty_count >= 7:
        row = row.fillna('-')
    return row
df.iloc[172:294] = df.iloc[172:294].apply(fill_empty_cells, axis=1)
df.iloc[5296:5385] = df.iloc[5296:5385].apply(fill_empty_cells1, axis=1)
print(df.columns.values[1])
colcount = df.count(axis='columns')[0]  # &&
with open("1.html", 'r', encoding='utf-8') as f:  # save before change
    get_all2 = f.readlines()
for fl in range(3, colcount+1):  # перебор столбцов и нарезка их на файлы
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
                try:
                    tx = df[df["{0}"] == found][nstolb].values[0]  # header=False, # ловим глюк разметки - несуществующие строки
                except:
                    print(found)
                    tx = df[df["{0}"] == found][nstolb].values[0]  # header=False,
                try:
                    float(tx)
                    tx = round(tx, 2)
                    tx = str(tx).replace(".", ",")
                except:
                    asd = 0
                try:
                    get_all[isch] = get_all[isch][:usch] + str(tx) + get_all[isch][usch + len(found):]
                except:
                    print(tx)
                    print(found)
                    get_all[isch] = get_all[isch][:usch] + str(tx) + get_all[isch][usch + len(found):]
                found = ""
            usch = usch - 1
        isch = isch + 1
    #/n делит одноименных б
    my_file = open(df.iloc[0][fl], "w",  encoding="utf-8")
    my_file.writelines(get_all)
    my_file.close()
print("FINISH")
