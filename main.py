# INVOICE READER
# Davide Quaranta 2022
import pip

try:
    import pdfplumber as pdf
except:
    print("Error : pdfplumber not installed!")
    ans = input("Do you want to install it [y\n] : ")
    if ans == 'y':
        pip.main(['install', 'pdfplumber'])
        import pdfplumber as pdf
try:
    import os
except:
    print("Error : os not installed!")
    ans = input("Do you want to install it [y\n] : ")
    if ans == 'y':
        pip.main(['install', 'os'])
        import os
try:
  import pandas as pd
except:
    print("Error : pandas not installed!")
    ans = input("Do you want to install it [y\n] : ")
    if ans == 'y':
        pip.main(['install', 'pandas'])
        import pandas as pd

try:
  import openpyxl
except:
    print("Error : openpyxl not installed!")
    ans = input("Do you want to install it [y\n] : ")
    if ans == 'y':
        pip.main(['install', 'openpyxl'])
        import openpyxl


print("                  INVOICE READER                 ")
print("-------------------------------------------------\n")
print("Hi! Welcome to invoice reader!")
check = False
while check == False:
    path = input("Please write the invoice directory : ")
    if os.path.exists(path):
        check = True
    else:
        print("ERROR : Path not Found! Try Again! \n")

# selects only pdf files
files = os.listdir(path)
files = [files[i] for i in range(len(files)) if files[i].endswith('pdf') == True]

# opens a pdf object
open_files = [pdf.open(files[i]) for i in range(len(files))]
pages = [open_files[i].pages for i in range(len(files))]

print(f"{len(files)} pdf found!\n")




words_dict = []
text = ''
# for every page the text is extracted
for k in range(len(pages)):
    for i in range(len(pages[0])):
        text += pages[k][i].extract_text()
    words_dict.append(text)
    text = ''

# we split the text at each endline
lines = []
for i in range(len(words_dict)):
    lines.append(words_dict[i].split(sep = '\n'))

# we separate keys from values in every entry
# and delete lines who don't contain values
split = []
entries = []

for k in range(len(lines)):
    split.append([lines[k][i].split(sep = ':') for i in range(len(lines[k]))])
    entries.append([split[k][i] for i in range(len(split[k])) if len(split[k][i]) == 2])


# we store separately keys and values
# and put them in a dictionary
keys = [list(list(zip(*entries[i]))[0]) for i in range(len(entries))]
values = [list(list(zip(*entries[i]))[1]) for i in range(len(entries))]


new_dict = [{key: value for key,
            value in zip(keys[i], values[i])} for i in range(len(entries))]

# converts dictionaries to dataframe and to excel
df = [pd.DataFrame(new_dict[i], index = [0]) for i in range(len(new_dict))]
final = df[0]
for k in range(len(df)-1):
    final = pd.merge(final,df[k+1], 'outer')

final.to_excel("output.xlsx")
print("Excel file created succesfully! Search for 'output.xlsx' ")
