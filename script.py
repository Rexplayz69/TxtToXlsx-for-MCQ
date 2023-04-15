import xlsxwriter as Excel
import tkinter as Tk
from tkinter import filedialog as TkFileDialog
import tkinter.font as TkFont
from tkinter import messagebox as TkMsgBox

SPLIT_CHAR: str = '~;~'

fileName: str = ''
filePath: str = ''
fileData: list[str] = []


def selectTxtFile():
    global fileName, filePath

    filePath = TkFileDialog.askopenfilename(title='Pick a Text File', filetypes=(('Text Files', '*.txt',),))
    if filePath.endswith('.txt'):
        fileName = filePath.split('/')[-1][:-4]
        fileNameVar.set(f"Selected Text File: {fileName}")
        pickBtnVar.set('Repick Text File')

def getAnsFromOptions(optionsList: list[str]) -> str:
    if optionsList[4].strip() == optionsList[0].strip():
        return 'A'
    elif optionsList[4].strip() == optionsList[1].strip():
        return 'B'
    elif optionsList[4].strip() == optionsList[2].strip():
        return 'C'
    elif optionsList[4].strip() == optionsList[3].strip():
        return 'D'
    else:
        return ''


#############
# Question? #
# A~;~      #
# B~;~      #
# C~;~      #
# D~;~      #
# Ans~;~    #
# Diff~;~   #
#############
def isPattern1(index: int) -> bool:
    global fileData

    if (index + 5) >= len(fileData):
        return False

    if fileData[index].count(SPLIT_CHAR) != 1:
        return False

    if fileData[index + 1].count(SPLIT_CHAR) != 1:
        return False

    if fileData[index + 2].count(SPLIT_CHAR) != 1:
        return False

    if fileData[index + 3].count(SPLIT_CHAR) != 1:
        return False

    if fileData[index + 4].count(SPLIT_CHAR) != 1:
        return False

    if fileData[index + 5].count(SPLIT_CHAR) != 1:
        return False

    return True

#############
# Question? #
# A~;~      #
# B~;~      #
# C~;~      #
# D~;~      #
# A~;~ Diff #
#############
def isPattern2(index: int) -> bool:
    global fileData

    if index + 4 >= len(fileData):
        return False

    if fileData[index].count(SPLIT_CHAR) != 1:
        return False

    if fileData[index + 1].count(SPLIT_CHAR) != 1:
        return False

    if fileData[index + 2].count(SPLIT_CHAR) != 1:
        return False

    if fileData[index + 3].count(SPLIT_CHAR) != 1:
        return False

    if fileData[index + 4].count(SPLIT_CHAR) >= 1:
        return True

    return False

####################
# Question?        #
# A~;~ B~;~ C~;~ D #
# Ans~;~ Diff      #
####################
def isPattern3(index: int) -> bool:
    global fileData

    if (index + 1) >= len(fileData):
        return False

    if fileData[index].count(SPLIT_CHAR) < 3:
        return False

    if fileData[index + 1].count(SPLIT_CHAR) >= 1:
        return True

    return False

#################################
# Question?                     #
# A~;~ B~;~ C~;~ D~;~ A~;~ Diff #
#################################
def isPattern4(index: int) -> bool:
    global fileData

    if index >= len(fileData):
        return False

    if fileData[index].count(SPLIT_CHAR) >= 5:
        return True

    return False

def convert():
    if filePath == '':
        return

    global fileData

    with open(filePath, 'r', encoding="utf8") as txtFile:
        fileData = txtFile.readlines()

    while fileData.__contains__('\n'):
        fileData.remove('\n')

    numQues: int = 0
    for data in fileData:
        if not data.__contains__(SPLIT_CHAR):
            numQues += 1

    tempDict: dict = {}
    quesMapList: list[dict] = []

    index: int = 0
    dupQues: int = 0

    while index < len(fileData):
        skip: bool = False
        for ques in quesMapList:
            if ques['ques'] == fileData[index].strip():
                if isPattern1(index + 1):
                    skip = True
                    dupQues += 1
                    index += 7
                    break
                elif isPattern2(index + 1):
                    skip = True
                    dupQues += 1
                    index += 6
                    break
                elif isPattern3(index + 1):
                    skip = True
                    dupQues += 1
                    index += 3
                    break
                elif isPattern4(index + 1):
                    skip = True
                    dupQues += 1
                    index += 2
                    break
                else:
                    skip = True
                    print(f'Error for duplicate Question on {index}: {fileData[index]}')
                    index += 1
                    break

        if skip:
            continue

        tempDict.update({'ques': fileData[index].strip()})

        if isPattern1(index + 1):
            tempDict.update({'a': fileData[index + 1].split(SPLIT_CHAR)[0].strip()})
            tempDict.update({'b': fileData[index + 2].split(SPLIT_CHAR)[0].strip()})
            tempDict.update({'c': fileData[index + 3].split(SPLIT_CHAR)[0].strip()})
            tempDict.update({'d': fileData[index + 4].split(SPLIT_CHAR)[0].strip()})
            tempDict.update({'ans': fileData[index + 5].split(SPLIT_CHAR)[0].strip()})
            tempDict.update({'diff': fileData[index + 6].split(SPLIT_CHAR)[0].strip()})

            quesMapList.append(tempDict.copy())
            tempDict.clear()
            index += 7
        elif isPattern2(index + 1):
            tempDict.update({'a': fileData[index + 1].split(SPLIT_CHAR)[0].strip()})
            tempDict.update({'b': fileData[index + 2].split(SPLIT_CHAR)[0].strip()})
            tempDict.update({'c': fileData[index + 3].split(SPLIT_CHAR)[0].strip()})
            tempDict.update({'d': fileData[index + 4].split(SPLIT_CHAR)[0].strip()})

            splitList: list[str] = fileData[index + 5].split(SPLIT_CHAR)
            tempDict.update({'ans': splitList[0].strip()})
            tempDict.update({'diff': splitList[1].strip()})

            quesMapList.append(tempDict.copy())
            tempDict.clear()
            index += 6
        elif isPattern3(index + 1):
            splitList: list[str] = fileData[index + 1].split(SPLIT_CHAR)
            tempDict.update({'a': splitList[0].strip()})
            tempDict.update({'b': splitList[1].strip()})
            tempDict.update({'c': splitList[2].strip()})
            if splitList[3].strip().endswith('.'):
                tempDict.update({'d': splitList[3].strip()[:-1]})
            else:
                tempDict.update({'d': splitList[3].strip()})

            splitList: list[str] = fileData[index + 2].split(SPLIT_CHAR)
            tempDict.update({'ans': splitList[0].strip()})

            if splitList[1].strip().endswith('.'):
                tempDict.update({'diff': splitList[1].strip()[:-1]})
            else:
                tempDict.update({'diff': splitList[1].strip()})

            quesMapList.append(tempDict.copy())
            tempDict.clear()
            index += 3
        elif isPattern4(index + 1):
            splitList: list[str] = fileData[index + 1].split(SPLIT_CHAR)

            tempDict.update({'a': splitList[0].strip()})
            tempDict.update({'b': splitList[1].strip()})
            tempDict.update({'c': splitList[2].strip()})
            tempDict.update({'d': splitList[3].strip()})

            if splitList[4].strip() in ['A', 'B', 'C', 'D']:
                tempDict.update({'ans': splitList[4].strip()})
            else:
                tempDict.update({'ans': getAnsFromOptions(splitList[:5])})

            tempDict.update({'diff': splitList[5].strip()})

            quesMapList.append(tempDict.copy())
            tempDict.clear()
            index += 2
        else:
            print(f'Error while reading on {index}: {fileData[index]}')
            index += 1

    numQuesWritten = len(quesMapList)

    excelWorkbook: Excel.Workbook = Excel.Workbook(f"{filePath[:-4]}.xlsx")
    excelSheet: Excel.workbook.Worksheet = excelWorkbook.add_worksheet()

    for i in range(0, len(quesMapList)):
        style: Excel.workbook.Format = excelWorkbook.add_format({'font_color': 'black'})
        if quesMapList[i]['diff'] == 'Easy':
            style = excelWorkbook.add_format({'font_color': '#00B050'})
        elif quesMapList[i]['diff'] == 'Medium':
            style = excelWorkbook.add_format({'font_color': '#FC000'})
        elif quesMapList[i]['diff'] == 'Hard':
            style = excelWorkbook.add_format({'font_color': '#FF0000'})
        elif quesMapList[i]['diff'] == 'Extreme':
            style = excelWorkbook.add_format({'font_color': '#7030A0'})
        else:
            print(f'Color Error on: {quesMapList[i]["ques"]}')

        excelSheet.write(i, 0, quesMapList[i]['ques'], style)
        excelSheet.write(i, 1, quesMapList[i]['a'], style)
        excelSheet.write(i, 2, quesMapList[i]['b'], style)
        excelSheet.write(i, 3, quesMapList[i]['c'], style)
        excelSheet.write(i, 4, quesMapList[i]['d'], style)
        excelSheet.write(i, 5, quesMapList[i]['ans'], style)

    excelWorkbook.close()

    quesMissed: bool = True
    if (numQues - dupQues) == numQuesWritten:
        quesMissed = False

    TkMsgBox.showinfo(title='Converted Successfully',
                      message=f'Question in Text File: {numQues}\nDuplicate Question: {dupQues}\nQuestion written in Excel File: {numQuesWritten}\nSome Questions missed: {quesMissed}\n\nExcel File is Saved in the same Directory as the Text File')


win: Tk.Tk = Tk.Tk()
win.title('Text To Excel')
win.iconphoto(True, Tk.PhotoImage(file='logo.png'))
win.resizable(False, False)

winWidth, winHeight = 900, 600
screenWidth, screenHeight = win.winfo_screenwidth(), win.winfo_screenheight()

x, y = int(screenWidth / 2 - winWidth / 2), int(screenHeight / 2 - winHeight / 2)

win.geometry('{}x{}+{}+{}'.format(winWidth, winHeight, x, y))

fileNameVar: Tk.Variable = Tk.StringVar(value='No Text File Select')
pickBtnVar: Tk.Variable = Tk.StringVar(value='Pick a Text File')

pickBtn: Tk.Button = Tk.Button(win, textvariable=pickBtnVar, font=TkFont.Font(size=10), command=selectTxtFile)
convertBtn: Tk.Button = Tk.Button(win, text='Convert', font=TkFont.Font(size=10), command=convert)

fileNameLabel: Tk.Label = Tk.Label(win, textvariable=fileNameVar, font=TkFont.Font(size=10, slant='italic'))
titleLabel: Tk.Label = Tk.Label(win, text='Text to Excel Convertor', font=TkFont.Font(size=40))
infoLabel: Tk.Label = Tk.Label(win,
                               text='This tool allows you to convert text file into excel. It is created to convert MCQ in Text file to be converted into Excel File Format\n\nText File Format:\nQuestion\nOption A~;~ Option B~;~ Option C~;~ Option D~;~ Answer~;~ Difficulty~;~\n\nNote: This tool ignores duplicate question and empty lines',
                               font=TkFont.Font(size=15), justify='left', wraplength=winWidth)
madeByLabel: Tk.Label = Tk.Label(win, text='Created By Ammar Rangwala', font=TkFont.Font(size=10, weight='bold'))

titleLabel.pack(padx=5, pady=5)
infoLabel.pack(expand=1, padx=5, pady=5)
pickBtn.pack(padx=5, pady=5)
fileNameLabel.pack(padx=5, pady=5)
convertBtn.pack(padx=5, pady=5)
madeByLabel.pack(anchor='s', expand=1, padx=5, pady=5)

win.mainloop()
