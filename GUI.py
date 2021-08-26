import tkinter as tk
import MainControl
from tkinter import *
import doc2docx
import os

#global var that stores error message from making contract obj
#into strings
errorMSGstr = ''

def EnterData():
    global errorMSGstr
    try:
        MainControl.EnterData(currentPath)

        for contract in MainControl.errorList:
            errorMSGstr += '\n' + contract.errorMsg
        if len(MainControl.error) == 0:
            if len(MainControl.errorList) == 0 and len(MainControl.correctList) != 0 :
                runState["text"] = "All Contracts are Sucessfully Summarized and Outputted ！"
                runState['fg'] = 'green'
                MainControl.correctList = []
                MainControl.errorList = []
                errorMSGstr = ''
            elif len(MainControl.errorList) != 0 and len(MainControl.correctList) != 0:
                runState["text"] = "Some Contracts are Sucessfully Summarized and Outputted ！"
                runState['fg'] = 'orange'
                errorMSG['text'] = errorMSGstr
                MainControl.correctList = []
                MainControl.errorList = []
                errorMSGstr = ''
            elif len(MainControl.errorList) == 0 and len(MainControl.correctList) == 0:
                runState["text"] = "‘前台拷贝’ 文件夹内无合同！"
                runState['fg'] = 'orange'
                errorMSG['text'] = errorMSGstr
                MainControl.correctList = []
                MainControl.errorList = []
                errorMSGstr = ''
            else:
                runState["text"] = "All Contracts Failed to Summarize and Output ！！"
                runState['fg'] = 'red'
                errorMSG['text'] = errorMSGstr
                MainControl.correctList = []
                MainControl.errorList = []
                errorMSGstr = ''
        else:
            runState["text"] = MainControl.error
            runState['fg'] = 'red'

    except FileNotFoundError:
        print('file not found')
        runState["text"] = "File not Found, recheck path"
        runState['fg'] = 'red'
    # except:
    #     runState["text"] = "Fail to extract contract, contact admin"
    #     runState['fg'] = 'red'

#takes in a string and erase a existing txt file to put the string in and save
def editTxt(data):

    f = open('pathSetting.txt', 'r+',encoding="utf8")
    f.truncate(0)
    f.close()
    f = open('pathSetting.txt', 'r+',encoding="utf8")
    f.write(data)
    f.close()
#function that updates the path
def updatePath():
    global currentPath
    currentPath = pathEntry.get()
    os.chdir(defaultPath)
    editTxt(pathEntry.get())
    runState["text"] = "Path Changed"
    runState["fg"] = 'green'
    return pathEntry.get()
#function that converts doc files to docx
def convertFile():
    try:
        errorMSGConv['text'] = 'Converting File'
        errorMSGConv['fg'] = 'orange'
        window.update()
        doc2docx.convert2docx(currentPath)
        if len(doc2docx.convertError) == 0:
            errorMSGConv['text'] = 'All Files Converted Successfully'
            errorMSGConv['fg'] = 'green'
        else:
            errorMSGConv['text'] = doc2docx.convertError
            errorMSGConv['fg'] = 'red'
    except:
        errorMSGConv['text'] = 'Error in File Conversion'
        errorMSGConv['fg'] = 'red'
        print(window.report_callback_exception())



#creating window for tinker
#setting the icon and title
window = tk.Tk()
window.title("Contract Management System")
# photo = PhotoImage(file = "2logo.png")
# window.iconphoto(False, photo)

#get the stored path and the path of the location of fiel
f = open('pathSetting.txt', 'r',encoding="utf8")
currentPath = f.read()
f.close()
defaultPath = os.getcwd()

#title and description of the program
title = tk.Label(text="Contract Management System",borderwidth=1, relief="solid")
title.config(font=("microsoft yahei", 18))
title.pack()
description = tk.Label(text="Please put all contracts in 'InputFile' folder",borderwidth=5)
description.config(font=("microsoft yahei", 12))
description.pack()

#convert button
buttonConv = tk.Button(
    text="Change Format",
    width=17,
    height=1,
    fg = 'green',
    command = convertFile
)
buttonConv.config(font=("microsoft yahei", 20))
buttonConv.pack()

#error message for the conversion process
errorMSGConv = tk.Label(text = '', fg  ='red' )
errorMSGConv.config(font=("microsoft yahei", 12))
errorMSGConv.pack()


#entry button
buttonLuRu = tk.Button(
    text="Summarize",
    width=10,
    height=1,
    fg = 'green',
    command = EnterData
)
buttonLuRu.config(font=("microsoft yahei", 20))
buttonLuRu.pack()

#white space and layout
whitespace = tk.Label(text = '')
whitespace.pack()


#entry bar and description for changing path
dsrpText = tk.Label(text = "Path to Folder：")
dsrpText.config(font=("microsoft yahei", 12))
dsrpText.pack()
pathEntry = tk.Entry(width = 50)
pathEntry.insert(0,currentPath)
pathEntry.pack()
whitespace2 = tk.Label(text = '')
whitespace2.pack()

#button for updating the path
buttonGengxin = tk.Button(
    text="Change Path",
    width=13,
    height=1,
    fg = 'green',
    command = updatePath
)
buttonGengxin.config(font=("microsoft yahei", 20))
buttonGengxin.pack()

#label that shows the current state of the the entering process
runState = tk.Label(text = 'Running ！', fg  = "green")
runState.config(font=("microsoft yahei", 12))
runState.pack()

#error message for the conversion process
errorMSG = tk.Label(text = errorMSGstr, fg  ='red' )
errorMSG.config(font=("microsoft yahei", 12))
errorMSG.pack(side = 'bottom', anchor="c",expand = True, fill = 'both')

window.mainloop()