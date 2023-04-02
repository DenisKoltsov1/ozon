from tkinter import *
from tkinter import ttk
#import re
#import os
#import undetected_chromedriver 
from flashscore import parser,get_str_number,get_count,OpenPixx
import pyperclip
'''

    215973958
   154369096
   646065674



829002092
724581115
825952201
154369096
215973958


875752810
     '''
def bufer(event):
    q = pyperclip.paste()
    e1.insert(0, q)
    return e1
    
    return 1 
def func_parce(event):
    #while e1.get()==None:
        #print('введите id')
    id=e1.get()
    parser(id)
    file ="C:\\ozon1\\Парфюмерия.xlsx"
    r=get_str_number(file)
    count=get_count(r)
    OpenPixx(file,id,r)
    #print('Hello')
    return 1
#окно класс Tk
root = Tk() 
root.title("Парсер")
#размер окна
root.geometry("400x450")

#lbl = Label(root, text="Введите id")
label = ttk.Label(root, text="Введите id (9 цифр)")
label.pack(fill=X, padx=[135, 10], pady=20)
e1 = Entry(width=50)
e1.pack(fill=X, padx=[60, 60], pady=30)

btn1 = ttk.Button(root,  text="вставить скопированный текст",width=40)


btn = ttk.Button(root,  text="начать парсинг",width=20)

'''
label1 = ttk.Label(root, text="вывод данных")
label1.pack(fill=X, padx=[155, 10], pady=10)
text = Text(width=25, height=5,
            fg='white', wrap=WORD)
 
text.pack()
'''
#label.pack(anchor=NW, padx=6, pady=6)
btn1.pack()
btn.pack(expand = 1)

btn1.bind("<Button-1>", bufer)
btn.bind("<Button-1>", func_parce)

#show_message()
root.mainloop()

