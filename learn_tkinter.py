from tkinter import *

def finish():
    root.destroy()
    print('ручное закрыте окна')

root=Tk()
root.title('Приложение')
root.geometry('400x400+200+200')
root.resizable(False, False)
#root.iconbitmap(default='brend.ico')
icon=PhotoImage(file='Brend.png')
root.iconphoto(False,icon)

label=Label(text='Привет')
label.pack()
root.protocol('WM_DELETE_WINDOW', finish)

root.mainloop()