from tkinter import *

clicks=0

def finish():
    root.destroy()
    print('ручное закрыте окна')

def click_button():
    global clicks
    clicks+=1
    btn['text']=f'Нажали{clicks}раз'

def entered(event):
    btn['text']='Навел'

def lefted(event):
    btn['text']='Убрал'

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

btn=Button(text='Жми', command=click_button)
#btn.pack()
btn.place(x=20,y=20)
btn.bind('<Enter>', entered)
btn.bind('<Leave>', lefted)

root.mainloop()