from tkinter import *
from pathlib import Path
from tkinter.filedialog import askopenfilename, asksaveasfile, askdirectory

root = Tk()
root.geometry('250x100')
root.title('Reader')

def save_file():
    file = asksaveasfile(filetypes=(("Text Document", "*.txt"),
                                       ('All Files', '*.*')),
                         title='Save File',
                         initialdir=str(Path.home())
                         )
    if file:
        print(file)
    else:
        print('Cancelled')

btn = Button(root, text='save file', command=save_file())
btn.pack(side=TOP, pady=10)

mainloop()