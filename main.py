import os
from curses import BUTTON1_CLICKED, newwin
from email.errors import StartBoundaryNotFoundDefect
import tkinter as tk
from tkinter import Frame, ttk, END
from venv import create
import pandas as pd
import openpyxl
from random import randint


class AutocompleteEntry(tk.Entry):
    def set_completion_list(self, completion_list):
        self._completion_list = completion_list
        self._hits = []
        self._hit_index = 0
        self.position = 0
        self.bind('<KeyRelease>', self.handle_keyrelease)

    def autocomplete(self):
        self.position = len(self.get())
        _hits = [element for element in self._completion_list if element.lower().startswith(self.get().lower())]
        if _hits != self._hits:
            self._hit_index = 0
            self._hits = _hits
        if self._hits:
            self.delete(0, tk.END)
            self.insert(0, self._hits[self._hit_index])
            self.select_range(self.position, tk.END)

    def handle_keyrelease(self, event):
        if len(event.keysym) == 1:
            self.autocomplete()

    def clear_text(self):
        self.entry.delete(0, 'end')


def click_btn1():
    global entry_w
    global entry_d
    global button3
    inputWindow = tk.Toplevel(root)
    inputWindow.geometry("800x300")
    #tree.destroy()
    #hscrollbar.destroy()
    words_ls = list(df["words"])
    entry_w = AutocompleteEntry(
        inputWindow,
        fg='black', bg='white',
        insertbackground='white',
        font=(font1, 15))
    entry_w.set_completion_list(words_ls)
    entry_w.place(x=10, y=80)

    entry_d = tk.Entry(inputWindow, width=60, font=(font1, 18), fg="black", bg="white")
    entry_d.place(x=10, y=130)

    def click_btn_view():
        entry_w.destory()
        entry_d.destory()
        button3.destroy()
        show_File()

    def click_btn_save():
        df2 = pd.read_excel(PATH)
        word = entry_w.get()
        defn = entry_d.get()
        entry_w.delete(0, END)
        entry_d.delete(0, END)
        df2 = df2.append({"words": word, "def": defn}, ignore_index=True)
        df2 = df2.sort_values("words", ascending=True)
        df2.to_excel(PATH, index=False)

    def click_btn_show():
        df2 = pd.read_excel(PATH)
        word = entry_w.get()
        if word in list(df2["words"]):
            entry_d.insert(END, df2.iloc[list(df2["words"]).index(word), 1])

    def click_btn_edit():
        df2 = pd.read_excel(PATH)
        word = entry_w.get()
        defn = entry_d.get()
        entry_w.delete(0, END)
        entry_d.delete(0, END)
        if word in list(df2["words"]):
            df2.iloc[list(df["words"]).index(word), 1] = defn
            df2.to_excel(PATH, index=False)
    
    def click_btn_clear():
        entry_w.delete(0, END)
        entry_d.delete(0, END)

    button3 = tk.Button(inputWindow, text="save", font=(font1, 15), command=click_btn_save)
    button3.place(x=10, y=250)
    button4 = tk.Button(inputWindow, text="show", font=(font1, 15), command=click_btn_show)
    button4.place(x=100, y=250)
    button5 = tk.Button(inputWindow, text="edit", font=(font1, 15), command=click_btn_edit)
    button5.place(x=190, y=250)
    button6 = tk.Button(inputWindow, text="clear", font=(font1, 15), command=click_btn_clear)
    button6.place(x=270, y=250)

def click_btn2():
    tree.destroy()
    hscrollbar.destroy()

    label = tk.Label(root, text="TEST", font=(font1, 12))
    label.place(x=50, y=50)

    word_t = tk.Entry(frame, width=60,font=(font1, 18), fg="black", bg="white")
    word_t.place(x=10, y=180)
    def_t = tk.Entry(frame, width=60,font=(font1, 18), fg="black", bg="white")
    def_t.place(x=10, y=230)

    def click_btn_show_word():
        def click_btn_show_answer():
            def_t.insert(END, defn1)

            answer_btn.destroy()

            next_btn = tk.Button(root, text="next", font=(font1, 15), command=click_btn_show_word)
            next_btn.place(x=10, y=300)

        word_t.delete(0, END)
        def_t.delete(0, END)

        rand = randint(0, len(df["words"]))
        sr1 = list(df.iloc[rand])
        word1 = sr1[0]
        defn1 = sr1[1]

        word_t.insert(END, word1)

        start_btn.destroy()

        answer_btn = tk.Button(root, text="answer", font=(font1, 15), command=click_btn_show_answer)
        answer_btn.place(x=10, y=300)

    start_btn = tk.Button(root, text="start", font=(font1, 15), command=click_btn_show_word)

    start_btn.place(x=10, y=300)


def click_btn3():
    for row in tree.get_children():
        tree.delete(row)
    df = pd.read_excel(PATH)
    df_rows = df.to_numpy().tolist()
    for row in df_rows:
        tree.insert("", "end", values=row)
    
    words_ls = list(df["words"])
    entry_w.set_completion_list(words_ls)

def click_btn4():
    df2 = pd.read_excel(PATH)
    df2.to_excel("data_copy.xlsx", index=False)
    



def show_File():
    global hscrollbar
    global frame
    global tree
    # clear_treeview()

    frame = Frame(root)
    frame.pack()
    tree = ttk.Treeview(frame, height=23)

    tree["column"] = (0, 1)
    tree["show"] = "headings"

    tree.column(0, width=170)
    tree.column(1, width=1000)
    tree.heading(0, text="words")
    tree.heading(1, text="def")

    word_w = root.winfo_width()/8
    def_w = word_w*7

   # Put Data in Rows
    df_rows = df.to_numpy().tolist()
    for row in df_rows:
        tree.insert("", "end", values=row)

    tree.pack()

    hscrollbar = ttk.Scrollbar(frame, orient=tk.HORIZONTAL, command=tree.xview)
    tree.configure(xscrollcommand=lambda f, l: hscrollbar.set(f, l))
    hscrollbar.pack(fill='x')

    button1 = tk.Button(root, text="input", font=(
        font1, 15), command=click_btn1)
    button1.place(x=600, y=470)

    button2 = tk.Button(root, text="test", font=(
        font1, 15), command=click_btn2)
    button2.place(x=670, y=470)

    button3 = tk.Button(root, text="reload", font=(
        font1, 15), command=click_btn3)
    button3.place(x=520, y=470)

    button4 = tk.Button(root, text="backup", font=(
        font1, 15), command=click_btn4)
    button4.place(x=10, y=470)
    


root = tk.Tk()
root.title("en-jp")
root.configure(bg="black")
root.resizable(False, False)
root.geometry("800x500")

style = ttk.Style()
style.theme_use('aqua')


PATH = "en-jp.xlsx"

if not os.path.exists(PATH):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(['words', 'def'])
    wb.save(PATH)

df = pd.read_excel(PATH)
font1 = "arial"


show_File()

root.mainloop()
