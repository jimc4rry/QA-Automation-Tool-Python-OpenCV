import tkinter as tk
from tkinter import ttk
import pyautogui
import time
from tkinter import *
from tkinter.ttk import *
import xlwt
from xlwt import *
from PIL import ImageTk, Image
from tkinter import messagebox

class QA_Automation_Tool:

    def __init__(self, master, **kwargs):
        width = kwargs.pop('width', None)
        height = kwargs.pop('height', None)
        self.outer = tk.Frame(master, **kwargs)
        self.vsb = ttk.Scrollbar(self.outer, orient=tk.VERTICAL)
        self.vsb.grid(row=0, column=1, sticky='ns')
        self.hsb = ttk.Scrollbar(self.outer, orient=tk.HORIZONTAL)
        self.hsb.grid(row=1, column=0, sticky='ew')
        self.canvas = tk.Canvas(self.outer, highlightthickness=0, width=width, height=height)
        self.canvas.grid(row=0, column=0, sticky='nsew')
        self.outer.rowconfigure(0, weight=1)
        self.outer.columnconfigure(0, weight=1)
        self.canvas['yscrollcommand'] = self.vsb.set
        self.canvas['xscrollcommand'] = self.hsb.set
        # mouse scroll does not seem to work with just "bind"; You have
        # to use "bind_all". Therefore, to use multiple windows you have
        # to bind_all in the current widget
        self.canvas.bind("<Enter>", self._bind_mouse)
        self.canvas.bind("<Leave>", self._unbind_mouse)
        self.vsb['command'] = self.canvas.yview
        self.hsb['command'] = self.canvas.xview
        self.inner = tk.Frame(self.canvas)
        # pack the inner Frame into the Canvas with the top left corner 4 pixels offset
        self.canvas.create_window(4, 4, window=self.inner, anchor='nw')
        self.inner.bind("<Configure>", self._on_frame_configure)
        self.outer_attr = set(dir(tk.Widget))

    def __getattr__(self, item):
        if item in self.outer_attr:
            # geometry attributes etc. (eg pack, destroy, tkraise) are passed on to self.outer
            return getattr(self.outer, item)
        else:
            # all other attributes (_w, children, etc.) are passed to self.inner
            return getattr(self.inner, item)

    def _on_frame_configure(self, event=None):
        x1, y1, x2, y2 = self.canvas.bbox("all")
        height = self.canvas.winfo_height()
        width = self.canvas.winfo_width()
        self.canvas.config(scrollregion=(0, 0, max(x2, width), max(y2, height)))

    def _bind_mouse(self, event=None):
        self.canvas.bind_all("<4>", self._on_mousewheel)
        self.canvas.bind_all("<5>", self._on_mousewheel)
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

    def _unbind_mouse(self, event=None):
        self.canvas.unbind_all("<4>")
        self.canvas.unbind_all("<5>")
        self.canvas.unbind_all("<MouseWheel>")

    def _on_mousewheel(self, event):
        """Linux uses event.num; Windows / Mac uses event.delta"""
        func = self.canvas.xview_scroll if event.state & 1 else self.canvas.yview_scroll
        if event.num == 4 or event.delta > 0:
            func(-1, "units")
        elif event.num == 5 or event.delta < 0:
            func(1, "units")

    def __str__(self):
        return str(self.outer)

root = tk.Tk()
root.title("QA Automation Tool")
root.geometry('610x700')
p1 = PhotoImage(file='images/QA Automation tool.png')
root.iconphoto(False, p1)
lbl = tk.Label(root, text="Hold shift while using the scroll wheel to scroll horizontally")
lbl.pack()
f = ("Times bold", 14)  # change the font
main_frame = QA_Automation_Tool(root, width=900, borderwidth=2, relief=tk.SUNKEN, background="light gray")
main_frame.pack(fill=tk.BOTH, expand=True)

Top_frame = Frame(main_frame)
Top_frame.place(width=500)
Top_frame.pack(ipadx=20, ipady=10, expand=True)
# Create an object of tkinter ImageTk
img = ImageTk.PhotoImage(Image.open("images/QA Automation tool.PNG"))
# Create a Label Widget to display the text or Image
label = tk.Label(Top_frame, image=img)
label.config(bg="red")
label.pack()

# Frame for Group Senario 1
frame0 = tk.LabelFrame(main_frame, text="Group Senario 1", relief=RAISED)
frame0.place(width=500)
frame0.config(bg='light blue')
frame0.pack(ipadx=200, ipady=10, expand=True)
tk.Label(frame0, wraplength=290, font="helvetica 8", justify="center", bg='light blue',
      text="1)Enter description of senario here").pack()
tk.Label(frame0, wraplength=290, font="helvetica 8", justify="center", bg='light blue',
      text="2)Enter description of senario here").pack()

# Frame for test1
frame1 = tk.LabelFrame(main_frame, text="test1", relief=RAISED)
frame1.place(width=500)
frame1.config(bg='green')
frame1.pack(ipadx=200, ipady=10, expand=True)
tk.Label(frame1, wraplength=290, font="helvetica 8", justify="center", bg="green", fg="yellow",
      text="1)Enter description of senario here").pack()
tk.Label(frame1, wraplength=290, font="helvetica 8", justify="center", bg="green", fg="yellow",
      text="2)Enter description of senario here").pack()
v1 = tk.StringVar()
v1.set("No value")  # initializing the choice, i.e. Python

# Frame for test2
frame2 = tk.LabelFrame(main_frame, text="test2", relief=RAISED)
frame2.place(width=500)
frame2.config(bg='green')
frame2.pack(ipadx=200, ipady=10, expand=True)
tk.Label(frame2, wraplength=290, font="helvetica 8", justify="center", bg="green", fg="yellow",
      text="1)Enter description of senario here").pack()
tk.Label(frame2, wraplength=290, font="helvetica 8", justify="center", bg="green", fg="yellow",
      text="2)Enter description of senario here").pack()
v2 = tk.StringVar()
v2.set("No value")  # initializing the choice, i.e. Python

# Frame for test3
frame3 = tk.LabelFrame(main_frame, text="test3", relief=RAISED)
frame3.place(width=500)
frame3.config(bg='green')
frame3.pack(ipadx=200, ipady=10, expand=True)
tk.Label(frame3, wraplength=290, font="helvetica 8", justify="center", bg="green", fg="yellow",
      text="1)Enter description of senario here").pack()
tk.Label(frame3, wraplength=290, font="helvetica 8", justify="center", bg="green", fg="yellow",
      text="2)Enter description of senario here").pack()
v3 = tk.StringVar()
v3.set("No value")  # initializing the choice, i.e. Python

# Frame for test4
frame4 = tk.LabelFrame(main_frame, text="test4", relief=RAISED)
frame4.place(width=500)
frame4.config(bg='green')
frame4.pack(ipadx=200, ipady=10, expand=True)
tk.Label(frame4, wraplength=290, font="helvetica 8", justify="center", bg="green", fg="yellow",
      text="1)Enter description of senario here").pack()
tk.Label(frame4, wraplength=290, font="helvetica 8", justify="center", bg="green", fg="yellow",
      text="2)Enter description of senario here").pack()
v4 = tk.StringVar()
v4.set("No value")  # initializing the choice, i.e. Python

# Frame for test5
frame5 = tk.LabelFrame(main_frame, text="test5", relief=RAISED)
frame5.place(width=500)
frame5.config(bg='green')
frame5.pack(ipadx=200, ipady=10, expand=True)
tk.Label(frame5, wraplength=290, font="helvetica 8", justify="center", bg="green", fg="yellow",
      text="1)Enter description of senario here").pack()
tk.Label(frame5, wraplength=290, font="helvetica 8", justify="center", bg="green", fg="yellow",
      text="2)Enter description of senario here").pack()
v5 = tk.StringVar()
v5.set("No value")  # initializing the choice, i.e. Python

# Frame for test6
frame6 = tk.LabelFrame(main_frame, text="test6", relief=RAISED)
frame6.place(width=500)
frame6.config(bg='green')
frame6.pack(ipadx=200, ipady=10, expand=True)
tk.Label(frame6, wraplength=290, font="helvetica 8", justify="center", bg="green", fg="yellow",
      text="1)Enter description of senario here").pack()
tk.Label(frame6, wraplength=290, font="helvetica 8", justify="center", bg="green", fg="yellow",
      text="2)Enter description of senario here").pack()
v6 = tk.StringVar()
v6.set("No value")  # initializing the choice, i.e. Python

# Frame for test7
frame7 = tk.LabelFrame(main_frame, text="test7", relief=RAISED)
frame7.place(width=500)
frame7.config(bg='green')
frame7.pack(ipadx=200, ipady=10, expand=True)
tk.Label(frame7, wraplength=290, font="helvetica 8", justify="center", bg="green", fg="yellow",
      text="1)Enter description of senario here").pack()
tk.Label(frame7, wraplength=290, font="helvetica 8", justify="center", bg="green", fg="yellow",
      text="2)Enter description of senario here").pack()
v7 = tk.StringVar()
v7.set("No value")  # initializing the choice, i.e. Python

# Frame for test8
frame8 = tk.LabelFrame(main_frame, text="test8", relief=RAISED)
frame8.place(width=500)
frame8.config(bg='green')
frame8.pack(ipadx=200, ipady=10, expand=True)
tk.Label(frame8, wraplength=290, font="helvetica 8", justify="center", bg="green", fg="yellow",
      text="1)Enter description of senario here").pack()
tk.Label(frame8, wraplength=290, font="helvetica 8", justify="center", bg="green", fg="yellow",
      text="2)Enter description of senario here").pack()
v8 = tk.StringVar()
v8.set("No value")  # initializing the choice, i.e. Python

# Frame for test9
frame9 = tk.LabelFrame(main_frame, text="test9", relief=RAISED)
frame9.place(width=500)
frame9.config(bg='green')
frame9.pack(ipadx=200, ipady=10, expand=True)
tk.Label(frame9, wraplength=290, font="helvetica 8", justify="center", bg="green", fg="yellow",
      text="1)Enter description of senario here").pack()
tk.Label(frame9, wraplength=290, font="helvetica 8", justify="center", bg="green", fg="yellow",
      text="2)Enter description of senario here").pack()
v9 = tk.StringVar()
v9.set("No value")  # initializing the choice, i.e. Python

# Frame for test10
frame10 = tk.LabelFrame(main_frame, text="test10", relief=RAISED)
frame10.place(width=500)
frame10.config(bg='green')
frame10.pack(ipadx=200, ipady=10, expand=True)
tk.Label(frame10, wraplength=290, font="helvetica 8", justify="center", bg="green", fg="yellow",
      text="1)Enter description of senario here").pack()
tk.Label(frame10, wraplength=290, font="helvetica 8", justify="center", bg="green", fg="yellow",
      text="2)Enter description of senario here").pack()
v10 = tk.StringVar()
v10.set("No value")  # initializing the choice, i.e. Python

# Button senario1
button1 = tk.Button(frame1, fg='yellow', bg='blue', text="senario1", command=lambda: senario_1())
button1.place(rely=0, relx=0)

# Button senario2
button2 = tk.Button(frame2, fg='yellow', bg='blue', text="senario2")
button2.place(rely=0, relx=0)

# Button senario3
button3 = tk.Button(frame3, fg='yellow', bg='blue', text="senario3")
button3.place(rely=0, relx=0)

# Button senario4
button4 = tk.Button(frame4, fg='yellow', bg='blue', text="senario4")
button4.place(rely=0, relx=0)

# Button senario5
button5 = tk.Button(frame5, fg='yellow', bg='blue', text="senario5")
button5.place(rely=0, relx=0)

# Button senario6
button6 = tk.Button(frame6, fg='yellow', bg='blue', text="senario6")
button6.place(rely=0, relx=0)

# Button senario7
button7 = tk.Button(frame7, fg='yellow', bg='blue', text="senario7", )
button7.place(rely=0, relx=0)

# Button senario8
button8 = tk.Button(frame8, fg='yellow', bg='blue', text="senario8", )
button8.place(rely=0, relx=0)

# Button senario9
button9 = tk.Button(frame9, fg='yellow', bg='blue', text="senario9", )
button9.place(rely=0, relx=0)

# Button senario10
button10 = tk.Button(frame10, fg='yellow', bg='blue', text="senario10", )
button10.place(rely=0, relx=0)

"""      

      SENARIO2


"""
# Frame for Group of senario 2
frame0_1 = tk.LabelFrame(main_frame, text="Group of senario 2", relief=RAISED)
frame0_1.place(width=500)
frame0_1.config(bg='light blue')
frame0_1.pack(ipadx=200, ipady=10, expand=True)
tk.Label(frame0_1, wraplength=290, font="helvetica 8", justify="center", bg='light blue',
      text="1)Enter description of senario here").pack()
tk.Label(frame0_1, wraplength=290, font="helvetica 8", justify="center", bg='light blue',
      text="2)Enter description of senario here").pack()

# Frame for test1
frame1_1 = tk.LabelFrame(main_frame, text="test1", relief=RAISED)
frame1_1.place(width=500)
frame1_1.config(bg='green')
frame1_1.pack(ipadx=200, ipady=10, expand=True)
tk.Label(frame1_1, wraplength=290, font="helvetica 8", justify="center", bg="green", fg="yellow",
      text="1)Enter description of senario here").pack()
tk.Label(frame1_1, wraplength=290, font="helvetica 8", justify="center", bg="green", fg="yellow",
      text="2)Enter description of senario here").pack()
v11 = tk.StringVar()
v11.set("No value")  # initializing the choice, i.e. Python

# Frame for test2
frame2_1 = tk.LabelFrame(main_frame, text="test2", relief=RAISED)
frame2_1.place(width=500)
frame2_1.config(bg='green')
frame2_1.pack(ipadx=200, ipady=10, expand=True)
tk.Label(frame2_1, wraplength=290, font="helvetica 8", justify="center", bg="green", fg="yellow",
      text="1)Enter description of senario here").pack()
tk.Label(frame2_1, wraplength=290, font="helvetica 8", justify="center", bg="green", fg="yellow",
      text="2)Enter description of senario here").pack()
v12 = tk.StringVar()
v12.set("No value")  # initializing the choice, i.e. Python

# Frame for test3
frame3_1 = tk.LabelFrame(main_frame, text="test3", relief=RAISED)
frame3_1.place(width=500)
frame3_1.config(bg='green')
frame3_1.pack(ipadx=200, ipady=10, expand=True)
tk.Label(frame3_1, wraplength=290, font="helvetica 8", justify="center", bg="green", fg="yellow",
      text="1)Enter description of senario here").pack()
tk.Label(frame3_1, wraplength=290, font="helvetica 8", justify="center", bg="green", fg="yellow",
      text="2)Enter description of senario here").pack()
v13 = tk.StringVar()
v13.set("No value")  # initializing the choice, i.e. Python

# Frame for test4
frame4_1 = tk.LabelFrame(main_frame, text="test4", relief=RAISED)
frame4_1.place(width=500)
frame4_1.config(bg='green')
frame4_1.pack(ipadx=200, ipady=10, expand=True)
tk.Label(frame4_1, wraplength=290, font="helvetica 8", justify="center", bg="green", fg="yellow",
      text="1)Enter description of senario here").pack()
tk.Label(frame4_1, wraplength=290, font="helvetica 8", justify="center", bg="green", fg="yellow",
      text="2)Enter description of senario here").pack()
v14 = tk.StringVar()
v14.set("No value")  # initializing the choice, i.e. Python

# Frame for test5
frame5_1 = tk.LabelFrame(main_frame, text="test5", relief=RAISED)
frame5_1.place(width=500)
frame5_1.config(bg='green')
frame5_1.pack(ipadx=200, ipady=10, expand=True)
tk.Label(frame5_1, wraplength=290, font="helvetica 8", justify="center", bg="green", fg="yellow",
      text="1)Enter description of senario here").pack()
tk.Label(frame5_1, wraplength=290, font="helvetica 8", justify="center", bg="green", fg="yellow",
      text="2)Enter description of senario here").pack()
v15 = tk.StringVar()
v15.set("No value")  # initializing the choice, i.e. Python

# Frame for test6
frame6_1 = tk.LabelFrame(main_frame, text="test6", relief=RAISED)
frame6_1.place(width=500)
frame6_1.config(bg='green')
frame6_1.pack(ipadx=200, ipady=10, expand=True)
tk.Label(frame6_1, wraplength=290, font="helvetica 8", justify="center", bg="green", fg="yellow",
      text="1)Enter description of senario here").pack()
tk.Label(frame6_1, wraplength=290, font="helvetica 8", justify="center", bg="green", fg="yellow",
      text="2)Enter description of senario here").pack()
v16 = tk.StringVar()
v16.set("No value")  # initializing the choice, i.e. Python

# Frame for test7
frame7_1 = tk.LabelFrame(main_frame, text="test7", relief=RAISED)
frame7_1.place(width=500)
frame7_1.config(bg='green')
frame7_1.pack(ipadx=200, ipady=10, expand=True)
tk.Label(frame7_1, wraplength=290, font="helvetica 8", justify="center", bg="green", fg="yellow",
      text="1)Enter description of senario here").pack()
tk.Label(frame7_1, wraplength=290, font="helvetica 8", justify="center", bg="green", fg="yellow",
      text="2)Enter description of senario here").pack()
v17 = tk.StringVar()
v17.set("No value")  # initializing the choice, i.e. Python

# Frame for test8
frame8_1 = tk.LabelFrame(main_frame, text="test8", relief=RAISED)
frame8_1.place(width=500)
frame8_1.config(bg='green')
frame8_1.pack(ipadx=200, ipady=10, expand=True)
tk.Label(frame8_1, wraplength=290, font="helvetica 8", justify="center", bg="green", fg="yellow",
      text="1)Enter description of senario here").pack()
tk.Label(frame8_1, wraplength=290, font="helvetica 8", justify="center", bg="green", fg="yellow",
      text="2)Enter description of senario here").pack()
v18 = tk.StringVar()
v18.set("No value")  # initializing the choice, i.e. Python

# Frame for test9
frame9_1 = tk.LabelFrame(main_frame, text="test9", relief=RAISED)
frame9_1.place(width=500)
frame9_1.config(bg='green')
frame9_1.pack(ipadx=200, ipady=10, expand=True)
tk.Label(frame9_1, wraplength=290, font="helvetica 8", justify="center", bg="green", fg="yellow",
         text="1)Enter description of senario here").pack()
tk.Label(frame9_1, wraplength=290, font="helvetica 8", justify="center", bg="green", fg="yellow",
         text="2)Enter description of senario here").pack()
v19 = tk.StringVar()
v19.set("No value")  # initializing the choice, i.e. Python

# Frame for test10
frame10_1 = tk.LabelFrame(main_frame, text="test10", relief=RAISED)
frame10_1.place(width=500)
frame10_1.config(bg='green')
frame10_1.pack(ipadx=200, ipady=10, expand=True)
tk.Label(frame10_1, wraplength=290, font="helvetica 8", justify="center", bg="green", fg="yellow",
         text="1)Enter description of senario here").pack()
tk.Label(frame10_1, wraplength=290, font="helvetica 8", justify="center", bg="green", fg="yellow",
         text="2)Enter description of senario here").pack()
v20 = tk.StringVar()
v20.set("No value")  # initializing the choice, i.e. Python

# Button senario1
button1_1 = tk.Button(frame1_1, fg='yellow', bg='blue', text="senario1", command=lambda: senario_1())
button1_1.place(rely=0, relx=0)

# Button senario2
button2_1 = tk.Button(frame2_1, fg='yellow', bg='blue', text="senario2")
button2_1.place(rely=0, relx=0)

# Button senario3
button3_1 = tk.Button(frame3_1, fg='yellow', bg='blue', text="senario3")
button3_1.place(rely=0, relx=0)

# Button senario4
button4_1 = tk.Button(frame4_1, fg='yellow', bg='blue', text="senario4", )
button4_1.place(rely=0, relx=0)

# Button senario5
button5_1 = tk.Button(frame5_1, fg='yellow', bg='blue', text="senario5", )
button5_1.place(rely=0, relx=0)

# Button senario6
button6_1 = tk.Button(frame6_1, fg='yellow', bg='blue', text="senario6", )
button6_1.place(rely=0, relx=0)

# Button senario7
button7_1 = tk.Button(frame7_1, fg='yellow', bg='blue', text="senario7", )
button7_1.place(rely=0, relx=0)

# Button senario8
button8_1 = tk.Button(frame8_1, fg='yellow', bg='blue', text="senario8", )
button8_1.place(rely=0, relx=0)

# Button senario9
button9_1 = tk.Button(frame9_1, fg='yellow', bg='blue', text="senario9", )
button9_1.place(rely=0, relx=0)

# Button senario10
button10_1 = tk.Button(frame10_1, fg='yellow', bg='blue', text="senario10", )
button10_1.place(rely=0, relx=0)


def selection_R1():
    selection = str(v1.get())
    label_R1.config(text=selection)
def selection_R2():
    selection = str(v2.get())
    label_R2.config(text=selection)
def selection_R3():
    selection = str(v3.get())
    label_R3.config(text=selection)
def selection_R4():
    selection = str(v4.get())
    label_R4.config(text=selection)
def selection_R5():
    selection = str(v5.get())
    label_R5.config(text=selection)
def selection_R6():
    selection = str(v6.get())
    label_R6.config(text=selection)
def selection_R7():
    selection = str(v7.get())
    label_R7.config(text=selection)
def selection_R8():
    selection = str(v8.get())
    label_R8.config(text=selection)
def selection_R9():
    selection = str(v9.get())
    label_R9.config(text=selection)
def selection_R10():
    selection = str(v10.get())
    label_R10.config(text=selection)
def selection_R1_1():
    selection = str(v11.get())
    label_R1_1.config(text=selection)
def selection_R2_1():
    selection = str(v12.get())
    label_R2_1.config(text=selection)
def selection_R3_1():
    selection = str(v13.get())
    label_R3_1.config(text=selection)
def selection_R4_1():
    selection = str(v14.get())
    label_R4_1.config(text=selection)
def selection_R5_1():
    selection = str(v15.get())
    label_R5_1.config(text=selection)
def selection_R6_1():
    selection = str(v16.get())
    label_R6_1.config(text=selection)
def selection_R7_1():
    selection = str(v17.get())
    label_R7_1.config(text=selection)
def selection_R8_1():
    selection = str(v18.get())
    label_R8_1.config(text=selection)
def selection_R9_1():
    selection = str(v19.get())
    label_R9_1.config(text=selection)
def selection_R10_1():
    selection = str(v20.get())
    label_R10_1.config(text=selection)

choises = [("Passed", "Passed"), ("Not Completed", "Not Completed"), ("N/A", "N/A")]

for choises, val in choises:
    tk.Radiobutton(frame1, text=choises, padx=20, variable=v1, bg='green', value=val, command=selection_R1).pack()
    tk.Radiobutton(frame2, text=choises, padx=20, variable=v2, bg='green', value=val, command=selection_R2).pack()
    tk.Radiobutton(frame3, text=choises, padx=20, variable=v3, bg='green', value=val, command=selection_R3).pack()
    tk.Radiobutton(frame4, text=choises, padx=20, variable=v4, bg='green', value=val, command=selection_R4).pack()
    tk.Radiobutton(frame5, text=choises, padx=20, variable=v5, bg='green', value=val, command=selection_R5).pack()
    tk.Radiobutton(frame6, text=choises, padx=20, variable=v6, bg='green', value=val, command=selection_R6).pack()
    tk.Radiobutton(frame7, text=choises, padx=20, variable=v7, bg='green', value=val, command=selection_R7).pack()
    tk.Radiobutton(frame8, text=choises, padx=20, variable=v8, bg='green', value=val, command=selection_R8).pack()
    tk.Radiobutton(frame9, text=choises, padx=20, variable=v9, bg='green', value=val, command=selection_R9).pack()
    tk.Radiobutton(frame10, text=choises, padx=20, variable=v10, bg='green', value=val, command=selection_R10).pack()
    tk.Radiobutton(frame1_1, text=choises, padx=20, variable=v11, bg='green', value=val, command=selection_R1_1).pack()
    tk.Radiobutton(frame2_1, text=choises, padx=20, variable=v12, bg='green', value=val, command=selection_R2_1).pack()
    tk.Radiobutton(frame3_1, text=choises, padx=20, variable=v13, bg='green', value=val, command=selection_R3_1).pack()
    tk.Radiobutton(frame4_1, text=choises, padx=20, variable=v14, bg='green', value=val, command=selection_R4_1).pack()
    tk.Radiobutton(frame5_1, text=choises, padx=20, variable=v15, bg='green', value=val, command=selection_R5_1).pack()
    tk.Radiobutton(frame6_1, text=choises, padx=20, variable=v16, bg='green', value=val, command=selection_R6_1).pack()
    tk.Radiobutton(frame7_1, text=choises, padx=20, variable=v17, bg='green', value=val, command=selection_R7_1).pack()
    tk.Radiobutton(frame8_1, text=choises, padx=20, variable=v18, bg='green', value=val, command=selection_R8_1).pack()
    tk.Radiobutton(frame9_1, text=choises, padx=20, variable=v19, bg='green', value=val, command=selection_R9_1).pack()
    tk.Radiobutton(frame10_1, text=choises, padx=20, variable=v20, bg='green', value=val,
                   command=selection_R10_1).pack()

label_R1 = tk.Label(frame1, width=15)
label_R1.pack()
label_R2 = tk.Label(frame2, width=15)
label_R2.pack()
label_R3 = tk.Label(frame3, width=15)
label_R3.pack()
label_R4 = tk.Label(frame4, width=15)
label_R4.pack()
label_R5 = tk.Label(frame5, width=15)
label_R5.pack()
label_R6 = tk.Label(frame6, width=15)
label_R6.pack()
label_R7 = tk.Label(frame7, width=15)
label_R7.pack()
label_R8 = tk.Label(frame8, width=15)
label_R8.pack()
label_R9 = tk.Label(frame9, width=15)
label_R9.pack()
label_R10 = tk.Label(frame10, width=15)
label_R10.pack()
label_R1_1 = tk.Label(frame1_1, width=15)
label_R1_1.pack()
label_R2_1 = tk.Label(frame2_1, width=15)
label_R2_1.pack()
label_R3_1 = tk.Label(frame3_1, width=15)
label_R3_1.pack()
label_R4_1 = tk.Label(frame4_1, width=15)
label_R4_1.pack()
label_R5_1 = tk.Label(frame5_1, width=15)
label_R5_1.pack()
label_R6_1 = tk.Label(frame6_1, width=15)
label_R6_1.pack()
label_R7_1 = tk.Label(frame7_1, width=15)
label_R7_1.pack()
label_R8_1 = tk.Label(frame8_1, width=15)
label_R8_1.pack()
label_R9_1 = tk.Label(frame9_1, width=15)
label_R9_1.pack()
label_R10_1 = tk.Label(frame10_1, width=15)
label_R10_1.pack()

def show_widget():
    frame0.pack(ipadx=200, ipady=10, expand=True)
    frame1.pack(ipadx=200, ipady=10, expand=True)
    frame2.pack(ipadx=200, ipady=10, expand=True)
    frame3.pack(ipadx=200, ipady=10, expand=True)
    frame4.pack(ipadx=200, ipady=10, expand=True)
    frame5.pack(ipadx=200, ipady=10, expand=True)
    frame6.pack(ipadx=200, ipady=10, expand=True)
    frame7.pack(ipadx=200, ipady=10, expand=True)
    frame8.pack(ipadx=200, ipady=10, expand=True)
    frame9.pack(ipadx=200, ipady=10, expand=True)
    frame10.pack(ipadx=200, ipady=10, expand=True)
    frame0_1.pack(ipadx=200, ipady=10, expand=True)
    frame1_1.pack(ipadx=200, ipady=10, expand=True)
    frame2_1.pack(ipadx=200, ipady=10, expand=True)
    frame3_1.pack(ipadx=200, ipady=10, expand=True)
    frame4_1.pack(ipadx=200, ipady=10, expand=True)
    frame5_1.pack(ipadx=200, ipady=10, expand=True)
    frame6_1.pack(ipadx=200, ipady=10, expand=True)
    frame7_1.pack(ipadx=200, ipady=10, expand=True)
    frame8_1.pack(ipadx=200, ipady=10, expand=True)
    frame9_1.pack(ipadx=200, ipady=10, expand=True)
    frame10_1.pack(ipadx=200, ipady=10, expand=True)

def hide_widget():
    frame0.pack_forget()
    frame1.pack_forget()
    frame2.pack_forget()
    frame3.pack_forget()
    frame4.pack_forget()
    frame5.pack_forget()
    frame6.pack_forget()
    frame7.pack_forget()
    frame8.pack_forget()
    frame9.pack_forget()
    frame10.pack_forget()
    frame0_1.pack_forget()
    frame1_1.pack_forget()
    frame2_1.pack_forget()
    frame3_1.pack_forget()
    frame4_1.pack_forget()
    frame5_1.pack_forget()
    frame6_1.pack_forget()
    frame7_1.pack_forget()
    frame8_1.pack_forget()
    frame9_1.pack_forget()
    frame10_1.pack_forget()

def show_widget_Group_of_senario_1():
    frame0.pack(ipadx=200, ipady=10, expand=True)
    frame1.pack(ipadx=200, ipady=10, expand=True)
    frame2.pack(ipadx=200, ipady=10, expand=True)
    frame3.pack(ipadx=200, ipady=10, expand=True)
    frame4.pack(ipadx=200, ipady=10, expand=True)
    frame5.pack(ipadx=200, ipady=10, expand=True)
    frame6.pack(ipadx=200, ipady=10, expand=True)
    frame7.pack(ipadx=200, ipady=10, expand=True)
    frame8.pack(ipadx=200, ipady=10, expand=True)
    frame9.pack(ipadx=200, ipady=10, expand=True)
    frame10.pack(ipadx=200, ipady=10, expand=True)

def show_widget_Group_of_senario_2():
    frame0_1.pack(ipadx=200, ipady=10, expand=True)
    frame1_1.pack(ipadx=200, ipady=10, expand=True)
    frame2_1.pack(ipadx=200, ipady=10, expand=True)
    frame3_1.pack(ipadx=200, ipady=10, expand=True)
    frame4_1.pack(ipadx=200, ipady=10, expand=True)
    frame5_1.pack(ipadx=200, ipady=10, expand=True)
    frame6_1.pack(ipadx=200, ipady=10, expand=True)
    frame7_1.pack(ipadx=200, ipady=10, expand=True)
    frame8_1.pack(ipadx=200, ipady=10, expand=True)
    frame9_1.pack(ipadx=200, ipady=10, expand=True)
    frame10_1.pack(ipadx=200, ipady=10, expand=True)

b1 = tk.Button(Top_frame, bg="red", fg="yellow", text="Hide All", command=hide_widget)
b1.pack(side=LEFT)
b2 = tk.Button(Top_frame, bg="yellow", fg="blue", text="Show All", command=show_widget)
b2.pack(side=LEFT)
b3 = tk.Button(Top_frame, fg="red", text="Group of senario 1", command=show_widget_Group_of_senario_1)
b3.pack()
b4 = tk.Button(Top_frame, fg="red", text="Group of senario 2", command=show_widget_Group_of_senario_2)
b4.pack()

def export_to_excel():
    workbook = xlwt.Workbook()
    header_font = xlwt.Font()
    header_font.name = 'Arial'
    header_font.bold = True
    al = Alignment()
    al.horz = Alignment.HORZ_CENTER
    header_style = xlwt.XFStyle()
    header_style.font = header_font
    header_style.alignment = al

    colour_style = xlwt.easyxf('align: wrap yes,vert centre, horiz center;pattern: pattern solid, \
                                           fore-colour light_orange;border: left thin,right thin,top thin,bottom thin')
    base_style = xlwt.easyxf('align: wrap yes,vert centre, horiz left; pattern: pattern solid, \
                                             fore-colour white;border: left thin,right thin,top thin,bottom thin')
    base_2_style = xlwt.easyxf('align: wrap yes,vert centre, horiz center;pattern: pattern solid, \
                                               fore-colour white;border: left thin,right thin,top thin,bottom thin')
    float_style = xlwt.easyxf('align: wrap yes,vert centre, horiz right ; pattern: pattern solid,\
                                              fore-colour light_yellow;border: left thin,right thin,top thin,bottom thin')
    date_style = xlwt.easyxf('align: wrap yes; pattern: pattern solid,fore-colour light_yellow;border: left thin,right thin,top thin,bottom thin\
                                             ', num_format_str='YYYY-MM-DD')
    datetime_style = xlwt.easyxf('align: wrap yes; pattern: pattern solid, fore-colour light_yellow;\
                                                 protection:formula_hidden yes;border: left thin,right thin,top thin,bottom thin',
                                 num_format_str='YYYY-MM-DD HH:mm:SS')


    sheet = workbook.add_sheet('Group of senario 1')
    sheet.write(0, 0, "Group of senario 1", header_style)
    sheet.write(0, 1, "Description", header_style)
    sheet.write(0, 2, "Results", header_style)
    sheet.write(1, 0, "Senario 1", header_style)
    sheet.write(1, 1, "1)Enter description of senario here\n2)Enter description of senario here", base_2_style)
    sheet.write(1, 2, v1.get(), colour_style)
    sheet.write(2, 0, "Senario 2", header_style)
    sheet.write(2, 1, "1)Enter description of senario here\n2)Enter description of senario here", base_2_style)
    sheet.write(2, 2, v2.get(), colour_style)
    sheet.write(3, 0, "Senario 3", header_style)
    sheet.write(3, 1, "1)Enter description of senario here\n2)Enter description of senario here", base_2_style)
    sheet.write(3, 2, v3.get(), colour_style)
    sheet.write(4, 0, "Senario 4", header_style)
    sheet.write(4, 1, "1)Enter description of senario here\n2)Enter description of senario here", base_2_style)
    sheet.write(4, 2, v4.get(), colour_style)
    sheet.write(5, 0, "Senario 5", header_style)
    sheet.write(5, 1, "1)Enter description of senario here\n2)Enter description of senario here", base_2_style)
    sheet.write(5, 2, v5.get(), colour_style)
    sheet.write(6, 0, "Senario 6", header_style)
    sheet.write(6, 1, "1)Enter description of senario here\n2)Enter description of senario here", base_2_style)
    sheet.write(6, 2, v6.get(), colour_style)
    sheet.write(7, 0, "Senario 7", header_style)
    sheet.write(7, 1, "1)Enter description of senario here\n2)Enter description of senario here", base_2_style)
    sheet.write(7, 2, v7.get(), colour_style)
    sheet.write(8, 0, "Senario 8", header_style)
    sheet.write(8, 1, "1)Enter description of senario here\n2)Enter description of senario here", base_2_style)
    sheet.write(8, 2, v8.get(), colour_style)
    sheet.write(9, 0, "Senario 9", header_style)
    sheet.write(9, 1, "1)Enter description of senario here\n2)Enter description of senario here", base_2_style)
    sheet.write(9, 2, v9.get(), colour_style)
    sheet.write(10, 0, "Senario 10", header_style)
    sheet.write(10, 1, "1)Enter description of senario here\n2)Enter description of senario here", base_2_style)
    sheet.write(10, 2, v10.get(), colour_style)
    sheet.col(0).width = 8000
    sheet.col(1).width = 8000
    sheet.col(2).width = 8000
    sheet.col(3).width = 8000
    sheet.col(4).width = 8000

    sheet1 = workbook.add_sheet('Group of senario 2')
    sheet1.write(0, 0, "Group of senario 2", header_style)
    sheet1.write(0, 1, "Description", header_style)
    sheet1.write(0, 2, "Results", header_style)
    sheet1.write(1, 0, "Senario 1", header_style)
    sheet1.write(1, 1, "1)Enter description of senario here\n2)Enter description of senario here", base_2_style)
    sheet1.write(1, 2, v11.get(), colour_style)
    sheet1.write(2, 0, "Senario 2", header_style)
    sheet1.write(2, 1, "1)Enter description of senario here\n2)Enter description of senario here", base_2_style)
    sheet1.write(2, 2, v12.get(), colour_style)
    sheet1.write(3, 0, "Senario 3", header_style)
    sheet1.write(3, 1, "1)Enter description of senario here\n2)Enter description of senario here", base_2_style)
    sheet1.write(3, 2, v13.get(), colour_style)
    sheet1.write(4, 0, "Senario 4", header_style)
    sheet1.write(4, 1, "1)Enter description of senario here\n2)Enter description of senario here", base_2_style)
    sheet1.write(4, 2, v14.get(), colour_style)
    sheet1.write(5, 0, "Senario 5", header_style)
    sheet1.write(5, 1, "1)Enter description of senario here\n2)Enter description of senario here", base_2_style)
    sheet1.write(5, 2, v15.get(), colour_style)
    sheet1.write(6, 0, "Senario 6", header_style)
    sheet1.write(6, 1, "1)Enter description of senario here\n2)Enter description of senario here", base_2_style)
    sheet1.write(6, 2, v16.get(), colour_style)
    sheet1.write(7, 0, "Senario 7", header_style)
    sheet1.write(7, 1, "1)Enter description of senario here\n2)Enter description of senario here", base_2_style)
    sheet1.write(7, 2, v17.get(), colour_style)
    sheet1.write(8, 0, "Senario 8", header_style)
    sheet1.write(8, 1, "1)Enter description of senario here\n2)Enter description of senario here", base_2_style)
    sheet1.write(8, 2, v18.get(), colour_style)
    sheet1.write(9, 0, "Senario 9", header_style)
    sheet1.write(9, 1, "1)Enter description of senario here\n2)Enter description of senario here", base_2_style)
    sheet1.write(9, 2, v19.get(), colour_style)
    sheet1.write(10, 0, "Senario 10", header_style)
    sheet1.write(10, 1, "1)Enter description of senario here\n2)Enter description of senario here", base_2_style)
    sheet1.write(10, 2, v20.get(), colour_style)
    sheet1.col(0).width = 8000
    sheet1.col(1).width = 8000
    sheet1.col(2).width = 8000
    sheet1.col(3).width = 8000
    sheet1.col(4).width = 8000
    workbook.save('QA Results.xls')

def senario_1():
    time.sleep(1)
    start = pyautogui.locateCenterOnScreen("images/start.png", confidence=0.9)
    if start is None:
        messagebox.showinfo("Error", "Process not completed")
        return
    else:
        pyautogui.click(start, button='left')
    time.sleep(1)
    recycle_bin = pyautogui.locateCenterOnScreen("images/recycle bin.png", confidence=0.8)
    if recycle_bin is None:
        messagebox.showinfo("Error", "Process not completed")
        return
    else:
        pyautogui.click(recycle_bin, button='left')

def donothing():
    filewin = Toplevel(root)
    button = Button(filewin, text="Do nothing button")
    button.pack()

def clear():
    label_R1.config(text="")
    label_R2.config(text="")
    label_R3.config(text="")
    label_R4.config(text="")
    label_R5.config(text="")
    label_R6.config(text="")
    label_R7.config(text="")
    label_R8.config(text="")
    label_R9.config(text="")
    label_R10.config(text="")
    label_R1_1.config(text="")
    label_R2_1.config(text="")
    label_R3_1.config(text="")
    label_R4_1.config(text="")
    label_R5_1.config(text="")
    label_R6_1.config(text="")
    label_R7_1.config(text="")
    label_R8_1.config(text="")
    label_R9_1.config(text="")
    label_R10_1.config(text="")
    return (v1.set(5), v2.set(5), v3.set(5), v4.set(5), v5.set(5), v6.set(5),
            v7.set(5), v8.set(5), v9.set(5), v10.set(5), v11.set(5), v12.set(5),
            v13.set(5), v14.set(5), v15.set(5), v16.set(5), v17.set(5), v18.set(5),
            v19.set(5), v20.set(5))

menubar = Menu(main_frame)
filemenu = Menu(menubar, tearoff=0)
filemenu.add_command(label="Export to excel", command=lambda: export_to_excel())
filemenu.add_command(label="Clear", command=lambda: clear())
filemenu.add_separator()
filemenu.add_command(label="Exit", command=lambda: on_closing())
menubar.add_cascade(label="File", menu=filemenu)
editmenu = Menu(menubar, tearoff=0)
editmenu.add_command(label="Undo", command=donothing)

root.config(menu=menubar)

def on_closing():
    if messagebox.askokcancel("Quit", "Do you want to quit?"):
        root.destroy()

root.protocol("WM_DELETE_WINDOW", on_closing)
root.mainloop()