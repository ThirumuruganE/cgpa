#!/usr/bin/env python3
from tkinter import *
from tkinter.messagebox import showerror, showinfo

root = Tk()
root.title("CGPA Calcuator")
rightframe = Frame(root, borderwidth=3,width =454)
rightframe.pack(side=RIGHT)
leftframe = Frame(root, borderwidth=3, bg='grey')
leftframe.pack(side=LEFT)


def calc():
    import pandas as pd
    global file1
    s = R1.get()
    p = R2.get()
    if s == 1:
        file1 = "r2015.xls"
    elif p == 1:
        file1 = "r2019.xls"
    op = pd.read_excel(file1)
    cr = (op.loc[:, 'course title'])
    c = (op.loc[:, 'grade'])
    d = (op.loc[:, 'credits'])

    total_number = cr.count()
    print(total_number)
    global f, u
    f = 0
    u = 0
    for i in range(0, total_number):
        if c[i] in range(5, 11):
            f = f + c[i] * d[i]
            print(f)
            u = u + d[i]
            print(u)
    tot = f / u
    u = int(u)
    if s == 1:
        showinfo("CGPA-r2015", "your current CGPA is {:.2f} and\nno of credits completed  is %d/81".format(tot) % (u))
    elif p == 1:
        showinfo("CGPA-r2019", "your current CGPA is {:.2f} and\nno of credits completed  is %d/85".format(tot) % (u))


R1 = IntVar()
R2 = IntVar()
chk_1 = Checkbutton(leftframe, text='R2015', variable=R1)
chk_1.pack()

chk_2 = Checkbutton(leftframe, text='R2019', variable=R2)
chk_2.pack()

but_1 = Button(leftframe, text='Calculate_CGPA', width=20, height=1, borderwidth=3, command=calc)
but_1.pack()


def lic():
    new_frame2 = Toplevel(root)
    new_frame2.title('License')
    scr = Scrollbar(new_frame2)
    tex1 = Text(new_frame2, height=30, width=70)
    scr.pack(side=RIGHT, fill=Y)
    tex1.pack(side=LEFT, fill=Y)
    scr.config(command=tex1.yview)
    tex1.config(yscrollcommand=scr.set)
    lic1 = '''MIT License
    \nCopyright (c) 2020 Thirumurugan Elango\nPermission is hereby granted, free of charge, to any person 
obtaining a copy of this software and associated documentation files (the "cgpa_calculator"), to deal in the Software
without restriction,including without limitation the rights to use, copy, modify, merge,  publish, distribute, 
sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do 
so, subject to the following conditions:
    
The above copyright notice and this permission notice shall be included in all copies or substantial 
portions of the Software.
    
THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE 
WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR 
COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR 
OTHERWISE, ARISING FROM,OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.'''
    tex1.insert(END, lic1)


scroll = Scrollbar(rightframe)
tex2 = Text(rightframe, height=30, width=80)
scroll.pack(side=RIGHT, fill=Y)
tex2.pack(side=LEFT, fill=Y)
scroll.config(command=tex2.yview)
tex2.config(yscrollcommand=scroll.set)
img = PhotoImage(file="cgpa.gif")
tex2.image_create(END, image=img)
readme = '''
Can be used to calculate current CGPA,
INSTRUCTIONS
1) If you belong to regulation 2015 fill the grade column of r2015.xls
2) If a course is not completed leave that cell of the grade column blank
3) select which regulation you belong to check box
4) click calculate button
5) Or If you belong to 2019 regulation fill r2019.xls and select R2019 option
6) Click calculate
'''
tex2.insert(END, readme)
from prettytable import PrettyTable

xy = PrettyTable()
xy.field_names = ['Letter Grade', 'Grade Points']
xy.add_row(["O (Outstanding)", 10])
xy.add_row(['A + (Excellent)', 9])
xy.add_row(['A (Very Good)  ', 8])
xy.add_row(['B + (Good)     ', 7])
xy.add_row(['B (Average)    ', 6])
tex2.insert(END, xy)
formu = '''
Formula used for calculation '''
tex2.insert(END, formu)
img2 = PhotoImage(file="cgppa.png")
img2 = img2.subsample(1,1)
tex2.image_create(END, image=img2)

def helping():
    new_wind = Toplevel(root)
    new_wind.title("help/contact")
    tex3 = Text(new_wind,height=10,width=50)
    tex3.pack()
    contact ='''
    Author: Thirumurugan Elango
    Masters Student,
    Department of Medical Physics,
    Anna University
    email: thiru20@protonmail.com'''
    tex3.insert(END,contact)
menubar = Menu(root)
root.config(menu=menubar)
submenu = Menu(menubar, tearoff=0)
menubar.add_cascade(label="about", menu=submenu)
submenu.add_command(label="License", command=lic)
helpmenu= Menu(menubar,tearoff=0)
menubar.add_cascade(label="Help",menu=helpmenu)
helpmenu.add_command(label='help/contact',command=helping)
root.mainloop()
