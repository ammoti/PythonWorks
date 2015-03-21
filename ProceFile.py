from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from mmap import mmap,ACCESS_READ
import tkinter.font as tkFont
from xlrd import open_workbook
import recommendations
from os.path import basename

class ProgramULagnet(Frame):
    def __init__(self):
        Frame.__init__(self)
        w = Canvas(root, width=500, height=400)
        w.create_text(250, 30, text="Virtual Advisor v1.0", font=("Bold", 14))
        #w.create_oval(10, 50, 70, 90, width=2, fill='blue')

        car_header = ['car', 'repair']
        car_list = [('Hyundai', 'brakes') ,('Honda', 'light') ,('Lexus', 'battery') , ('Benz', 'wiper') ,('Ford', 'tire') ,('Chevy', 'air') , ('Chrysler', 'piston') , ('Toyota', 'brake pedal') , ('BMW', 'seat')]

        w.create_rectangle(490, 45, 10, 10)
        w.create_rectangle(50, 100, 490, 150)
        w.create_text(110, 125, text="Collobrative \nFiltering Type : ", font=("Bold", 10))
        w.create_text(350, 125, text="Similiratiy \nMeasure : ", font=("Bold", 10))

        tree =ttk.Treeview(root,selectmode="extended",columns=("FirstColumn","SecondColumn"))
        tree.pack(expand=YES, fill=BOTH)
        tree.heading("#0", text="")
        tree.column("#0",minwidth=0,width=0, stretch=NO)
        tree.heading("FirstColumn", text="Reccommended Course")
        tree.column("FirstColumn",minwidth=0,width=200, stretch=NO)
        tree.heading("SecondColumn", text="Predicted Grade")
        tree.column("SecondColumn",minwidth=0,width=300)
        tree.tag_configure("column",foreground="red")

        tree.place(x=0,y=350)
      #  Style().configure("Treeview", background="#383838", foreground="white", fieldbackground="red")
        def hello(vapi):
            w.create_text(350, 70, text=vapi, font=("Bold", 14))

        def seeReccomondion(self):
            currentValue = c

        Radiobutton(root,
            text="User-Based",
            value=1).place(x=170,y=105)
        Radiobutton(root,
            text="Item-Based",
            value=2).place(x=170,y=125)
        countryvar = ('Toronto', 'Ottawa', 'Montreal', 'Vancouver', 'St. John')
        cmdBox = ttk.Combobox(root)
        # values=countryvar,width=13).place(x=380,y=120)
        cmdBox.place(x=380,y=120)
        cmdBox["values"]=("Pearson","Euclidian","Jaccard")
        #cmdBox.bind("<<ComboboxSelected>>", defocus)

        button2=Button(root, text='Load your current grades')
        button2.place(x=70,y=250)
        button3=Button(root, text='See the Recommended Courses', command=self.seeReccomondion)
        button3.place(x=100,y=300)
        button=Button(root, text='Load Past Student Grades',command=self.askopenfile)
        button.place(x=30,y=200)
        w.pack()



    def askopenfile(self):
            self.file_opt = options = {}
            options['initialdir'] = 'C:\\'
            options['parent'] = root
            options['title'] = 'This is a title'
            """Returns an opened file in read mode."""
            usame = filedialog.askopenfiles(filetypes = (("Excel files", "*.xls;*.xlsx"),("All files", ".*")))

            datas={}

            for files in usame:
                result ={}
                dosya = files.name
                fileName = basename(dosya)
                rows = []
                wb = open_workbook(dosya)
                sh = wb.sheet_by_index(0)
                for rownum in range(sh.nrows):
                    rows.append(sh.row_values(rownum))
                rows.__delitem__(0)

                courseDatas={}
                for itemDatas in rows:
                    courseName = itemDatas[0]+" "+itemDatas[1]
                    courseGrade = itemDatas[2]
                    courseDatas[courseName]= courseGrade
                datas[fileName] = courseDatas


    def read_rows(inputfile):
        rows = []
        wb = open_workbook(inputfile)
        sh = wb.sheet_by_index(0)
        for rownum in range(sh.nrows):
            rows.append(sh.row_values(rownum))
            return rows

if __name__=='__main__':
    root = Tk()
    root.configure(width=500,height=400)
    ProgramULagnet().mainloop()