from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from mmap import mmap,ACCESS_READ
import tkinter.font as tkFont
from xlrd import open_workbook
from recommendations import *
from os.path import basename

class virtualAdvisor(Frame):
        def _init_(self):
            self.veriler={}
            radioValue=1
            def askopenpastgrade():
                    self.file_opt = options = {}
                    options['parent'] = root
                    excelFiles = filedialog.askopenfiles(filetypes = (("Excel files", "*.xls;*.xlsx"),("All files", ".*")))
                    for files in excelFiles:
                        save ={}
                        dosya_name = files.name
                        fileName = basename(dosya_name)
                        rows = []
                        wb = open_workbook(dosya_name)
                        shindex = wb.sheet_by_index(0)
                        for rownum in range(shindex.nrows):
                            rows.append(shindex.row_values(rownum))
                        rows.__delitem__(0)
                        self.veriler[fileName] = converter(rows)


            def converter(x):
                courses=dict()
                for itemmm in x:
                        courseName = itemmm[0]+" "+itemmm[1]
                        if itemmm[2] == 'A+':
                            itemmm[2] = 4.1
                        elif itemmm[2] == 'A':
                            itemmm[2] = 4.0
                        elif itemmm[2] == 'A-':
                            itemmm[2] = 3.7
                        elif itemmm[2] == 'B+':
                            itemmm[2] = 3.3
                        elif itemmm[2] == 'B':
                            itemmm[2] = 3.0
                        elif itemmm[2] == 'B-':
                            itemmm[2] = 2.7
                        elif itemmm[2] == 'C+':
                            itemmm[2] = 2.3
                        elif itemmm[2] == 'C':
                            itemmm[2] = 2.0
                        elif itemmm[2] == 'C-':
                            itemmm[2] = 1.7
                        elif itemmm[2] == 'D+':
                            itemmm[2] = 1.3
                        elif itemmm[2] == 'D':
                            itemmm[2] = 1.0
                        elif itemmm[2] == 'D-':
                            itemmm[2] = 0.5
                        else:
                            itemmm[2] = 0.0
                        courseGrade = itemmm[2]
                        courses[courseName]= courseGrade
                return courses

            def askopentranscript():
                self.file_opt = options = {}
                options['initialdir'] = 'C:\\'
                options['parent'] = root
                options['title'] = 'This is a title'
                """Returns an opened file in read mode."""
                excelFiles = filedialog.askopenfiles(filetypes = (("Excel files", "*.xls;*.xlsx"),("All files", ".*")))


                for files in excelFiles:
                    result ={}
                    dosya = files.name
                    fileName = basename(dosya)
                    self.useLaterAsPerson = fileName

                    rows = []
                    wb = open_workbook(dosya)
                    sh = wb.sheet_by_index(0)
                    for rownum in range(sh.nrows):
                        rows.append(sh.row_values(rownum))
                    rows.__delitem__(0)
                    self.veriler[fileName] = converter(rows)


            def cmdSeeRecomended():
                veriler = self.veriler
                finalResult = dict()
                comboValue = kutu.current()
                grade = list()
                name = list()
                if radioValue==1:
                    if comboValue ==0:
                        AllCourses = getRecommendations(self.veriler,self.useLaterAsPerson,sim_pearson)
                        First6Courses= AllCourses[:6]
                        for i in range(len(First6Courses)):
                            finalResult[First6Courses[i][0]]=First6Courses[i][1]
                            courseName =First6Courses[i][1]
                    elif comboValue ==1:
                        AllCourses = getRecommendations(veriler,self.useLaterAsPerson,sim_distance)
                        First6Courses= AllCourses[:6]
                        for i in range(len(First6Courses)):
                            finalResult[First6Courses[i][0]]=First6Courses[i][1]
                    elif comboValue ==2:
                        AllCourses = getRecommendations(veriler,self.useLaterAsPerson,sim_jaccard())
                        First6Courses= AllCourses[:6]
                        for i in range(len(First6Courses)):
                            finalResult[First6Courses[i][0]]=First6Courses[i][1]
                elif radioValue==2:
                    AllCourses = calculateSimilarItems(self.veriler,10)
                    AllCourses2 = getRecommendedItems(veriler,AllCourses,self.useLaterAsPerson)
                    First6Courses = AllCourses2[:6]
                    for i in range(len(First6Courses)):
                            courseGrade = First6Courses[i][0]
                            grade.append(self.writing(courseGrade))
                            finalResult[First6Courses[i][0]]=First6Courses[i][1]
                            courseName =First6Courses[i][1]
                            name.append(courseName)


                for item in finalResult:
                    self.treeView.insert("","end",values=(finalResult[item],item))


            birincibuton=Label(self,width=4,height=3)
            birincibuton.config(text="1")
            birincibuton.place(relx=0.12,rely=0.04)

            ikincibuton=Button(self,width=22,height=3)
            ikincibuton.config(text="Load Past Student Grade",command=askopenpastgrade)
            ikincibuton.place(relx=0.17,rely=0.04)

            ucuncubuton=Label(self,width=4,height=3)
            ucuncubuton.config(text="2")
            ucuncubuton.place(relx=0.41,rely=0.04)

            dorduncubuton=Button(self,width=23,height=3)
            dorduncubuton.config(text="Load Your Current Transcript", command=askopentranscript)
            dorduncubuton.place(relx=0.47,rely=0.04)

            besincibuton=Label(self,width=4,height=3)
            besincibuton.config(text="3")
            besincibuton.place(relx=0.12,rely=0.16)
            v = IntVar()
            v.set(1)
            languages = [
                ("User-Based",1),
                ("Item-Based",2),
            ]
            radioButtonCount = 0
            for txt, val in languages:
                items= Radiobutton(self,text=txt,padx = 20,variable=v,
                            value=val).place(relx=0.32,rely=0.16+radioButtonCount)
                radioButtonCount+=0.04

            altincibuton=Label(self,width=15,height=3)
            altincibuton.config(text="colloborative\n filtering:",bg="gray")
            altincibuton.place(relx=0.181,rely=0.16)

            yedincibuton=Label(self,width=15,height=3)
            yedincibuton.config(text="similarty Measure:",bg="gray")
            yedincibuton.place(relx=0.50,rely=0.158)

            St=StringVar()
            comboliste=["Pearson,Euclidean,Jaccard"]
            kutu=ttk.Combobox(width=8, textvariable=St,state="readonly")
            kutu["values"]=("Pearson","Euclidian","Jaccard")
            kutu.current(0)
            kutu.place(relx=0.65,rely=0.178)

            sekizincibuton=Label(self,width=4,height=3)
            sekizincibuton.config(text="4")
            sekizincibuton.place(relx=0.12,rely=0.25)

            dokuzuncubuton=Button(self,width=23,height=3)
            dokuzuncubuton.config(text="See The Recomended Course",command=cmdSeeRecomended)
            dokuzuncubuton.place(relx=0.18,rely=0.25)

            onuncubuton=Label(self,width=23,height=1)
            onuncubuton.config(text="Vitual Advisor v1.0")
            onuncubuton.place(relx=0.35,rely=00)

            self.treeView =ttk.Treeview(root,columns=("FirstColumn","SecondColumn"))
            self.treeView.pack()
            self.treeView.heading("#0", text="")
            self.treeView.column("#0",minwidth=0,width=0,)
            self.treeView.heading("FirstColumn", text="Reccommended Course")
            self.treeView.column("FirstColumn",minwidth=0,width=200)
            self.treeView.heading("SecondColumn", text="Predicted Grade")
            self.treeView.column("SecondColumn",minwidth=0,width=300)

            self.treeView.place(x=0,y=200)







root = Tk()
virtualAdvisor._init_(root)
root.configure(width=800,height=500)
root.mainloop()