from tkinter import *
import tkinter as tk
from tkinter import filedialog
from xml.sax.saxutils import unescape
#from PIL import Image, ImageTk


class App:
    in_path=""
    out_path=""
    html_src=""
    caption=""
    extract_info=""
    caldic={}
    


    def __init__(self, master):
        frame = Frame(master)
        frame.pack()


        self.lbl=Label(frame,text="Tableau Metadata Reader Tool", font=(16)).grid(row=0, column=0, columnspan=3, padx=10, pady=5,sticky=W)
       # self.lbl.pack(side=LEFT)

        # path="C:/Users/nitin/Desktop/hig.jpg"
        # img=ImageTk.PhotoImage(Image.open(path))
        #
        # self.lbl_img=Label(frame,image=img).grid(row=0, column=1)
        #
        # #self.lbl_img.pack(side=LEFT)

       # self.blank_lbl=Label(frame,text="").grid(row=0,column=2)
       # self.blank_lbl.pack(side=BOTTOM)

        self.open_btn = Button(frame, text="OPEN", command=self.open_file)
        self.open_btn.grid(row=1,column=0, padx=5, pady=5, sticky=W)
     #   self.open_btn.pack(side=LEFT , padx=40, pady=10)

        self.html_btn=Button(frame,text="HTML", command=self.convert_to_html)
        self.html_btn.grid(row=1, column=1, padx=5, pady=5)
     #   self.html_btn.pack(side=LEFT)

        self.button = Button(frame, text="QUIT", command=frame.quit )
        self.button.grid(row=1,column=2, padx=5, pady=5, sticky=E)
        #self.button.pack(side=LEFT)

        self.lbl_file=Label(frame, text="Open TWB file")
        self.lbl_file.grid(row=2, column=0, columnspan=3, padx=10, pady=5,sticky=W)
       # self.lbl_file.pack()

    def open_file(self):
       # tk.Tk().withdraw()  # Close the root window
        root.filename = filedialog.askopenfilename(initialdir="/", title="Select file",
                                                     filetypes=(("Tableau Workbook", "*.twb"), ("all files", "*.*")))
        self.in_path=root.filename
        self.lbl_file.configure(text="File Opened: "+self.in_path)

        print(self.in_path)

    def save_file(self):
        root.filename=filedialog.asksaveasfilename(initialdir="/",title="Select File",
                                                     filetypes=(("HTML", "*.html"), ("all files", "*.*")))

 

    
    #def replace_all(text):
        #for i, j in self.caldic.items():
            #if i:
                #text = text.replace(i, j)
            #print "text--  "+text
        #return text
        

    def convert_to_html(self):

        print("convert to html input path file: "+ self.in_path)
        self.save_file()

        if not root.filename:
            self.lbl_file.configure(text="Open Correct File")
        else:
            #if not root.filename.endswith(".html"):
            self.out_path = root.filename +".html"

            self.lbl_file.configure(text="File saved at: "+self.out_path)

            print(self.out_path)
            f = open(self.out_path, 'w')
            self.read_twb()
            #print("src : " + self.html_src)
            message = """<html>
                    <head></head>
                    <body><p>""" + self.html_src + """</p></body>
                    </html>"""
            f.write(message)
            f.close()


    def read_twb(self):
        dash_name = "<table border=\"1\" width=\"100%\"><tr><td bgcolor=\"#98B5E3\"><b>ID</b></td><td bgcolor=\"#98B5E3\"><b>DASHBOARD NAME</b></td></tr>"
        worksht_name = "<table border=\"1\" width=\"100%\"><tr><td bgcolor=\"#98B5E3\"><b>ID</b></td><td bgcolor=\"#98B5E3\"><b>WORKSHEET NAME</b></td></tr>"
        version_info = 0
        
        DLM = 0
        Calcs = ['0']
        Cleaner = ['0']
        WLM = 0
        k = 0

        self.html_src = '<h2>VERSION INFORMATION</h2>'

        with open(self.in_path) as filer3:
            for line in filer3:
                if line.find('datasource caption')>-1 and line.find('inline')>-1:
                    line=filer3.readline()
                    line=filer3.readline()
                    while line.find('</metadata-records>')<0:
                        prev=line
                        line=filer3.readline()
                    while line.find('</datasource>')<0 :

                        prev1=line
                        line=filer3.readline()
                        if line.find('<calculation class=')>-1:
                            if prev1.find('param-domain-type')<0:
                                if prev1[prev1.find('name')+6:prev1.find('role')-2] :
                                    self.caldic[prev1[prev1.find('name')+6:prev1.find('role')-2]] = prev1[prev1.find('caption')+9:prev1.find('datatype')-2]            
                                    #calFormula[prev1[prev1.find('caption')+9:prev1.find('datatype')-2]]= unescape(line[line.find('formula')+9:line.find(' />')-1], {"&apos;": "'", "&quot;": '"', "&#13;":" ", "&#10;":" "})

        filer3.close()

        
        with open(self.in_path) as filer:
            for line in filer:
                ##read version information
                if line.find('<!-- build') > -1:
                    version_info  = "<table border=\"1\" width=\"100%\"><tr><td bgcolor=\"#98B5E3\"><b>FILENAME</b></td><td bgcolor=\"#98B5E3\"><b>TABLEAU BUILD</b></td><td bgcolor=\"#98B5E3\"><b>PLATFORM</b></td>"
                    version_info = version_info  + "<td bgcolor=\"#98B5E3\"><b>TABLEAU VERSION</b></td></tr>"
                    version_info = version_info  + '<tr><td style=\"text-align:center;\">' + self.in_path + '</td>'
                    version_info = version_info  + '<td style=\"text-align:center;\">' + line[11:11 + 19] + '</td>'
                if line.find('source-platform=') > -1:
                    version_info = version_info  + '<td style=\"text-align:center;\">' + line[line.find('source-platform=') + 17:line.find('source-platform=') + 20] + '</td>'
                if line.find('source-platform') > -1:
                    version_info  = version_info  + '<td style=\"text-align:center;\">' + line[line.find('version=') + 9:line.find('version=') + 13] + '</td></tr></table>'
                    self.html_src = self.html_src + version_info

                    ##Datasource information
                    self.html_src = self.html_src + '<h2>DATASOURCES</h2>'

                if line.find('datasource caption') > -1 and line.find('inline') > -1:
                    k = k + 1
                    self.html_src = self.html_src + '<h3>DATASOURCE ' + str(k) + ': ' + line[line.find(
                        'datasource caption') + 20:line.find('inline') - 2] + '</h3>'

                    data_src = "<table border=\"1\" width=\"100%\"><tr><td bgcolor=\"#98B5E3\"><b>DATASOURCE</b></td><td bgcolor=\"#98B5E3\"><b>TYPE</b></td><td bgcolor=\"#98B5E3\"><b>DRIVER VERSION</b></td>"
                    data_src = data_src + "<td bgcolor=\"#98B5E3\"><b>FILE NAME</b></td>"
                    data_src = data_src + "<td bgcolor=\"#98B5E3\"><b>SERVER</b></td><td bgcolor=\"#98B5E3\"><b>TABLE</b></td><td bgcolor=\"#98B5E3\"><b>TYPE</b></td></tr> "
                    data_src = data_src + '<tr><td>' + line[line.find('datasource caption') + 20:line.find(
                        'inline') - 2] + '</td>'
                    data_src = data_src + '<td>' + line[line.find('name') + 6:line.find('version') - 3] + '</td>'
                    data_src = data_src + '<td>' + line[line.find('version') + 9:line.find('version') + 12] + '</td>'

                    line = filer.readline()
                    if line.find('filename') > -1:
                        data_src = data_src + '<td>' + line[
                                                       line.find('filename') + 10:line.find('password') - 2] + '</td>'

                    if line.find('server') > -1:
                        data_src = data_src + '<td>' + line[
                                                       line.find('server') + 8:line.find('server-oauth') - 2] + '</td>'
                    line = filer.readline()
                    if line.find('table') > -1:
                        data_src = data_src + '<td>' + line[line.find('table') + 7:line.find('type') - 2] + '</td>'
                        data_src = data_src + '<td>' + line[
                                                       line.find('type') + 6:line.find('/>') - 2] + '</td></tr></table>'
                        self.html_src = self.html_src + data_src

                    ##Reading field name information

                    field_name = "<table border=\"1\" width=\"100%\"><tr><td bgcolor=\"#98B5E3\"><b>ID</b></td><td bgcolor=\"#98B5E3\"><b>FIELD NAME</b></td><td bgcolor=\"#98B5E3\"><b>TABLE</b></td><td bgcolor=\"#98B5E3\"><b>ALIAS</b></td><td bgcolor=\"#98B5E3\"><b>DATA TYPE</b></td></tr>"
                    index=1
                    while line.find('</metadata-records>') < 0:
                        prev = line
                        line = filer.readline()
                        if line.find('<remote-name') > -1:
                            calculated_field_name= line[line.find('remote-name') + 12:line.find('</remote')]
                            if calculated_field_name=="/>":
                                break
                            field_name=field_name+ '<tr><td style=\"text-align:center;\">' + str(index) + '</td>'
                            index+=1
							
                            field_name = field_name + '<td style=\"text-align:center;\">' + calculated_field_name + '</td>'
                        if line.find('<remote-alias') > -1:
                            field_name = field_name + '<td style=\"text-align:center;\">' + line[line.find('remote-alias') + 13:line.find(
                                '</remote')] + '</td>'
                        if line.find('<parent-name') > -1:
                            field_name = field_name + '<td style=\"text-align:center;\">' + line[line.find('parent-name') + 12:line.find(
                                '</parent')] + '</td>'
                        if line.find('<local-type') > -1:
                            field_name = field_name + '<td style=\"text-align:center;\">' + line[line.find('local-type') + 11:line.find(
                                '</local')] + '</td></tr>'
                    self.html_src = self.html_src + field_name + '</table>'

                    ##Calculation information
                    self.html_src = self.html_src + '<h2>CALCULATIONS</h2>'

                    CAL = "<table border=\"1\" width=\"100%\"><tr><td bgcolor=\"#98B5E3\"><b>ID</b></td><td bgcolor=\"#98B5E3\"><b>CALCULATED FIELD NAME</b></td><td bgcolor=\"#98B5E3\"><b>DATA TYPE</b></td><td bgcolor=\"#98B5E3\">"
                    CAL = CAL + "<b>ALIAS</b></td><td bgcolor=\"#98B5E3\"><b>ROLE</b></td><td bgcolor=\"#98B5E3\"><b>FORMULA</b></td></tr>"

                    index=1
                    while line.find('</datasource>') < 0:
                        prev = line
                        line = filer.readline()
                       
                        if line.find('<calculation class=') > -1:
                            if prev.find('param-domain-type') < 0:
                                Calcs.append(prev[prev.find('name') + 6:prev.find('role') - 2] + "|" + prev[prev.find(
                                    'caption') + 9:prev.find('datatype') - 2])
                                Cleaner.append(line[line.find('formula') + 9:line.find(' />') - 1] + '\n')
                                
                                caption=prev[prev.find('caption') + 9:prev.find('datatype') - 2]
                                if caption == "olum" :
                                    break
                                CAL = CAL + "<tr><td>" + str(index) + "</td>"
                                index+=1
                                CAL = CAL + "<td>" + caption  + "</td>"
                                CAL = CAL + "<td>" + prev[prev.find('datatype') + 10:prev.find('name') - 2] + "</td>"
                                CAL = CAL + "<td>" + prev[prev.find('name') + 6:prev.find('role') - 2] + "</td>"
                                CAL = CAL + "<td>" + prev[prev.find('role') + 6:prev.find(' type=') - 1] + "</td>"

                                formula=unescape(line[line.find('formula')+9:line.find(' />')-1], {"&apos;": "'", "&quot;": '"', "&#13;":" ", "&#10;":" "})
                                for i, j in self.caldic.items():
                                    if i:
                                        formula = formula.replace(i, j)
                                #formula=self.replace_all(formula)
                                CAL = CAL + "<td>" + formula  + "</td></tr>"
                                
                                #CAL = CAL + "<td>" + line[line.find('formula') + 9:line.find(' />') - 1] + "</td></tr>"

                        ##Extract Information
                        if line.find('<extract count') > -1:
                            line = filer.readline()
                            self.extract_info= '<table border=\"1\" width=\"100%\"><tr><td bgcolor=\"#98B5E3\"><b>LAST EXTRACT DATE:</b></td><td>' + line[
                                                                                                                            line.find(
                                                                                                                                'update-time') + 12:line.find(
                                                                                                                                '>')] + '</td></tr></table>'


                    self.html_src = self.html_src + CAL + "</table>"

 

                    CAL = ""
                    #print('\n')


                if line.find('dashboard name') > -1:
                    DLM += 1
                    dash_name = dash_name + '<tr><td>' + str(DLM) + '</td><td>' + line[line.find('name') + 6:line.find(
                        '>') - 1] + '</td></tr>'

                if line.find('worksheet name') > -1:
                    WLM += 1

                    worksht_name = worksht_name + '<tr><td>' + str(WLM) + '</td><td>' + line[
                                                                                        line.find('name') + 6:line.find(

                                                                                            '>') - 1] + '</td></tr>'
        self.html_src = self.html_src + '<h2>EXTRACT INFORMATION</h2>'
        if self.extract_info:
            self.html_src = self.html_src + self.extract_info


        self.html_src = self.html_src + '<h2>DASHBOARDS</h2>'
        self.html_src = self.html_src + dash_name + "</table>"

        self.html_src = self.html_src + '<h2>WORKSHEETS</h2>'
        self.html_src = self.html_src+ worksht_name + "</table>"
        filer.close()

        ##Parameter information
        self.html_src = self.html_src + '<h2>PARAMETERS</h2>'

        with open(self.in_path) as filer2:
            index=1
            for line in filer2:
                if line.find('datasource hasconnection') > -1 and line.find('Parameters') > -1:
                    D = "<table border=\"1\" width=\"100%\"><tr><td bgcolor=\"#98B5E3\"><b>ID</b></td><td bgcolor=\"#98B5E3\"><b>PARAMETER NAME</b></td><td bgcolor=\"#98B5E3\"><b>VALUES</b></td><td bgcolor=\"#98B5E3\"><b>"
                    D = D + "MINIMUM VALUE</b></td><td bgcolor=\"#98B5E3\"><b>MAXIMUM VALUE</b></tr>"
                    while line.find('</datasource>') < 0:
                        line = filer2.readline()
                        if line.find('caption') > -1:
                            k = 1
                            D = D + '<tr><td><b>' + str(index)+ '</b></td>'+ '<td colspan=\"4\" ><b>' + line[line.find('caption') + 9:line.find(
                                'datatype') - 2] + '</b></td></tr>'
                            index+=1
                            while line.find('</column') < 0:
                                line = filer2.readline()

                                if line.find('<range granularity=') > -1:
                                    D = D + '<tr><td></td><td></td><td>' + line[line.find('min=') + 5: line.find(
                                        '/>') - 2] + '</td>'
                                    D = D + '<td>' + line[line.find('max=') + 5: line.find('min') - 2] + '</td>'
                                    D = D + '<td>' + line[line.find('ity=') + 5: line.find('max') - 2] + '</td></tr>'
                                if line.find('<alias key') > -1:
                                    D = D + '<tr><td>    Option ' + str(k) + '</td><td>' + str(line[line.find(
                                        'value') + 7:line.find('/>') - 2]).replace('"',"") + '</td><td></td><td></td><td></td></tr>'
                                if line.find('<member value') > -1:
                                    D = D + '<tr><td>    Option ' + str(k) + '</td><td>' + (line[line.find(
                                        'value') + 6:line.find('/>') - 2]).replace('"',"") + '</td><td></td><td></td><td></td></tr>'
                                    k += 1

                    self.html_src = self.html_src + D + '</table>'

 

 


root = Tk()
root.title("Tableau Metadata Reader")
#root.geometry("600x500-400+50")
app = App(root)
root.resizable(0,0)
#root.iconbitmap("hig.ico")
root.mainloop()
root.destroy()  
