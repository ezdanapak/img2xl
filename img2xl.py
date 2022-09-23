import openpyxl as xl, os, tkinter, exifread, sv_ttk
from tkinter import TOP, filedialog, messagebox

class imgtoxl:
    def __init__(self, geometry, resizable1, resizable2, title, theme):
        self.window=window
        self.window.geometry(geometry)
        self.window.resizable(resizable1,resizable2)
        self.window.title(title)
        sv_ttk.set_theme(theme)
        self.path=''

    def choosefolder(self):
        self.path=filedialog.askdirectory()
        self.folderlabel.config(text=self.path)
        
    def mainfunc(self):
        self.newpath=''
        try:
            for i in self.path:
                if i=='/':
                    i='\\\\'
                self.newpath+=i
        except:
            pass     

        if os.path.isdir(self.newpath):    
            self.datalist=[]

            self.filetype='.JPG'
            os.chdir(self.newpath)

            for file in os.listdir():
                self.ext=os.path.splitext(file)[1]
                if self.ext==self.filetype:
                    tags=exifread.process_file(open(file, 'rb'))
                    geo={i:tags[i] for i in tags.keys() if i.startswith('GPS')}

                    lat=str(geo['GPS GPSLatitude'])
                    lat=lat.lstrip('[')
                    lat=lat.rstrip(']')
                    latlist=lat.split(', ')
                    latdeg=int(latlist[0])
                    latmin=int(latlist[1])
                    latsec=int(latlist[2].split('/')[0])/int(latlist[2].split('/')[1])
                    declat=(latsec/60+latmin)/60+latdeg
                    
                    lon=str(geo['GPS GPSLongitude'])
                    lon=lon.lstrip('[')
                    lon=lon.rstrip(']')
                    lonlist=lon.split(', ')
                    londeg=int(lonlist[0])
                    lonmin=int(lonlist[1])
                    lonsec=int(lonlist[2].split('/')[0])/int(lonlist[2].split('/')[1])
                    declon=(lonsec/60+lonmin)/60+londeg
                    
                    alt=str(geo['GPS GPSAltitude'])
                    altlist=alt.split('/')
                    alt=int(altlist[0])/int(altlist[1])

                    row=(file, declat, declon, alt)
                    self.datalist.append(row)
            
            if len(self.datalist)>0:
                wb=xl.Workbook()
                sheet=wb.active
                for i in self.datalist:
                    sheet.append(i)
                try:
                    wb.save('Image_coordinates.xlsx')
                    messagebox.showinfo('Infobox','Process complete')
                    self.window.destroy()
                except PermissionError:
                    messagebox.showerror('Error', 'Please close excel file and try again')
            else:
                messagebox.showerror('Error', 'Folder contains no images or images have no cooridnates')        
        else:
            messagebox.showerror('Error', 'Please choose a valid folder')
    
    def choosebtn(self):
        self.choosebutton=tkinter.Button(self.window, text='Choose folder', command=lambda:app.choosefolder())
        self.choosebutton.place(x=266,y=30)

    def folderlbl(self):
        if os.path.isdir(self.path):
            self.folderlabel=tkinter.Label(window, text=app.choosefolder)
            self.folderlabel.pack(side=TOP)
        else:
            self.folderlabel=tkinter.Label(window, text='Folder path will be shown here')
            self.folderlabel.pack(side=TOP)

    def runbtn(self):
        self.runbutton=tkinter.Button(self.window, text='Run', font='Bold 12', padx=30, pady=20, command=app.mainfunc)
        self.runbutton.place(x=260, y=150)

window=tkinter.Tk()

#initialize with window size, resizable parameters, window label and theme paramaters
app=imgtoxl('600x360', 0,0, 'IMG2XL', 'dark')
app.choosebtn()
app.folderlbl()
app.runbtn()

window.mainloop()
