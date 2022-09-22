import openpyxl as xl, os, tkinter, exifread, pyproj
from tkinter import filedialog, messagebox
from PIL.ExifTags import TAGS as tags

def choosefolder():
    global path
    path=filedialog.askdirectory()
    folderlabel.config(text=path)
    
    
def mainfunc():
    newpath=''
    try:
        for i in path:
            if i=='/':
                i='\\\\'
            newpath+=i
    except NameError:
        pass     

    if os.path.isdir(newpath):    
        datalist=[]

        filetype='.JPG'
        os.chdir(newpath)

        for file in os.listdir():
            ext=os.path.splitext(file)[1]
            if ext==filetype:
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
                

                '''
                projmet="+proj=utm +zone=38 +north +datum=WGS84 +units=m +no_defs "
                mycoords=pyproj.Proj(projmet)(declon,declat)
                x=mycoords[0]
                y=mycoords[1]
                '''

                
                alt=str(geo['GPS GPSAltitude'])
                altlist=alt.split('/')
                alt=int(altlist[0])/int(altlist[1])

                row=(file, declat, declon, alt)
                datalist.append(row)
        
        if len(datalist)>0:
            wb=xl.Workbook()
            sheet=wb.active
            for i in datalist:
                sheet.append(i)
            try:
                wb.save('Image_coordinates.xlsx')
                messagebox.showinfo('Infobox','Process complete')
                window.destroy()
            except PermissionError:
                messagebox.showerror('Error', 'Please close excel file and try again')
        else:
            messagebox.showerror('Error', 'Folder contains no images or images have no cooridnates')        
    else:
        messagebox.showerror('Error', 'Please choose a valid folder')
    


window=tkinter.Tk()
window.geometry('600x360')
window.resizable(0,0)
window.title('IMG2XL')
window.iconbitmap('uav.ico')
window.configure(bg='yellow')
   
bglabel=tkinter.Label(window)
bglabel.place(x=0, y=0)

choosebutton=tkinter.Button(window, text='Choose folder', font='Bold 12', padx=30, pady=20, command=lambda:choosefolder())
choosebutton.place(x=10,y=250)

folderlabel=tkinter.Label(window, text=choosefolder)
folderlabel.place(x=10, y=325)

runbutton=tkinter.Button(window, text='Run', font='Bold 12', padx=30, pady=20, command=mainfunc)
runbutton.place(x=200, y=250)

window.mainloop()


