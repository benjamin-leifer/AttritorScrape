from lxml import html
import requests
import datetime
import time
import re
import xlwt
import pandas as pd
from pandas import DataFrame
import tkinter.tix as tk
import threading
import matplotlib.pyplot as plt
import io
from base64 import encodebytes
from tkinter.filedialog import askopenfilenames
from scipy import integrate

#matplotlib.style.use('ggplot')

global isRunning
isRunning=False
global Hz
Hz=[]
global HP
HP=[]
global RPM
RPM=[]
global Time
Time=[]

import smtplib, os
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate
from email import encoders

def send_mail( send_from, send_to, subject, text, files=[], server="localhost", port=587, username='', password='', isTls=True):
    print('send to is: '+ send_to)
    msg = MIMEMultipart()
    msg['From'] = send_from
    msg['To'] = COMMASPACE.join(send_to)
    msg['Date'] = formatdate(localtime = True)
    msg['Subject'] = subject

    msg.attach( MIMEText(text) )

    for f in files:
        part = MIMEBase('application', "octet-stream")
        part.set_payload( open(f,"rb").read() )
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment; filename="{0}"'.format(os.path.basename(f)))
        msg.attach(part)
    """
    for f in files:
        fp = open(f, 'r',errors='ignore')
        part = MIMEBase('application', "octet-stream")
        part.set_payload(encodebytes(fp.read()).decode())
        fp.close()
        part.add_header('Content-Transfer-Encoding', 'base64')
        part.add_header('Content-Disposition', 'attachment; filename="%s"' % file)
        msg.attach(part)
    """

    smtp = smtplib.SMTP('smtp.gmail.com:587')
    if isTls: smtp.starttls()
    smtp.login(username,password)
    smtp.sendmail(send_from, send_to, msg.as_string())
    smtp.quit()


class AttritorScrape(tk.Frame):

    def __init__(self, master):
        tk.Frame.__init__(self,master)
        self.master = master
        self.isRunning = False
        self.Hz = []
        self.HP = []
        self.RPM = []
        self.Time = []
        self.Notes = []

        #master.title("Attritor Scrape")

        #print(Hz)

        pageLabel = tk.Label(master, text='HTML page')
        pageLabel.pack()
        self.pageVar = tk.StringVar()
        #pageVar.set('Feature Currently not Working')
        #pageEntry = tk.Entry(master,textvariable=pageVar, width = 50)
        #pageEntry.pack()
        """
        pages = [('VHM1','http://10.10.92.175/process_display_0.htm'),
                 ('VHM2','http://10.10.92.63/process_display_0.html'),
                 ('VS1','http://10.10.92.51/process_display_0.html'),
                 ('VS2','http://10.10.92.106/process_display_0.html'),
                ]
        pageVar.set('http://10.10.92.175/process_display_0.html')
        """
        pages = [('VHM1','VHM1'),
                 ('VHM2','VHM2'),
                 ('VS1','VS1'),
                 ('VS2','VS2'),
                ]
        self.pageVar.set('VHM1')
        for text, page in pages:
            b = tk.Radiobutton(master, text = text, variable = self.pageVar, value = page)
            b.pack()

        filenameLabel = tk.Label(master, text='Filename')
        filenameLabel.pack()
        self.fileVar = tk.StringVar()
        fileEntry = tk.Entry(master,textvariable=self.fileVar, width = 50)
        fileEntry.pack()

        sheetLabel = tk.Label(master, text='Sheet Name')
        sheetLabel.pack()
        sheetVar = tk.StringVar()
        sheetEntry = tk.Entry(master,textvariable=sheetVar, width = 50)
        sheetEntry.pack()

        emailLabel = tk.Label(master, text='Email')
        emailLabel.pack()
        self.emailVar = tk.StringVar()
        emailEntry = tk.Entry(master,textvariable=self.emailVar, width = 50)
        emailEntry.pack()

        runtimeLabel = tk.Label(master, text='Run Time')
        runtimeLabel.pack()
        self.runtimestartVar = tk.StringVar()
        self.runtimecurrentVar = tk.StringVar()
        runtimestartEntry = tk.Entry(master,textvariable=self.runtimestartVar)
        runtimestartEntry.pack()
        runtimecurrentEntry = tk.Entry(master,textvariable=self.runtimecurrentVar)
        runtimecurrentEntry.pack()
        self.NotesVar = tk.StringVar()
        notesLabel = tk.Label(master, text = 'Notes')
        notesLabel.pack()
        notesEntry = tk.Entry(master, textvariable = self.NotesVar)
        notesEntry.pack()

        recordbutton = tk.Button(master,text = 'Start Recording Data', command = lambda: self.runThread(self.pageVar.get()))
        recordbutton.pack()

        erasebutton = tk.Button(master, text = 'Erase Stored Data', command = lambda: self.eraseStoredData())
        erasebutton.pack()

        writeButton = tk.Button(master, text = 'Write to Excel', command = lambda: self.writeToExcel(self.fileVar.get(),sheetVar.get()))
        writeButton.pack()

        plotButton = tk.Button(master, text = 'Plot HP vs Time', command = lambda: self.plotData())
        plotButton.pack()

        stoprecordbutton = tk.Button(master,text = 'Stop Recording Data', command = lambda: self.stopThread())
        stoprecordbutton.pack()

        resetCountbutton = tk.Button(master, text = 'Reset Count', command = lambda: self.resetCount())
        resetCountbutton.pack()

        subject = 'Attritor Data'
        text = 'Test'
        #emailgetter = lambda y: ['bleifer@gmail.com'] if (emailVar.get() is not '') else [emailVar.get()]
        #emailgetter = self.getEmail()
        emailButton = tk.Button(master, text = 'Send Email', command = lambda: send_mail('veloxintattritor@gmail.com',self.getEmail(),subject, text, files = [self.writeToExcel(self.fileVar.get(),sheetVar.get())], username = 'veloxintattritor@gmail.com',password = 'tnix0lev!!'))
        emailButton.pack()

        importFileButton = tk.Button(master, text = 'Import Files', command = lambda: self.importFiles())
        importFileButton.pack()

        plotAllDataButton = tk.Button(master, text = 'Plot All Data', command = lambda: self.plotData2())
        plotAllDataButton.pack()

        RunDataButton = tk.Button(master, text = 'RunData', command = lambda: self.updateRunData())
        RunDataButton.pack()

        #Print2Button = tk.Button(master, text = 'Print2', command = lambda: self.writeToExcel2(self.whileCount-30,-1,self.pageVar.get(),self.fileVar.get()))
        #Print2Button.pack()

        NotifyMeButton = tk.Button(master, text = 'Change Notification', command = lambda: self.changeNotifyToFalse())
        NotifyMeButton.pack()

        self.HP_Cut_Off= tk.StringVar()
        self.HP_Cut_Off.set('5.0')
        HP_Cut_OffLabel = tk.Label(master, text = 'Cut Off HP')
        HP_Cut_OffLabel.pack()
        HP_Cut_OffEntry = tk.Entry(master, textvariable= self.HP_Cut_Off)
        HP_Cut_OffEntry.pack()

        self.TotalRunTime = tk.StringVar()
        totalRunTimeLabel = tk.Label(master, text = 'Total Run Time')
        totalRunTimeLabel.pack()
        totalRunTimeEntry = tk.Entry(master, textvariable= self.TotalRunTime)
        totalRunTimeEntry.pack()

        self.PowderWeight = tk.StringVar()
        powderWeightLabel = tk.Label(master, text = 'Powder Weight (kg)')
        powderWeightLabel.pack()
        powderWeightEntry = tk.Entry(master, textvariable= self.PowderWeight)
        powderWeightEntry.pack()

        self.CumulativePower = tk.StringVar()
        CumulativePowerLabel = tk.Label(master, text = 'Total Power in Joules')
        CumulativePowerLabel.pack()
        CumulativePowerEntry = tk.Entry(master, textvariable= self.CumulativePower)
        CumulativePowerEntry.pack()

        self.powerStartTime=tk.StringVar()
        self.powerEndTime=tk.StringVar()



    def recordRunData(self, pageInit):
        #self.recording = True
        print('page is: ')
        print (pageInit)
        isRunning = True
        print('isRunning is: '+str(self.isRunning))
        timestart = time.time()
        count = 0
        non_decimal = re.compile(r'[^\d.]+')
        self.whileCount=0
        self.notified=False
        while self.isRunning:
        #while time.time()<timestart+3:
            self.whileCount = self.whileCount +1
            #print('in loop')
            #print(pageInit)
            if pageInit == 'VHM1':
                #print('entered correct if place')
                page = requests.get('http://10.10.92.175/process_display_0.html')
            elif pageInit == 'VHM2':
                page = requests.get('http://10.10.92.63/process_display_0.html')
            elif pageInit == 'VS1':
                page = requests.get('http://10.10.92.51/process_display_0.html')
            elif pageInit == 'VS2':
                page = requests.get('http://10.10.92.106/process_display_0.html')
            #page = requests.get('http://10.10.92.63/process_display_0.html')
            #page = requests.get(page)
            tree = html.fromstring(page.content)

            Tree = tree.xpath('//td/text()')

            for element in Tree:
                if count ==15:
                    self.Hz.append(float(non_decimal.sub('',element)))
                if count == 17:
                    self.RPM.append(float(non_decimal.sub('',element)))
                if count == 19:
                    self.HP.append(float(non_decimal.sub('',element)))
                count=count+1
            self.Time.append(datetime.datetime.now())
            self.Notes.append(self.NotesVar.get())
            #print(time.time())
            time.sleep(.5)
            count = 0

            if self.whileCount%10 ==0:
                print (self.Time[-1],self.Hz[-1],self.RPM[-1],self.HP[-1],self.whileCount)
                self.currentStatus()
            """
            if self.whileCount==5000:
                #self.stopThread()
                filename = str(self.Time[0])+'-'+str(self.Time[-1])
                filename = re.sub(r':',' ',filename)
                filename = self.pageVar.get() + '-- ' + filename
                sheetname = filename
                self.writeToExcel(filename, sheetname)
                self.eraseStoredData()
                self.currentStatus()
                self.whileCount=0
            """
            if self.whileCount%7200==0:
                self.writeToExcel2(self.whileCount-7200,-1,self.pageVar.get(),self.fileVar.get())
                print('Saved File')
            if len(self.HP)>35:
                try:
                    if self.notified==False and (sum(self.HP[-30:])/len(self.HP[-30:]))>float(self.HP_Cut_Off.get()):
                        subject = 'High HP Alert'
                        text = 'HP_High: '+str(self.HP_Cut_Off.get())
                        send_mail('veloxintattritor@gmail.com',self.getEmail(),subject, text, files = [], username = 'veloxintattritor@gmail.com',password = 'tnix0lev!!')
                        self.notified=True
                except:
                    print('Error as expected')

        #print(Time,Hz,RPM,HP)

    def writeToExcel(self, filename, sheetname):

        if filename == '':
            filename = str(self.Time[0])+'-'+str(self.Time[-1])
            filename = re.sub(r':',' ',filename)
            sheetname = re.sub(r':', ' ', str(self.Time[-1]))

        df = DataFrame({'HP':self.HP,'RPM':self.RPM,'Hz':self.Hz,'Time':self.Time,'Notes':self.Notes})
        filename = filename+'.xlsx'
        #output = io.BytesIO()
        writer = pd.ExcelWriter(filename, engine = 'xlsxwriter')
        df.to_excel(writer, sheet_name='Sheet1')
        #df.to_excel(writer, sheet_name = 'Sheet1')
        #df.to_excel(filename, sheet_name = 'sheet1', index=False)
        writer.save()
        #xlsx_data = output.getvalue()
        return filename

    def writeToExcel2(self, startIndex, endIndex, millName, filename):
        #if filename == '':
        print ('Start index is: '+ str(startIndex)+': '+ str(self.Time[startIndex]))
        print ('End index is: '+ str(endIndex)+': '+ str(self.Time[endIndex]))
        dateCompName = str(self.Time[startIndex])+'-'+str(self.Time[endIndex])
        dateCompName = re.sub(r':',' ',dateCompName)
        preFileName = millName + '--' + dateCompName + ' '
        print('preFileName is: '+ preFileName)
        #print ('MillName is: ' +str(millName))
        #sheetname = re.sub(r':', ' ', str(self.Time[-1]))

        df = DataFrame({'HP':self.HP[startIndex:endIndex],'RPM':self.RPM[startIndex:endIndex],'Hz':self.Hz[startIndex:endIndex],'Time':self.Time[startIndex:endIndex],'Notes':self.Notes[startIndex:endIndex]})
        filename = preFileName+filename+'.xlsx'
        print('filename is: '+filename)
        #output = io.BytesIO()
        writer = pd.ExcelWriter(filename, engine = 'xlsxwriter')
        df.to_excel(writer, sheet_name='Sheet1')
        #df.to_excel(writer, sheet_name = 'Sheet1')
        #df.to_excel(filename, sheet_name = 'sheet1', index=False)
        writer.save()
        #xlsx_data = output.getvalue()
        return filename

    def changeNotifyToFalse(self):
        self.notified = False
        print('Notification Status Set to: '+ str(self.notified))

    def eraseStoredData(self):
        self.Hz = []
        self.RPM = []
        self.HP = []
        self.Time = []

        self.currentStatus()

    def currentStatus(self):
        if self.Time:
            self.runtimestartVar.set(self.Time[0])
            self.runtimecurrentVar.set(self.Time[-1])
            #print(self.Time[0],self.Time[-1])
            #print(self.runtimestartVar,self.runtimecurrentVar)
        else:
            self.runtimestartVar.set('')
            self.runtimecurrentVar.set('')

    def runThread(self, page):

        isRunning = True
        print(isRunning)

        Thread = scrapeThread(self, 1, 'thread', 1, page)
        Thread.start()

    def stopThread(self):
        self.isRunning = False
        print('isRunning is: '+str(self.isRunning))

    def plotData(self):

        HP = self.HP
        Time = self.Time[0:len(HP)]

        plt.plot(Time, HP)
        x1,x2,y1,y2 = plt.axis()
        plt.axis([x1,x2,0,6])
        plt.show()

    def plotData2(self):
        print('printing All Data')
        df = DataFrame({'HP':self.HP,'RPM':self.RPM,'Hz':self.Hz,'Time':self.Time,'Notes':self.Notes})
        self.DataFrames = pd.concat([self.DataFrames,df]).drop_duplicates()
        self.DataFrames.plot(x='Time',y ='HP')
        plt.show()
        #self.DataFrames.plot()
    def resetCount(self):
        self.writeToExcel(self.fileVar.get(),sheetVar.get())
        self.whileCount = 0
        print('whileCount reset')

    def getEmail(self):
        #emailgetter = lambda y: ['bleifer@gmail.com'] if (emailVar.get() is not '') else [emailVar.get()]
        if self.emailVar.get() is '':
            email = 'bleifer@veloxint.com'
        else:
            email = self.emailVar.get()
        return email

    def importFiles(self):
        files = askopenfilenames(filetypes=(('Excel files', '*.xlsx'),
                                   ('All files', '*.*')),
                                   title='Select Input File'
                                   )
        fileList = root.tk.splitlist(files)
        print('Files = ', fileList)
        self.combineExcels(fileList)

    def combineExcels(self, filesList):
        self.files={}
        self.filesList = []
        for c,file in enumerate(filesList):
            self.files[file] = pd.read_excel(file)
            self.filesList.append(pd.read_excel(file))
        self.DataFrames = pd.concat(self.filesList)
        self.DataFrames = self.DataFrames.drop_duplicates()
        print('self.DataFrames is: ')
        #with pd.option_context('display.max_rows', None, 'display.max_columns', 3):
            #print(self.DataFrames)
        self.DataFrames['Time'] = pd.to_datetime(self.DataFrames['Time'])
        #self.DataFrames.sort_values('Time')
        self.DataFrames.index =self.DataFrames['Time']
        #del self.DataFrames['Time']
        self.DataFrames= self.DataFrames.sort_index()
        #with pd.option_context('display.max_rows', None, 'display.max_columns', 3):
            #print(self.DataFrames)

        #print (self.DataFrames.columns.tolist())
    def updateRunData(self):

        def onclick(event):
            self.dblclickCount=0
            print('%s click: button=%d, x=%d, y=%d, xdata=%f, ydata=%f' %
                ('double' if event.dblclick else 'single', event.button,
                event.x, event.y, event.xdata, event.ydata))
            if event.dblclick and self.dblclickCount==0:
                self.beginningTime=event.xdata
                self.dblclickCount=1
                print('beginningTime: '+ str(event.xdata))
            elif event.dblclick and self.dblclick==1:
                self.endTime = event.xdata
                print('endTime: '+ str(event.xdata))

        df = DataFrame({'HP':self.HP,'RPM':self.RPM,'Hz':self.Hz,'Time':self.Time,'Notes':self.Notes})
        self.DataFrames = pd.concat([self.DataFrames,df]).drop_duplicates(subset='Time')

        #self.DataFrames['Time'] = pd.to_datetime(self.DataFrames['Time'])
        self.DataFrames.index =self.DataFrames['Time']
        #del self.DataFrames['Time']
        self.DataFrames= self.DataFrames.sort_index()

        self.DataFrames['TimeDelta']= self.DataFrames.Time - self.DataFrames.Time.shift()
        print(self.DataFrames)
        runMax = self.DataFrames['Hz'].max(axis=0)
        print('runMax is: ')
        print(runMax)
        print(self.DataFrames['Time'].values)
        print(self.DataFrames.columns.tolist())
        #self.DataFrames.drop(self.DataFrames[self.DataFrames.Hz<runMax-25].index)
        self.DataFrames_noZero = self.DataFrames.drop(self.DataFrames[self.DataFrames.Hz<runMax-25].index)
        with pd.option_context('display.max_rows', None, 'display.max_columns', 3):
            print(self.DataFrames_noZero)
            #print(self.DataFrames)
        fig, ax = plt.subplots()
        ax.plot(self.DataFrames['Time'].values,self.DataFrames['HP'].values)
        #self.DataFrames.plot(x='Time',y ='HP')

        #cid = fig.canvas.mpl_connect('button_press_event', onclick)
        plt.show()
        #print(self.dblclickCount)
        self.dblclickCount=0
        print('TimeDelta Max then Min is: ')
        print(self.DataFrames_noZero['TimeDelta'].max(axis=0))
        print(self.DataFrames_noZero['TimeDelta'].min(axis=0))
        self.TotalRunTime.set(self.DataFrames_noZero['TimeDelta'].sum())
        self.CumulativePower.set(((745.7*self.DataFrames_noZero['HP']*self.DataFrames_noZero['TimeDelta']).sum().total_seconds())/(1000000*float(self.PowderWeight.get())))
    #def processRunData(self):



class scrapeThread (threading.Thread):

    def __init__(self, tkParent, threadID, name, counter, page):
        threading.Thread.__init__(self)
        self.tkParent = tkParent
        print(self.tkParent)
        self.threadID = threadID
        self.name = name
        self.counter = counter
        self.page = page

    def run(self):
        print("Starting "+self.name)
        self.tkParent.isRunning = True
        AttritorScrape.recordRunData(self.tkParent, self.page)
        print('Exiting '+self.name)

root = tk.Tk()
scroll = tk.ScrolledWindow(root, scrollbar=tk.BOTH)
scroll.pack(fill=tk.BOTH, expand=1)
AttritorScrape(scroll.window).pack()
root.mainloop()
