from PyQt5.QtWidgets import QWidget, QApplication, QMainWindow, QFileDialog, QColorDialog, QDialog, QMessageBox
from PyQt5 import uic
from PyQt5.QtCore import QCoreApplication, Qt, QUrl
from PyQt5.QtGui import QColor, QPixmap, QImage, QDesktopServices, QIcon
import sys, os, openpyxl, qrgen, time, qrcode, io
from shutil import rmtree
import pickle
from cryptography.fernet import Fernet


class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow,self).__init__()
        self.x=['','','']
        self.validities=[False,False]
        self.scond=False
        self.header=1
        self.prefixes=[]
        self.suffixes=[]
        self.delims=[]
        self.worksheet=[]
        self.indx=0
        self.clmindex=[]
        self.fixedkey=b'sJeXPvBz2WeQ8-WRreLeg1ntaASes6BscwJvMmSk_QU='
        self.loadkey=crypt().loadkey()

        self.finalstring=''

        # Load the UI file
        self.ui_file = "src/ui/mainwindow.ui"
        uic.loadUi(self.ui_file, self)
        self.stacks.setCurrentWidget(self.setup)
        self.save=self.setup
        self.Browse1.clicked.connect(lambda: self.browsefiles(self.lineEdit,0,'Database File','Excel File (*.xlsx)'))
        self.checkBox.stateChanged.connect(self.toggleStatus)
        self.Browse2.clicked.connect(lambda: self.browsefiles(self.lineEdit_2,1,'Logo File','Logo Image (*.png *.jpg *.jpeg *.bmp *.gif)'))
        self.lineEdit_3.setText(str(self.hslider.value()))
        self.hslider.valueChanged.connect(self.ValueUpdate)
        self.lineEdit_3.textChanged.connect(self.sliderUpdate)
        self.colorPicker.clicked.connect(self.pick_color)
        self.colorline.textChanged.connect(self.colorupdate)
        self.colorvalidity.setVisible(0)
        self.logovalidity.setVisible(0)
        self.datavalidity.setVisible(0)
        self.filevalidity.setVisible(0)
        self.lineEdit.textChanged.connect(self.strupdate1)
        self.lineEdit_2.textChanged.connect(self.strupdate2)
        self.helpbtn.clicked.connect(self.helpdialog)
        self.nextbutton.clicked.connect(self.next1)
        self.aboutbtn.clicked.connect(lambda: self.stacks.setCurrentWidget(self.aboutpage))
        self.simpleqr.clicked.connect(self.tosimpleqr)


     #------------------------------operation page-----------------------------------------------------------------------------------
        self.bckbtn1.clicked.connect(lambda: self.stacks.setCurrentWidget(self.setup))
        self.headercheck.stateChanged.connect(self.toggleHeader)
        self.loadbtn.clicked.connect(self.loadxl)
        self.deliName=['None','Comma','Tab','Colon','Semicolon','Newline','Forward Slash']
        self.deliValu=['',',','\t',':',';','\n','/']
        for name, value in zip(self.deliName, self.deliValu):
             self.delibox.addItem(name,value)
        self.addbtn.clicked.connect(self.addData)
        self.prefixline.textChanged.connect(self.prefixupdate)
        self.suffixline.textChanged.connect(self.suffixupdate)
        self.takebtn.clicked.connect(self.take)
        self.exprtbtn.clicked.connect(self.export)
        self.advncbtn.clicked.connect(lambda: self.stacks.setCurrentWidget(self.advancepage))
        self.samplebtn.clicked.connect(self.check)
        self.folderbtn.clicked.connect(lambda: self.browsefolder(self.folderline))
        self.clrbtn.clicked.connect(self.clearbutton)
        

     #------------------------------about page-----------------------------------------------------------------------------------
        self.okabout.clicked.connect(lambda: self.stacks.setCurrentWidget(self.save))

     #------------------------------advanced page--------------------------------------------------------------------------------
        self.errorlevels=[qrcode.constants.ERROR_CORRECT_L, qrcode.constants.ERROR_CORRECT_M, qrcode.constants.ERROR_CORRECT_Q, qrcode.constants.ERROR_CORRECT_H]
        self.errornames=['Low (7%)','Medium (15%)', 'Quartile (25%)', 'High (30%)']
        for name,level in zip(self.errornames,self.errorlevels):
          self.errorlvlbox.addItem(name,level)
        self.errorlvlbox.setCurrentIndex(1)
        self.resetbtn.clicked.connect(self.reset)
        self.savebtn.clicked.connect(self.savebutton)
        self.checkbtn.clicked.connect(self.checkbutton)

     #------------------------------Simple QR page--------------------------------------------------------------------------------
        self.entertext.textChanged.connect(self.simpleqrtext)
        self.folderbtnsimp.clicked.connect(lambda: self.browsefolder(self.folderlinesimp))
        self.folderlinesimp.textChanged.connect(lambda: self.exprtbtnsimp.setEnabled(True))
        self.exprtbtnsimp.clicked.connect(self.export2)
        self.samplebtnsimp.clicked.connect(self.checksimple)
        self.advncbtnsimp.clicked.connect(lambda: self.stacks.setCurrentWidget(self.advancepage))
        self.batchqrbtn.clicked.connect(lambda: self.stacks.setCurrentWidget(self.setup))
        self.clrbtnsimp.clicked.connect(lambda: self.entertext.setPlainText(''))
        self.aboutbtnsimp.clicked.connect(lambda: self.stacks.setCurrentWidget(self.aboutpage))
        
    def checkbutton(self):
         if self.save==self.setup:
              self.check()
         else:
              self.checksimple()

    def savebutton(self):
         if self.save==self.setup:
              self.stacks.setCurrentWidget(self.operation)
         else:
              self.stacks.setCurrentWidget(self.save)

    def simpleqrtext(self):
         if self.entertext.toPlainText()=='':
              x=False
         else:
              x=True
         self.advncbtnsimp.setEnabled(x)
         self.clrbtnsimp.setEnabled(x)
         self.samplebtnsimp.setEnabled(x)
         self.exportlabelsimp.setEnabled(x)
         self.folderlinesimp.setEnabled(x)
         self.folderbtnsimp.setEnabled(x)
         self.filelabelsimp.setEnabled(x)
         self.filesimp.setEnabled(x)
         self.png.setEnabled(x)

    def checksimple(self):
         img=qrgen.simpleqr(self.entertext.toPlainText(),self.colorline.text(),self.version(),self.errorlvlbox.currentData(),self.sizespinbox.value(),self.borderspinbox.value())
         buffer = io.BytesIO()
         img.save(buffer, format='png')
         qimage = QImage.fromData(buffer.getvalue())        
         self.check_clicked(QPixmap.fromImage(qimage))

    def reset(self):
         self.versionspinbox.setValue(0)
         self.errorlvlbox.setCurrentIndex(1)
         self.sizespinbox.setValue(10)
         self.borderspinbox.setValue(4)
    def version(self):
         if self.versionspinbox.value()>0:
              return self.versionspinbox.value()
         else: 
              return None

    def tosimpleqr(self):
         self.stacks.setCurrentWidget(self.simpleqrpage)
         self.outfilebox.setVisible(0)
         self.outlabel.setVisible(0)
         self.save=self.simpleqrpage

    def tobatchqr(self):
         self.stacks.setCurrentWidget(self.setup)
         self.outfilebox.setVisible(1)
         self.outlabel.setVisible(1)
         self.save=self.operation

    def check(self):
         if self.checkBox.isChecked():
              img=qrgen.qrwlogo(self.lineEdit_2.text(),self.lineEdit_3.text(),self.outputtxt.toPlainText(),self.colorline.text(),self.version(),self.errorlvlbox.currentData(),self.sizespinbox.value(),self.borderspinbox.value())
         else:
              img=qrgen.simpleqr(self.outputtxt.toPlainText(),self.colorline.text(),self.version(),self.errorlvlbox.currentData(),self.sizespinbox.value(),self.borderspinbox.value())
         buffer = io.BytesIO()
         img.save(buffer, format='png')
         qimage = QImage.fromData(buffer.getvalue())        
         self.check_clicked(QPixmap.fromImage(qimage))

    def check_clicked(self,x):
        # Load the second window UI file
        self.second_window = uic.loadUi("src/ui/sampleui.ui")
        self.second_window.qrdisplay.setPixmap(x)
        self.second_window.close()
        self.second_window.setWindowFlags(self.second_window.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        # Show the second window
        self.second_window.show()
        self.second_window.editbtn.clicked.connect(self.edit)
        self.second_window.looksgoodbtn.clicked.connect(self.looksgood)
        self.second_window.exec_()

    def browsefolder(self,text):
        cryptpath=os.path.join(os.getcwd(),r'bin\redlof.bin')
        try:
             path=crypt().dcrypt(cryptpath, self.loadkey)  
        except:
             path=''
        folder=QFileDialog.getExistingDirectory(self, 'Select Folder', path)
        if not folder=='':
          text.setText(folder)
          crypt().ncrypt(folder,cryptpath,self.loadkey)
          self.exprtbtn.setEnabled(1)

    def clearbutton(self):
         self.outputtxt.setPlainText('')
         self.clrbtn.setEnabled(0)
         self.prefixes=[]
         self.suffixes=[]
         self.delims=[]
         self.clmindex=[]
         self.finalstring=''
         self.indx=0
         self.prefixline.clear()
         self.suffixline.clear()
         self.exprtbtn.setEnabled(0)
         self.advncbtn.setEnabled(0)
         self.samplebtn.setEnabled(0)
         self.folderbtn.setEnabled(0)
         self.folderline.setEnabled(0)
         self.exportlabel.setEnabled(0)
               
        
    def edit(self):
         self.second_window.close()
         self.stacks.setCurrentWidget(self.advancepage)

    def looksgood(self):
         self.second_window.close()
         if self.save==self.setup:
              self.stacks.setCurrentWidget(self.operation)
         else:
              self.stacks.setCurrentWidget(self.save)


    def next1(self):
         if (os.path.exists(self.x[0]) and os.path.exists(self.x[1])) or (os.path.exists(self.x[0]) and not self.checkBox.isChecked()):
              self.filevalidity.setVisible(0)
              self.stacks.setCurrentWidget(self.operation)  
              self.sheetbox.clear()            
              self.sheetbox.addItems(xl.sheets(self.x[0]))
              
         else:
              self.filevalidity.setVisible(1)
              
    def toggleHeader(self,state):
         if state==2:
              self.header=2
         else:
              self.header=1
    def addData(self):
         self.prefixlabel.setEnabled(1)
         self.prefixline.setEnabled(1)
         self.suffixlabel.setEnabled(1)
         self.suffixline.setEnabled(1)
         self.datalabel.setEnabled(1)
         self.databox2.setEnabled(1)
         self.samvlabel.setEnabled(1)
         self.samview.setEnabled(1)
         self.takebtn.setEnabled(1)
         self.databox2.setText(self.databox.currentText())
         self.samview.setText(self.delibox.currentData()+self.prefixline.text()+self.databox2.text()+self.suffixline.text())

    def prefixupdate(self,value):
         text=self.delibox.currentData()+value+self.databox2.text()+self.suffixline.text()
         self.samview.setText(text)

    def suffixupdate(self,value):
         text=self.delibox.currentData()+self.prefixline.text()+self.databox2.text()+value
         self.samview.setText(text)

    def take(self):
         self.prefixes.append(self.prefixline.text())
         self.suffixes.append(self.suffixline.text())
         self.delims.append(self.delibox.currentData())
         self.clmindex.append(self.databox.currentData())
         self.finalstring=self.finalstring+self.stringmkr(self.header)
         self.outputtxt.setPlainText(self.finalstring)
         self.indx+=1
         self.prefixline.clear()
         self.suffixline.clear()
         self.advncbtn.setEnabled(1)
         self.samplebtn.setEnabled(1)
         self.folderbtn.setEnabled(1)
         self.folderline.setEnabled(1)
         self.exportlabel.setEnabled(1)
         self.clrbtn.setEnabled(1)
         if not self.folderline.text()=='':
              self.exprtbtn.setEnabled(1)

    def stringmkr(self,row):
         text=f'{self.delims[self.indx]}{self.prefixes[self.indx]}{self.worksheet.cell(row=row,column=self.clmindex[self.indx]).value}{self.suffixes[self.indx]}'
         return text
                       
    def export2(self):
         img=qrgen.simpleqr(self.entertext.toPlainText(),self.colorline.text(),self.version(),self.errorlvlbox.currentData(),self.sizespinbox.value(),self.borderspinbox.value())
         path=self.folderlinesimp.text()
         if self.filesimp.text()=='':
               fname='QR Code'
         else:
              fname=self.filesimp.text()
         file=os.path.join(path, f'{fname}.png') 
         img.save(file)
         from PyQt5.QtWidgets import QMessageBox

         msg = QMessageBox()
         msg.setIcon(QMessageBox.Information)
         msg.setText("Done!")
         msg.setWindowTitle("Message")
         msg.setStandardButtons(QMessageBox.Ok)
         msg.resize(500, 200)
         msg.exec_()
       

    def export(self):
         self.prowindow=uic.loadUi('src/ui/progress.ui')
         rownum=self.worksheet.max_row
         ndx=1
         if self.headercheck.isChecked():
              rownum-=1
              ndx+=1
         self.prowindow.show()
         self.prowindow.progressBar.setValue(0)
         path=self.folderline.text()
         os.makedirs(os.path.dirname(path), exist_ok=True)
         
         if self.checkBox.isChecked():
              for row in range(rownum):
                    text=''                              
                    for i in range(self.indx):
                         text+=f'{self.delims[i]}{self.prefixes[i]}{self.worksheet.cell(row=row+ndx,column=self.clmindex[i]).value}{self.suffixes[i]}'
                    
                    self.prowindow.progressBar.setValue(int(((row+ndx)/rownum)*100))
                    time.sleep(0.01)
                    QCoreApplication.processEvents()                    
                    img=qrgen.qrwlogo(self.lineEdit_2.text(),self.lineEdit_3.text(),text,self.colorline.text(),self.version(),self.errorlvlbox.currentData(),self.sizespinbox.value(),self.borderspinbox.value())
                    img.save(self.outfilename(row+ndx,path))
         else:
              for row in range(rownum):
                    text=''
                    for i in range(self.indx):
                         text+=f'{self.delims[i]}{self.prefixes[i]}{self.worksheet.cell(row=row+ndx,column=self.clmindex[i]).value}{self.suffixes[i]}'
                         
                    self.prowindow.progressBar.setValue(int(((row+ndx)/rownum)*100))
                    time.sleep(0.01)
                    QCoreApplication.processEvents()                    
                    img=qrgen.simpleqr(text,self.colorline.text(),self.version(),self.errorlvlbox.currentData(),self.sizespinbox.value(),self.borderspinbox.value())
                    img.save(self.outfilename(row+ndx,path))
                         
        
         self.prowindow.okbtn.clicked.connect(self.prowindow.close)
         self.prowindow.folderbtn.clicked.connect(lambda: self.folder(self.folderline.text()))
         self.prowindow.okbtn.setEnabled(1)
         self.prowindow.folderbtn.setEnabled(1)
         self.prowindow.exec_()
         self.delims=[]
         self.prefixes=[]
         self.suffixes=[]
         self.clminde=[]
         self.indx=0
         self.prowindow.okbtn.setEnabled(0)
         self.prowindow.folderbtn.setEnabled(0)
         self.finalstring=''
         
                 
    def folder(self,x):
         self.prowindow.close()
         QDesktopServices.openUrl(QUrl.fromLocalFile(x))
                      

    def loadxl(self):
         self.worksheet, headers =xl.data(self.x[0],self.sheetbox.currentText(),self.header)
         self.databox.clear()
         val=1
         for head in headers:
              self.databox.addItem(head,val)
              self.outfilebox.addItem(head,val)
              val+=1


         self.label_10.setEnabled(1)
         self.databox.setEnabled(1)   
         self.delilabel.setEnabled(1)
         self.delibox.setEnabled(1)
         self.addbtn.setEnabled(1)           
              
    def outfilename(self, row, path):
         col = self.outfilebox.currentData()
         fname = self.worksheet.cell(row=row, column=col).value
         full = os.path.join(path, f'{fname}.png')
         i = 1
         while os.path.exists(full):
              full = os.path.join(path, f'{fname}_{i}.png')
              i += 1
         return full

              
         

    def browsefiles(self,text,i,dialg,typ):
            fname=QFileDialog.getOpenFileName(self,'Open file',dialg,typ)
            text.setText(fname[0])
            self.x[i]=fname[0]
            self.nextenable()
            

    def toggleStatus(self,state):         
         if state==2:
              self.label_4.setEnabled(1)
              self.lineEdit_2.setEnabled(1)
              self.Browse2.setEnabled(1)
              self.hslider.setEnabled(1)
              self.lineEdit_3.setEnabled(1)
              self.label_5.setEnabled(1)
              if not self.validities[1] and self.scond:
                   self.logovalidity.setVisible(1)
              else: 
                   self.logovalidity.setVisible(0)
              

         else:
              self.label_4.setEnabled(0)
              self.lineEdit_2.setEnabled(0)
              self.Browse2.setEnabled(0)
              self.hslider.setEnabled(0)
              self.lineEdit_3.setEnabled(0)
              self.label_5.setEnabled(0)
              self.logovalidity.setVisible(0)             
         self.nextenable()
    
    def ValueUpdate(self,value):
         self.lineEdit_3.setText(str(value))
         self.x[2]=value
    
    def sliderUpdate(self,value):
         if value.isdigit():
              value=int(value)
              if value>0 and value<101:
                   self.hslider.setValue(value)
                   self.x[2]=value
              
    def pick_color(self):
        color = QColorDialog.getColor()
        if color.isValid():
            # Do something with the selected color
            self.colorline.setText(color.name())
            self.colordisplay.setStyleSheet(f"background-color: {color.name()};")
            self.colorvalidity.setVisible(0)
    
    def colorupdate(self,cl):
         color=QColor(cl)

         if color.isValid():
              self.colorline.setText(color.name())
              self.colordisplay.setStyleSheet(f"background-color: {color.name()};")
              self.colorvalidity.setVisible(0)
              
         else:
              self.colorvalidity.setVisible(1)
    
    def strupdate1(self,value):
         self.x[0]=value
         if value.endswith('.xlsx') or value=='':
              self.validities[0]=True
              self.datavalidity.setVisible(0)
         elif value=='':
              self.validities[0]=False
              self.datavalidity.setVisible(0)
         else:
              self.validities[0]=False
              self.datavalidity.setVisible(1)
         
         self.nextenable()
         
    def strupdate2(self,value):
         self.x[1]=value
         if self.x[1].endswith('.png') or self.x[1].endswith('.jpg') or self.x[1].endswith('.jpeg') or self.x[1].endswith('.bmp') or self.x[1].endswith('.gif'):
              self.validities[1]=True
              self.logovalidity.setVisible(0)
         else:
              self.validities[1]=False
              self.logovalidity.setVisible(1)
         
         self.nextenable()
         self.scond=True
         
    
    def nextenable(self):
        if (self.validities[0] and self.checkBox.isChecked() and self.validities[1]) or ((not self.checkBox.isChecked()) and self.validities[0]):
             self.nextbutton.setEnabled(1)
        else:
             self.nextbutton.setEnabled(0)

    def helpdialog(self):
        # Create and show the Help dialog
        help_dialog = QDialog()
        help_ui_file = "src/ui/dialog.ui"
        uic.loadUi(help_ui_file, help_dialog)
        help_dialog.okhelp.clicked.connect(help_dialog.accept)
        help_dialog.exec_()

 
                
#-----------------------------------------------Common Functions-----------------------------------------------------------
class xl:
    def sheets(datafile):
        workbook=openpyxl.load_workbook(datafile) 
        sheets=workbook.sheetnames
        return sheets
    def data(datafile,sheet,toggle):
         workbook=openpyxl.load_workbook(datafile)
         worksheet=workbook[sheet]
         num_columns = worksheet.max_column
         header=[]
         if not toggle ==2:
              for i in range(num_columns):
                   header.append(f"Column {i+1}")
         else:
              header_row = next(worksheet.iter_rows(min_row=1, max_row=1))
              header = [cell.value for cell in header_row]
                   
         return worksheet, header

#-----------------------------------------------Encryption and Decryption-----------------------------------------------------------
class crypt:

     def keygen(self,file):
          key = Fernet.generate_key()
          with open(file, 'wb') as f:
               pickle.dump(key,f)

     def loadkey(self):
          bin=os.path.join(os.getcwd(),'bin') 
          clip=os.path.join(bin,'clip.bin')
          if not os.path.exists(clip):
              try:
                   self.keygen(clip)
              except:
                   os.mkdir(bin)
                   self.keygen(clip)

          with open(clip, "rb") as f:
               key = pickle.load(f)
          return key

     def ncrypt(self,data,filename,key):
          fernet=Fernet(key)
          encrypted=fernet.encrypt(data.encode())
          with open( filename ,'wb') as f:
               pickle.dump(encrypted,f)

     def dcrypt(self,filename,key):
          with open(filename,'rb') as f:
               cryptdata=pickle.load(f)
          fernet=Fernet(key)
          
          data=fernet.decrypt(cryptdata).decode()
          
          return data
          


def cleanup():
     x=os.path.join(os.getcwd(),'__pycache__')
     if os.path.exists(x):
          rmtree(x)
     else:
          pass


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    icon = QIcon("logo.ico")
    app.setWindowIcon(icon)
    window.show()
    app.aboutToQuit.connect(cleanup)
    sys.exit(app.exec_())

    