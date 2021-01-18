import Cust_Functions as F

from PyQt5 import QtWidgets, QtCore, QtGui
from PyQt5.QtWinExtras import QtWin
import os
import time
import subprocess
from mydesign import Ui_MainWindow  # импорт нашего сгенерированного файла
#from mydesign2 import Ui_Dialog  # импорт нашего сгенерированного файла
import sys

def showDialog(self, msg):
    msgBox = QtWidgets.QMessageBox()
    msgBox.setIcon(QtWidgets.QMessageBox.Information)
    msgBox.setText(msg)
    msgBox.setWindowTitle("Внимание!")
    msgBox.setStandardButtons(QtWidgets.QMessageBox.Ok)  # | QtWidgets.QMessageBox.Cancel)
    returnValue = msgBox.exec()

#class mywindow2(QtWidgets.QDialog):  # диалоговое окно
#    def __init__(self,parent=None,item_o="",p1=0,p2=0):
#        self.item_o = item_o
#        self.p1 = p1
#        self.p2 = p2
#        self.myparent = parent
#        super(mywindow2, self).__init__()
#        self.ui2 = Ui_Dialog()
#        self.ui2.setupUi(self)
#        self.setWindowModality(QtCore.Qt.ApplicationModal)

class mywindow(QtWidgets.QMainWindow):
    resized = QtCore.pyqtSignal()
    def __init__(self):
        super(mywindow, self).__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.setFocusPolicy(QtCore.Qt.StrongFocus)

        #self.resized.connect(self.widths)

        tabl_mk = self.ui.table_bd_mk
        spis_bd_mk = F.otkr_f(F.tcfg('bd_mk'),separ='|')
        s = []
        for i in spis_bd_mk:
            if i[2] != 'Закрыта':
                s.append(i)
        F.zapoln_wtabl(self,s,tabl_mk,0,0,'','',isp_shapka=True,separ='')
        tabl_mk.setSelectionBehavior(1)
        tabl_mk.clicked.connect(self.zagruz_det_iz_mk)
        tabl_mk.setSelectionMode(1)

        tabl_rc = self.ui.table_bd_rab_c
        spis_bd_rc = F.otkr_f(F.tcfg('bd_rab_c'),separ='|')
        F.zapoln_wtabl(self, spis_bd_rc, tabl_rc, 0, 0, '', '', isp_shapka=False, separ='')
        tabl_rc.setSelectionBehavior(1)
        tabl_rc.setSelectionMode(1)

        tabl_tar = self.ui.table_bd_tara
        spis_tar = F.otkr_f(F.scfg('osn_tara') + os.sep + 'osn_Тара.txt',separ='|')
        F.zapoln_wtabl(self, spis_tar, tabl_tar, 0, 0, '', '', isp_shapka=False, separ='')
        tabl_tar.setSelectionBehavior(1)
        tabl_tar.setSelectionMode(1)

        but_cr_tar = self.ui.push_create_tar
        but_cr_tar.clicked.connect(self.ct_tar)


    def ct_tar(self):
        tabl_mk = self.ui.table_bd_mk
        tabl_rc = self.ui.table_bd_rab_c
        tabl_tar = self.ui.table_bd_tara
        tabl_det = self.ui.table_sod_mk
        text_prim = self.ui.line_prim_cr_tar

        if tabl_rc.currentRow() == None or tabl_rc.currentRow() == -1:
            showDialog(self,'Не выбран рабочий центр')
            return
        if tabl_tar.currentRow() == None or tabl_tar.currentRow() == -1:
            showDialog(self,'Не выбран вид тары')
            return
        s =[]
        spis = F.spisok_iz_wtabl(tabl_det,'',shapka=True)
        for i in range(1,len(spis)):
            if spis[i][0] != '':
                if F.is_numeric(spis[i][0]) == False:
                    showDialog(self,'В кол-во деталей введены некорректные данные ' + spis[i][0])
                    return
                s.append([spis[i][0],spis[i][1].strip(),spis[i][2].strip()])
        if len(s) == 0:
            showDialog(self, 'Не выбраны детали')
            return

        sp_tar = F.otkr_f(F.tcfg('arh_tar'),separ='|')
        nom = int(sp_tar[len(sp_tar) - 1][0]) + 1
        nom = '0' * (8 - len(str(nom))) + str(nom)
        vid_tar = tabl_tar.item(tabl_tar.currentRow(),0).text()
        rab_c = tabl_rc.item(tabl_rc.currentRow(),0).text()
        sp_tar.append([nom,F.tek_polz(),F.date(2),'',vid_tar,'открыта',text_prim.text(),rab_c])
        F.zap_f(F.tcfg('arh_tar'),sp_tar,'|')
        F.zap_f(F.scfg('bd_tara')+ os.sep + nom + '.txt',s,'|')
        showDialog(self,'Taра ' + vid_tar + " успено сформирована. №" + nom)


    def zagruz_det_iz_mk(self):
        tabl_mk = self.ui.table_bd_mk
        tabl_det = self.ui.table_sod_mk

        nom = tabl_mk.item(tabl_mk.currentRow(),0).text()
        spis = F.otkr_f(F.scfg('bd_mk')+os.sep + nom+ '.txt',separ="|")
        s = []
        for i in spis:
            s.append(['',i[0],i[1],i[2],i[3],i[4],i[5],i[6],i[7],i[8]])
        s[0][0] = 'Факт. кол.'
        ed_kol = {0}
        F.zapoln_wtabl(self, s, tabl_det, 0, ed_kol,'', '', isp_shapka=True, separ='')
        tabl_det.setSelectionMode(1)
        for i in range(0,len(s)-1):
            for j in range(1, len(s[i])):
                F.dob_color_wtab(tabl_det,i,j,16,16,16)
                #tabl_det.item().background().color().rgb(QtGui.QColor.red())

app = QtWidgets.QApplication([])

myappid = 'Powerz.BAG.SustControlWork.0.0.0'  # !!!
QtWin.setCurrentProcessExplicitAppUserModelID(myappid)
app.setWindowIcon(QtGui.QIcon(os.path.join("icons", "icon.png")))

S = F.scfg('Stile').split(",")
app.setStyle(S[1])

application = mywindow()
application.show()

sys.exit(app.exec())