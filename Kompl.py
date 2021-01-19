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
        F.ust_cvet_videl_tab(tabl_mk)

        tabl_rc = self.ui.table_bd_rab_c
        spis_bd_rc = F.otkr_f(F.tcfg('bd_rab_c'),separ='|')
        F.zapoln_wtabl(self, spis_bd_rc, tabl_rc, 0, 0, '', '', isp_shapka=False, separ='')
        tabl_rc.setSelectionBehavior(1)
        tabl_rc.setSelectionMode(1)
        F.ust_cvet_videl_tab(tabl_rc)

        tabl_tar = self.ui.table_bd_tara
        spis_tar = F.otkr_f(F.scfg('osn_tara') + os.sep + 'osn_Тара.txt',separ='|')
        F.zapoln_wtabl(self, spis_tar, tabl_tar, 0, 0, '', '', isp_shapka=False, separ='')
        tabl_tar.setSelectionBehavior(1)
        tabl_tar.setSelectionMode(1)
        F.ust_cvet_videl_tab(tabl_tar)

        tabl_det = self.ui.table_sod_mk
        F.ust_cvet_videl_tab(tabl_det)

        tabl_potok = self.ui.table_potok
        self.obn_potok()
        tabl_potok.setSelectionBehavior(1)
        tabl_potok.setSelectionMode(1)
        F.ust_cvet_videl_tab(tabl_potok)

        tabl_rc_potok = self.ui.table_bd_rab_c_potok
        self.obn_rc_potok()
        tabl_rc_potok.setSelectionBehavior(1)
        tabl_rc_potok.setSelectionMode(1)
        F.ust_cvet_videl_tab(tabl_rc_potok)
        tabl_potok.doubleClicked.connect(self.spis_det_tar)

        but_cr_tar = self.ui.push_create_tar
        but_cr_tar.clicked.connect(self.ct_tar)

        but_dvig = self.ui.push_dvijenie
        but_dvig.clicked.connect(self.dvijenie)

        but_rasform = self.ui.push_rasform
        but_rasform.clicked.connect(self.rasformirovat)

        but_vydat = self.ui.push_vydat
        but_vydat.clicked.connect(self.vydat)

    def spis_det_tar(self):
        tabl_potok = self.ui.table_potok
        tabl_rc_potok = self.ui.table_bd_rab_c_potok
        if tabl_potok.currentRow() == None or tabl_potok.currentRow() == -1:
            showDialog(self, 'Не выбрана тара')
            return
        nom_tar = tabl_potok.item(tabl_potok.currentRow(), 0).text()
        spis = F.otkr_f(F.scfg('bd_tara') + os.sep + nom_tar + '.txt', separ="|")
        s = 'Список деталей в ' + tabl_potok.item(tabl_potok.currentRow(), 5).text()  + ' №' + nom_tar + ":" + '\n'
        for i in spis:
            s = s + i[1] + ' ' + i[2] + ' ' + i[0] + 'шт.' + '\n'
        showDialog(self,'*скопирован в буфер.\n' + s)
        F.copy_bufer(s)

    def vydat(self):
        tabl_potok = self.ui.table_potok
        tabl_rc_potok = self.ui.table_bd_rab_c_potok
        if tabl_potok.currentRow() == None or tabl_potok.currentRow() == -1:
            showDialog(self, 'Не выбрана тара')
            return
        nom_tar = tabl_potok.item(tabl_potok.currentRow(), 0).text()
        sp_tar = F.otkr_f(F.tcfg('arh_tar'), separ='|')
        for i in range(0, len(sp_tar)):
            if sp_tar[i][0] == nom_tar:
                sp_tar[i][6] = 'выдано'
                break
        F.zap_f(F.tcfg('arh_tar'), sp_tar, '|')
        self.obn_potok()
        return

    def rasformirovat(self):
        tabl_potok = self.ui.table_potok
        tabl_rc_potok = self.ui.table_bd_rab_c_potok
        if tabl_potok.currentRow() == None or tabl_potok.currentRow() == -1:
            showDialog(self, 'Не выбрана тара')
            return
        nom_tar = tabl_potok.item(tabl_potok.currentRow(), 0).text()
        sp_tar = F.otkr_f(F.tcfg('arh_tar'), separ='|')
        for i in range(0, len(sp_tar)):
            if sp_tar[i][0] == nom_tar:
                sp_tar[i][6] = 'закрыта'
                break
        F.zap_f(F.tcfg('arh_tar'), sp_tar, '|')
        self.obn_potok()
        return


    def dvijenie(self):
        tabl_potok = self.ui.table_potok
        tabl_rc_potok = self.ui.table_bd_rab_c_potok
        if tabl_potok.currentRow() == None or tabl_potok.currentRow() == -1:
            showDialog(self,'Не выбрана тара')
            return
        if tabl_rc_potok.currentRow() == None or tabl_rc_potok.currentRow() == -1:
            showDialog(self,'Не выбран рабочий центр')
            return
        rab_c = tabl_rc_potok.item(tabl_rc_potok.currentRow(),0).text()
        dop = F.tek_polz()+"$"+F.now()+"$"+rab_c
        nom_tar = tabl_potok.item(tabl_potok.currentRow(),0).text()
        sp_tar = F.otkr_f(F.tcfg('arh_tar'), separ='|')
        for i in range(0, len(sp_tar)):
            if sp_tar[i][0] == nom_tar:
                sp_tar[i][8] = sp_tar[i][8] + "-->" + dop
                break
        F.zap_f(F.tcfg('arh_tar'),sp_tar,'|')
        self.obn_potok()
        return


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
        if tabl_mk.currentRow() == None or tabl_mk.currentRow() == -1:
            showDialog(self,'Не выбрана маршрутная карта')
            return
        s =[]
        spis = F.spisok_iz_wtabl(tabl_det,'',shapka=True)
        for i in range(1,len(spis)):
            if spis[i][0] != '':
                if F.is_numeric(spis[i][0]) == False:
                    showDialog(self,'В кол-во деталей ' +  spis[i][3] + ' '  + spis[i][4] +
                               ' введены некорректные данные: ' + spis[i][0])
                    return
                if int(spis[i][0]) > int(spis[i][1]):
                    showDialog(self, 'Кол-во деталей ' +  spis[i][3] + ' '  + spis[i][4] +
                               '  превышает доступное ' + spis[i][1])
                    return
                s.append([spis[i][0],spis[i][3].strip(),spis[i][4].strip()])
        if len(s) == 0:
            showDialog(self, 'Не выбраны детали')
            return

        sp_tar = F.otkr_f(F.tcfg('arh_tar'),separ='|')
        nom = int(sp_tar[len(sp_tar) - 1][0]) + 1
        nom = '0' * (8 - len(str(nom))) + str(nom)
        vid_tar = tabl_tar.item(tabl_tar.currentRow(),0).text()
        rab_c = tabl_rc.item(tabl_rc.currentRow(),0).text()
        nom_mk = tabl_mk.item(tabl_mk.currentRow(),0).text()

        sp_tar.append([nom,F.tek_polz(),F.date(2),nom_mk,'',vid_tar,'открыта',text_prim.text(),F.tek_polz()+"$"+F.now()+"$"+rab_c])
        F.zap_f(F.tcfg('arh_tar'),sp_tar,'|')
        F.zap_f(F.scfg('bd_tara')+ os.sep + nom + '.txt',s,'|')
        tabl_det.clear()
        self.obn_potok()
        showDialog(self,'Taра  №' + nom + '  ' + vid_tar + "  успено сформирована.")

    def obn_rc_potok(self):
        tabl_rc_potok = self.ui.table_bd_rab_c_potok
        spis = F.otkr_f(F.tcfg('bd_rab_c'),separ='|')
        F.zapoln_wtabl(self, spis, tabl_rc_potok, 0, 0, '', '', isp_shapka=False, separ='')

    def obn_potok(self):
        tabl_potok = self.ui.table_potok
        bd_arh = F.otkr_f(F.tcfg('arh_tar'),separ='|')
        s= []
        s.append(bd_arh[0])
        for i in bd_arh:
            if i[6] == 'открыта':
                s.append(i)
        F.zapoln_wtabl(self,s,tabl_potok,0,0,"","",200,True,'',30)

    def dost_ostatok_det(self,spisok,nom_mk):
        for i in range(1,len(spisok)):
            spisok[i][1] = spisok[i][5]
            spisok[i][2] = '0'
        bd_arh_tar = F.otkr_f(F.tcfg('arh_tar'),separ='|')
        for i in bd_arh_tar:
            if i[3] == nom_mk:
                if i[6] == 'открыта' or i[6] == 'выдано':
                    sost = i[6]
                    nom = i[0]
                    det_tmp = F.otkr_f(F.scfg('bd_tara') + os.sep + nom + '.txt',separ='|')
                    for i in range(0,len(det_tmp)):
                        for j in range(1,len(spisok)):
                            if det_tmp[i][1].strip() == spisok[j][3].strip() and det_tmp[i][2].strip() == spisok[j][4].strip():
                                spisok[j][1] = int( spisok[j][1]) - int(det_tmp[i][0])
                                if sost == 'выдано':
                                    spisok[j][2] = int(spisok[j][2]) + int(det_tmp[i][0])
                                break
        return spisok

    def zagruz_det_iz_mk(self):
        tabl_mk = self.ui.table_bd_mk
        tabl_det = self.ui.table_sod_mk

        nom = tabl_mk.item(tabl_mk.currentRow(),0).text()
        spis = F.otkr_f(F.scfg('bd_mk')+os.sep + nom+ '.txt',separ="|")
        s = []
        max_dop_dlina = 0
        for i in range(0, len(spis)):
            dop_dlina = 0
            tmp = ['','','',spis[i][0],spis[i][1],spis[i][2],spis[i][3],spis[i][4],spis[i][5],spis[i][6],spis[i][7],spis[i][8]]

            for j in range(9, len(spis[0])):
                if spis[0][j] == 'комплектация' and spis[i][j-1] != "":
                    if i > 0:
                        tmp.append(spis[0][j-1])
                        dop_dlina+=1
            s.append(tmp)
            if dop_dlina > max_dop_dlina:
                max_dop_dlina = dop_dlina
        for i in range(0,max_dop_dlina):
            s[0].append('')
        s[0][0] = 'Факт. кол.'
        s[0][1] = 'Дост. кол.'
        s[0][2] = 'Выдано'

        nom_mk = tabl_mk.item(tabl_mk.currentRow(), 0).text()
        s = self.dost_ostatok_det(s,nom_mk)

        ed_kol = {0}
        F.zapoln_wtabl(self, s, tabl_det, 0, ed_kol,'', '', isp_shapka=True, separ='')
        tabl_det.setSelectionMode(1)
        for i in range(0,len(s)-1):
            for j in range(1, len(s[i])):
                F.dob_color_wtab(tabl_det,i,j,16,16,16)
            if i > 0 and int(s[i][2]) == int(s[i][5]):
                F.dob_color_wtab(tabl_det, i-1, 0, 16, 16, 16)
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