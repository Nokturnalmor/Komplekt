import Cust_Functions as F
import copy
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

def showDialogYNC(self, msg):
    msgBox = QtWidgets.QMessageBox()
    msgBox.setIcon(QtWidgets.QMessageBox.Information)
    msgBox.setText(msg)
    msgBox.setWindowTitle("Внимание!")
    msgBox.setStandardButtons(QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No | QtWidgets.QMessageBox.Cancel)
    returnValue = msgBox.exec()
    if returnValue == 123:
        return 'Yes'
    if returnValue == 123:
        return 'No'
    if returnValue == 123:
        return 'Cancel'
    return 'False'

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
        self.setWindowTitle("Диспетчирование")
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



        tabl_tar = self.ui.table_bd_tara
        spis_tar = F.otkr_f(F.scfg('osn_tara') + os.sep + 'osn_Тара.txt', separ='|')
        F.zapoln_wtabl(self, spis_tar, tabl_tar, 0, 0, '', '', isp_shapka=False, separ='')
        tabl_tar.setSelectionBehavior(1)
        tabl_tar.setSelectionMode(1)
        F.ust_cvet_videl_tab(tabl_tar)

        tabl_det = self.ui.table_sod_mk
        F.ust_cvet_videl_tab(tabl_det)
        tabl_det.doubleClicked.connect(self.dblclk_tab_det)
        tabl_det.clicked.connect(self.clk_tab_det)

        tabl_potok = self.ui.table_potok
        self.obn_potok()
        tabl_potok.setSelectionBehavior(1)
        tabl_potok.setSelectionMode(1)
        F.ust_cvet_videl_tab(tabl_potok)
        tabl_potok.doubleClicked.connect(self.spis_det_tar)
        #tabl_potok.clicked.connect(self.obn_ob_mar)


        tabl_rc_potok = self.ui.table_bd_rab_c_potok
        self.obn_rc_potok()
        tabl_rc_potok.setSelectionBehavior(1)
        tabl_rc_potok.setSelectionMode(1)
        F.ust_cvet_videl_tab(tabl_rc_potok)
        tabl_rc_potok.clicked.connect(self.clk_rc)



        #tabl_ob_mar = self.ui.table_obsh_marshrut
        #tabl_potok.setSelectionBehavior(1)
        #tabl_ob_mar.setSelectionMode(1)
        #F.ust_cvet_videl_tab(tabl_ob_mar)
        #tabl_ob_mar.verticalHeader().setVisible(False)
        #tabl_ob_mar.horizontalHeader().setVisible(False)

        but_cr_tar = self.ui.push_create_tar
        but_cr_tar.clicked.connect(self.ct_tar)

        but_dvig = self.ui.push_dvijenie
        but_dvig.clicked.connect(self.dvijenie)

        but_rasform = self.ui.push_rasform
        but_rasform.clicked.connect(self.rasformirovat)

        but_vydat = self.ui.push_vydat
        but_vydat.clicked.connect(self.vydat)

        tab = self.ui.tabWidget
        tab.tabBarClicked.connect(self.tab_click)

    def clk_rc(self):
        tabl_det = self.ui.table_sod_mk
        tabl_det.setCurrentCell(-1,-1)

    def tab_click(self):
        tab = self.ui.tabWidget
        if tab.currentIndex() == 0:
            self.zagruz_det_iz_mk()

    def clk_tab_det(self):
        vne_marshruta = self.obn_rc_potok()
        tabl_det = self.ui.table_sod_mk
        if vne_marshruta == True:
            F.ust_color_wtab(tabl_det,tabl_det.currentRow(),5,254, 115, 0)
        else:
            F.ust_color_wtab(tabl_det, tabl_det.currentRow(), 5, 233, 233, 233)

    def dblclk_tab_det(self):
        tabl_det = self.ui.table_sod_mk
        r = tabl_det.currentRow()
        k = tabl_det.currentColumn()
        if k == 6:
            self.sost_tary_iz_mk(r,k)
        if k > 14:
            self.sost_det_po_rc()

    def sost_det_po_rc(self):
        tabl_det = self.ui.table_sod_mk
        tabl_mk = self.ui.table_bd_mk
        id = tabl_det.item(tabl_det.currentRow(),11).text()
        arr = tabl_det.item(tabl_det.currentRow(),tabl_det.currentColumn()).text().split('_')
        rc = arr[0]
        nom_mk = tabl_mk.item(tabl_mk.currentRow(), 0).text()
        bd_arh_tar = F.otkr_f(F.tcfg('arh_tar'), separ='|')
        s = self.tek_pol_det(id,nom_mk,'',bd_arh_tar)
        s_mar = s[0]
        s_vne_mar = s[1]
        s_bez_tar_po_mar = s[2]
        s_bez_tar_vne_mar = s[3]
        s_vidano = s[4]
        s_v_tare_po_mar = s[5]
        s_v_tare_vne_mar = s[6]
        s_pererabot_k_perem = s[7]
        msg = 'По маршруту: ' + "\n" + rc + ' '
        for i in range(len(s_mar)):
            if s_mar[i][0] == rc:
                msg+= str(s_mar[i][1]) + ' шт.'+  "\n"
        msg+= "\n" + 'Вне маршурта:'+ "\n"
        for i in s_vne_mar.keys():
            if s_vne_mar[i] > 0:
                msg+= i + ' ' + str(s_vne_mar[i]) + ' шт.' +  "\n"
        msg += "\n" + 'Без тары:' + "\n" + rc + ' '
        for i in range(len(s_bez_tar_po_mar)):
            if s_bez_tar_po_mar[i][0] == rc:
                msg+= 'По маршруту: ' + str(s_bez_tar_po_mar[i][1]) + ' шт.'+  "\n"
        for i in s_bez_tar_vne_mar.keys():
            if s_bez_tar_vne_mar[i] > 0:
                msg += 'Вне маршурта: ' + i + ' ' + str(s_bez_tar_vne_mar[i]) + ' шт.' + "\n"
        msg += "\n" + 'Выдано:' + "\n" + rc + ' '
        for i in range(len(s_vidano)):
            if s_vidano[i][0] == rc:
                msg+= str(s_vidano[i][1]) + ' шт.'+  "\n"
        msg += "\n" + 'В таре:' + "\n" + rc + ' '
        for i in range(len(s_v_tare_po_mar)):
            if s_v_tare_po_mar[i][0] == rc:
                msg += 'По маршруту: ' + str(s_v_tare_po_mar[i][1]) + ' шт.'+  "\n"
        for i in s_v_tare_vne_mar.keys():
            if s_v_tare_vne_mar[i] > 0:
                msg += 'Вне маршурта: ' + i + ' ' + str(s_v_tare_vne_mar[i]) + ' шт.' + "\n"
        msg += "\n" + 'Переработано для перемещения:' + "\n" + rc + ' '
        for i in range(len(s_pererabot_k_perem)):
            if s_pererabot_k_perem[i][0] == rc:
                msg += str(s_pererabot_k_perem[i][1]) + ' шт.' + "\n"

        F.msgbox(msg)
        return



    def sost_tary_iz_mk(self,r,c):
        tabl_det = self.ui.table_sod_mk
        tabl_mk = self.ui.table_bd_mk
        naim = tabl_det.item(r,3).text().strip()
        nn = tabl_det.item(r,4).text().strip()
        nom_mk = tabl_mk.item(tabl_mk.currentRow(),0).text()
        bd_arh_tar = F.otkr_f(F.tcfg('arh_tar'), separ='|')
        s=''
        for i in bd_arh_tar:
            if i[3] == nom_mk:
                sost = i[6]
                nom = i[0]
                marsh = i[9]
                nazv = i[5]
                det_tmp = F.otkr_f(F.scfg('bd_tara') + os.sep + nom + '.txt', separ='|')
                for i in range(0, len(det_tmp)):
                    if det_tmp[i][1].strip() == naim and det_tmp[i][2].strip() == nn:
                        s+= sost + ' ' + nom + ' ' + nazv + '\n' + marsh + '\n' + '\n'
        showDialog(self, s)
        return





    def spis_det_tar(self):
        tabl_potok = self.ui.table_potok
        tabl_rc_potok = self.ui.table_bd_rab_c_potok
        if tabl_potok.currentRow() == None or tabl_potok.currentRow() == -1:
            showDialog(self, 'Не выбрана тара')
            return
        nom_mk = tabl_potok.item(tabl_potok.currentRow(), 3).text()
        nom_tar = tabl_potok.item(tabl_potok.currentRow(), 0).text()
        spis = F.otkr_f(F.scfg('bd_tara') + os.sep + nom_tar + '.txt', separ="|")
        sp_det_mar = self.spis_det_s_ih_marsh(nom_mk)
        s = 'Список деталей в ' + tabl_potok.item(tabl_potok.currentRow(), 5).text()  + ' №' + nom_tar + ":" + '\n'
        for i in spis:
            mar_tmp = []
            for j in range(len(sp_det_mar)):
                if sp_det_mar[j][3].strip() == i[1] and sp_det_mar[j][4].strip() == i[2]:
                    for k in range(16, len(sp_det_mar[j])):
                        mar_tmp.append(sp_det_mar[j][k])
                    break
            mar_tmp = '-->'.join(mar_tmp)
            s = s + i[1] + ' ' + i[2] + ' ' + i[0] + 'шт. Маршрут: ' + mar_tmp + '\n'
        showDialog(self,'*скопирован в буфер.\n' + s)
        F.copy_bufer(s)

    def spis_det_s_ih_marsh(self,nom_mk):
        spis = F.otkr_f(F.scfg('bd_mk') + os.sep + nom_mk + '.txt', separ="|")
        s = []
        max_dop_dlina = 0
        for i in range(0, len(spis)):
            dop_dlina = 0
            tmp = ['', '', '', spis[i][0], spis[i][1], spis[i][2], "", "", spis[i][3], spis[i][4], spis[i][5],
                   spis[i][6], spis[i][7], spis[i][8], spis[i][9], spis[i][10]]

            for j in range(11, len(spis[0])):
                if spis[0][j] == 'комплектация' and spis[i][j - 1] != "":
                    if i > 0:
                        tmp.append(spis[0][j - 1])
                        dop_dlina += 1
            s.append(tmp)
            if dop_dlina > max_dop_dlina:
                max_dop_dlina = dop_dlina
        for i in range(0, max_dop_dlina):
            s[0].append('        ')
        return s

    def vse_det_v_odin_rc(self,nom_mk,nom_tar,rc):
        s = self.spis_det_s_ih_marsh(nom_mk)
        spis_tar = F.otkr_f(F.scfg('bd_tara') + os.sep + nom_tar + '.txt', separ='|')

        for i in range(len(spis_tar)):
            flag_rc = False
            for j in range(len(s)):
                if spis_tar[i][1].strip() == s[j][3].strip() and spis_tar[i][2].strip() == s[j][4].strip():
                    for k in range(15,len(s[j])):
                        if rc == s[j][k]:
                            flag_rc = True
                            break
                    break
            if flag_rc == False:
                return False
        return True

    def vydat(self):

        tabl_potok = self.ui.table_potok
        tabl_rc_potok = self.ui.table_bd_rab_c_potok
        tabl_det = self.ui.table_sod_mk
        if tabl_potok.currentRow() == None or tabl_potok.currentRow() == -1:
            showDialog(self, 'Не выбрана тара')
            return
        if tabl_det.currentRow() == None or tabl_det.currentRow() == -1:
            showDialog(self, 'Не выбран РЦ')
            return
        if tabl_det.currentColumn() < 16 or tabl_det.currentItem() == None:
            showDialog(self, 'Не выбран РЦ')
            return
        nom_tar = tabl_potok.item(tabl_potok.currentRow(), 0).text()
        nom_mk = tabl_potok.item(tabl_potok.currentRow(), 3).text()

        arr_rc = tabl_det.currentItem().text().split('_')
        rc = arr_rc[0]
        if self.vse_det_v_odin_rc(nom_mk,nom_tar,rc) == False:
            showDialog(self, 'Нужно переформировать. Не все детали в таре ' + nom_tar + " имеют общий РЦ :" + rc)
            return
        sp_tar = F.otkr_f(F.tcfg('arh_tar'), separ='|')
        for i in range(0, len(sp_tar)):
            if sp_tar[i][0] == nom_tar:
                if rc != self.tek_rc(sp_tar[i][9]):
                    showDialog(self, 'Нельзя выдать в ' + rc + ", т.к. тара находится в " + self.tek_rc(sp_tar[i][9]))
                    return
                sp_tar[i][6] = 'выдано'
                sp_tar[i][7] = str(tabl_det.currentColumn()-15)
                break
        self.uchet_komplektacii(sp_tar,nom_tar,nom_mk,rc,tabl_det.currentColumn()-15)
        F.zap_f(F.tcfg('arh_tar'), sp_tar, '|')
        self.obn_potok()
        self.dost_ostatok_det()
        self.oform_cveta()
        showDialog(self, 'Выдача записана успешно')
        return

    def uchet_komplektacii(self,spis_arh_tar,nom_tar,nom_mk,rc,nomer_rc):

        spis_mk = F.otkr_f(F.scfg('bd_mk') + os.sep + nom_mk + '.txt', separ="|")
        spis_tar = F.otkr_f(F.scfg('bd_tara') + os.sep + nom_tar + '.txt',separ='|')
        for i in range(len(spis_tar)):
            nomer = nomer_rc
            for j in range(len(spis_mk)):
                if spis_tar[i][3].strip() == spis_mk[j][6].strip():
                    kol_vyd = 0
                    for k in range(len(spis_arh_tar)):
                        if spis_arh_tar[k][3] == nom_mk and spis_arh_tar[k][6] == 'выдано' and str(nomer_rc) == spis_arh_tar[k][7]:
                            if self.tek_rc(spis_arh_tar[k][9]) == rc:
                                spis_tar_tmp = F.otkr_f(F.scfg('bd_tara') + os.sep + spis_arh_tar[k][0] + '.txt', separ='|')
                                for l in range(len(spis_tar_tmp)):
                                    if spis_tar_tmp[l][3].strip() == spis_mk[j][6].strip():
                                        kol_vyd+= int(spis_tar_tmp[l][0])
                    for l in range(11,len(spis_mk[j]),4):
                        if spis_mk[j][l] != '' and nomer==1:
                            break
                        if spis_mk[j][l] != '':
                            nomer-=1
                    nomer_rc_v_mk = l+1
                    if kol_vyd == int(spis_mk[j][2]):
                        spis_mk[j][nomer_rc_v_mk] = str(kol_vyd) + ' шт. (полный компл.)'
                    elif kol_vyd < int(spis_mk[j][2]):
                        if kol_vyd//int(spis_mk[j][7]) >= 1:
                            spis_mk[j][nomer_rc_v_mk] = str(kol_vyd) + ' шт. (' + str(
                                int(kol_vyd//int(spis_mk[j][7]))) + 'компл.)'
                    break
        F.zap_f(F.scfg('bd_mk') + os.sep + nom_mk + '.txt',spis_mk,'|')
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
                sp_tar[i][4] = F.now()
                break
        F.zap_f(F.tcfg('arh_tar'), sp_tar, '|')
        self.obn_potok()
        return


    def dvijenie(self):
        tabl_potok = self.ui.table_potok
        tabl_rc_potok = self.ui.table_bd_rab_c_potok
        tabl_det = self.ui.table_sod_mk
        if tabl_potok.currentRow() == None or tabl_potok.currentRow() == -1:
            showDialog(self,'Не выбрана тара')
            return

        dop = ''
        if tabl_rc_potok.currentRow() >= 0:
            rab_c = tabl_rc_potok.item(tabl_rc_potok.currentRow(), 0).text()
            dop = F.tek_polz()+"$"+F.now()+"$0$"+rab_c
        if tabl_det.currentRow() >= 0:
            arr_rab_c = tabl_det.currentItem().text().split('_')
            rab_c = arr_rab_c[0]
            dop = F.tek_polz() + "$" + F.now() + "$" + str(tabl_det.currentColumn() - 15) + "$" + rab_c

        if dop == '':
            F.msgbox('Не выбран РЦ')
            return


        nom_tar = tabl_potok.item(tabl_potok.currentRow(),0).text()
        sp_tar = F.otkr_f(F.tcfg('arh_tar'), separ='|')
        for i in range(0, len(sp_tar)):
            if sp_tar[i][0] == nom_tar:
                tek_dvig = sp_tar[i][9].split("-->")[-1]
                tek_pol = tek_dvig.split("$")[-1]
                if tek_pol == rab_c:
                    F.msgbox('Тара уже находится в ' + tek_pol)
                    return
                break

        for i in range(0, len(sp_tar)):
            if sp_tar[i][0] == nom_tar:
                sp_tar[i][9] = sp_tar[i][9] + "-->" + dop
                break
        F.zap_f(F.tcfg('arh_tar'),sp_tar,'|')
        self.obn_potok()
        self.dost_ostatok_det()
        self.oform_cveta()
        return


    def ct_tar(self):
        tabl_mk = self.ui.table_bd_mk
        tabl_rc = self.ui.table_bd_rab_c_potok
        tabl_tar = self.ui.table_bd_tara
        tabl_det = self.ui.table_sod_mk
        text_prim = self.ui.line_prim_cr_tar

        if tabl_det.currentColumn() < 16 or tabl_det.currentRow() == -1:
            if tabl_rc.currentRow() != -1:
                rab_c = tabl_rc.item(tabl_rc.currentRow(),0).text()
                dost_kol = int(tabl_rc.item(tabl_rc.currentRow(),2).text())
                flag = 1
            else:
                showDialog(self,'Не выбран рабочий центр')
                return
        else:
            arr_rab_c = tabl_det.item(tabl_det.currentRow(), tabl_det.currentColumn()).text().split('_')
            rab_c = arr_rab_c[0]
            flag = 0

        if tabl_tar.currentRow() == None or tabl_tar.currentRow() == -1:
            showDialog(self,'Не выбран вид тары')
            return
        if tabl_mk.currentRow() == None or tabl_mk.currentRow() == -1:
            showDialog(self,'Не выбрана маршрутная карта')
            return
        s =[]
        spis = F.spisok_iz_wtabl(tabl_det,'',shapka=True)
        mk = tabl_mk.item(tabl_mk.currentRow(),0).text().strip()
        if F.nalich_file(F.scfg('mk_data') + os.sep + mk + '.txt') == False:
            showDialog(self, 'Не обнаружен файл')
            return
        sp_mk = F.otkr_f(F.scfg('mk_data') + os.sep + mk + '.txt',False,'|')
        if sp_mk == []:
            showDialog(self, 'Некорректное содержимое МК')
            return
        for i in range(1,len(spis)):
            if spis[i][0] != '':
                if F.is_numeric(spis[i][0]) == False:
                    showDialog(self,'В кол-во деталей ' +  spis[i][3] + ' '  + spis[i][4] +
                               ' введены некорректные данные: ' + spis[i][0])
                    return
                if flag == 0:
                    dost_kol_arr = tabl_det.currentItem().text().split('_')
                    dost_kol = dost_kol_arr[-1]

                    if int(spis[i][0]) > int(dost_kol):
                        showDialog(self, 'Кол-во деталей ' +  spis[i][3] + ' '  + spis[i][4] +
                                           '  превышает доступное ' + spis[i][1])
                        return

                bd_arh = F.otkr_f(F.tcfg('arh_tar'), separ='|')
                if flag == 0:
                    if self.tek_pol_det(spis[i][11].strip(),mk,tabl_det.currentColumn()-16,bd_arh) < 1:
                        showDialog(self,  spis[i][3] + ' '  + spis[i][4] +' не находится на ' + rab_c)
                        return

                id = spis[i][11]
                s_tmp = self.tek_pol_det(id, mk, '', bd_arh)
                s_bez_tar_po_mar = s_tmp[2]
                s_bez_tar_vne_mar = s_tmp[3]
                s_pererabot_k_perem = s_tmp[7]
                for j in range(len(s_bez_tar_po_mar)):
                    if s_bez_tar_po_mar[j][0] == rab_c:
                        bez_tar_po_mar = int(s_bez_tar_po_mar[j][1])
                        break
                for j in s_bez_tar_vne_mar.keys():
                    if j == rab_c:
                        bez_tar_vne_mar = int(s_bez_tar_vne_mar[j])
                        break
                for j in range(len(s_pererabot_k_perem)):
                    if s_pererabot_k_perem[j][0] == rab_c:
                        pererabot_k_perem = int(s_pererabot_k_perem[j][1])
                        break
                s_vidano = s_tmp[4]
                for j in range(len(s_vidano)):
                    if s_vidano[j][0] == rab_c:
                        vidano = int(s_vidano[j][1])
                        break

                if flag == 1:
                    bez_tar = bez_tar_vne_mar
                else:
                    bez_tar = bez_tar_po_mar

                ##if pererabot_k_perem > 0 and bez_tar > 0:
                 #   rez = showDialogYNC(self, 'Доступно: до обработки ' + str(bez_tar) + ' шт., ' +
                 #                             'после обработки ' + str(pererabot_k_perem) + ' шт.' + n/ +
                 #                             'взять для перемещения ДО ОБРАБОТКИ?')
                 #   if rez == 'Cancel':
                 #       return
                 #   elif rez == 'No':
                 #       return
                 #   elif rez == 'Yes':
                 #       return


                if int(spis[i][0]) > bez_tar:
                    if pererabot_k_perem == 0:
                        showDialog(self, 'Кол-во деталей ' +  spis[i][3] + ' '  + spis[i][4] +
                                   '  превышает доступное в цеху, не выданное, без тары: ' + str(bez_tar))
                        return

                if pererabot_k_perem <  int(spis[i][0]):
                    showDialog(self, 'Кол-во деталей ' + spis[i][3] + ' ' + spis[i][4] +
                               '  превышает доступное в цеху, после обработки: ' + str(pererabot_k_perem))
                    return


                if vidano == spis[i][5]:
                    tmp = self.sost_rabot_po_det_rc_pornom(sp_mk,rab_c,spis[i][11],tabl_det.currentColumn()-15)
                    if tmp == False:
                        showDialog(self, 'Полный объем работ на ' + rab_c + ' для ' + spis[i][3].strip() +
                                   ' ' + spis[i][4].strip() + ' не выполнен.')
                        return

                s.append([spis[i][0],spis[i][3].strip(),spis[i][4].strip(),spis[i][11].strip()])
        if len(s) == 0:
            showDialog(self, 'Не выбраны детали')
            return

        sp_tar = F.otkr_f(F.tcfg('arh_tar'),separ='|')
        nom = int(sp_tar[len(sp_tar) - 1][0]) + 1
        nom = '0' * (8 - len(str(nom))) + str(nom)
        vid_tar = tabl_tar.item(tabl_tar.currentRow(),0).text()

        nom_mk = tabl_mk.item(tabl_mk.currentRow(),0).text()

        if flag == 0:
            sp_tar.append([nom,F.tek_polz(),F.date(2),nom_mk,'',vid_tar,'открыта',"",text_prim.text(),F.tek_polz()+"$"
                           +F.now()+"$"+ str(tabl_det.currentColumn()-15) + "$" + rab_c])
        else:
            sp_tar.append(
                [nom, F.tek_polz(), F.date(2), nom_mk, '', vid_tar, 'открыта', "", text_prim.text(), F.tek_polz() + "$"
                 + F.now() + "$0$" + rab_c])
        F.zap_f(F.tcfg('arh_tar'),sp_tar,'|')
        F.zap_f(F.scfg('bd_tara')+ os.sep + nom + '.txt',s,'|')
        self.dost_ostatok_det()
        self.obn_potok()
        for i in range(tabl_det.rowCount()):
            tabl_det.item(i,0).setText('')

        showDialog(self,'Taра  №' + nom + '  ' + vid_tar + "  успено сформирована.")

    def tek_pol_det(self,id,mk,rab_c,bd_arh):
        """
        :param id:
        :param mk:
        :param rab_c:
        :param bd_arh:
        :return:
        По маршруту, вне маршрута, s_bez_tar_po_mar, s_bez_tar_vne_mar, Выдано, s_v_tare_po_mar, s_v_tare_vne_mar, s_pererabot_k_perem
        """
        tabl_det = self.ui.table_sod_mk
        sp_nar = F.otkr_f(F.tcfg('Naryad'),separ='|')
        s = []
        s2 = dict()
        bd_rab_c = F.otkr_f(F.tcfg('bd_rab_c'),separ='|')
        for i in bd_rab_c:
            s2[i[0]] = 0
        s32 = copy.deepcopy(s2)
        s52 = copy.deepcopy(s2)
        spis = F.spisok_iz_wtabl(tabl_det, '', shapka=True)
        for i in range(len(spis)):
            if spis[i][11] == id:
                for j in range(16,len(spis[i])):
                    arr = spis[i][j].split('_')
                    tmp = arr[0]
                    s.append([tmp,0])


        s4 = copy.deepcopy(s)
        s5 = copy.deepcopy(s)
        s6 = copy.deepcopy(s)

        s[0][1] = spis[i][5]
        s3 = copy.deepcopy(s)
        mar = []
        for i in bd_arh:
            if i[3] == mk:
                sost = i[6]
                tmp_tar = F.otkr_f(F.scfg('bd_tara') + os.sep + i[0] + '.txt',separ='|')
                for j in range(len(tmp_tar)):
                    if tmp_tar[j][3] == id:
                        arr = i[9].split('-->')
                        arr1 = arr[0].split('$')
                        arr2 = arr[-1].split('$')
                        mar.append([tmp_tar[j][0],i[7],arr1[-1],arr1[-2],arr2[-1],arr2[-2],sost])
        for i in range(len(mar)):
            if mar[i][3] == '0':
                s2[mar[i][2]] = int(s2[mar[i][2]]) - int(mar[i][0])
            else:
                s[int(mar[i][3]) - 1][1] = int(s[int(mar[i][3]) - 1][1]) - int(mar[i][0])

            if mar[i][5] == '0':
                s2[mar[i][4]] = int(s2[mar[i][4]]) + int(mar[i][0])
            else:
                s[int(mar[i][5]) - 1][1] = int(s[int(mar[i][5]) - 1][1]) + int(mar[i][0])

            if mar[i][6] == 'закрыта': #в цеху
                if mar[i][3] == '0':
                    s32[mar[i][2]] = int(s32[mar[i][2]]) - int(mar[i][0])
                else:
                    if int(s5[int(mar[i][3]) - 1][1]) - int(mar[i][0]) < 0:
                        delt = int(s5[int(mar[i][3]) - 1][1]) - int(mar[i][0])
                        s5[int(mar[i][3]) - 1][1] = 0
                        if int(s4[int(mar[i][3]) - 1][1]) + delt < 0:
                            delt2 = int(s4[int(mar[i][3]) - 1][1]) + delt
                            s4[int(mar[i][3]) - 1][1] = 0
                            s3[int(mar[i][3]) - 1][1] = int(s3[int(mar[i][3]) - 1][1]) + delt2
                        else:
                            s4[int(mar[i][3]) - 1][1] = int(s4[int(mar[i][3]) - 1][1]) + delt
                    else:
                        s5[int(mar[i][3]) - 1][1] = int(s5[int(mar[i][3]) - 1][1]) - int(mar[i][0])

                if mar[i][5] == '0':
                    s32[mar[i][4]] = int(s32[mar[i][4]]) + int(mar[i][0])
                else:
                    s3[int(mar[i][5]) - 1][1] = int(s3[int(mar[i][5]) - 1][1]) + int(mar[i][0])
                    #s4[int(mar[i][3]) - 1][1] = int(s4[int(mar[i][3]) - 1][1]) - int(mar[i][0])


            if mar[i][6] == 'выдано': #в работе
                if mar[i][3] == '0':
                    s32[mar[i][2]] = int(s32[mar[i][2]]) - int(mar[i][0])
                else:
                    if int(s3[int(mar[i][3]) - 1][1]) - int(mar[i][0]) < 0:
                        delt = int(s3[int(mar[i][3]) - 1][1]) - int(mar[i][0])
                        s3[int(mar[i][3]) - 1][1] = 0
                        s4[int(mar[i][3]) - 1][1] = int(s4[int(mar[i][3]) - 1][1]) + delt
                    else:
                        s3[int(mar[i][3]) - 1][1] = int(s3[int(mar[i][3]) - 1][1]) - int(mar[i][0])

                s4[int(mar[i][5]) - 1][1] = int(s4[int(mar[i][5]) - 1][1]) + int(mar[i][0])


                #s5[int(mar[i][3]) - 1][1] = int(s5[int(mar[i][3]) - 1][1]) - int(mar[i][0])
            if mar[i][6] == 'открыта': #забрана
                if mar[i][3] == '0':
                    s52[mar[i][2]] = int(s52[mar[i][2]]) - int(mar[i][0])
                else:
                    s4[int(mar[i][3]) - 1][1] = int(s4[int(mar[i][3]) - 1][1]) - int(mar[i][0])

                if mar[i][5] == '0':
                    s52[mar[i][4]] = int(s52[mar[i][4]]) + int(mar[i][0])
                else:
                    s5[int(mar[i][5]) - 1][1] = int(s5[int(mar[i][5]) - 1][1]) + int(mar[i][0])

                    #s3[int(mar[i][3]) - 1][1] = int(s3[int(mar[i][3]) - 1][1]) - int(mar[i][0])

        spis_mk = F.otkr_f(F.scfg('bd_mk') + os.sep + mk + '.txt', separ="|")

        for i in range(1, len(spis_mk)):
            if spis_mk[i][6] == id:
                koef = 0
                for j in range(12, len(spis_mk[i]), 4):
                    if spis_mk[i][j - 1] == '':
                        koef += 1
                    if spis_mk[i][j] != '':
                        if 'Полный' in spis_mk[i][j + 1]:
                            if s4[int((j-12-koef*4)/4)][1] > 0:
                                s6[int((j-12-koef*4)/4)][1] = spis_mk[i][2]
                        else:
                            summ_kol = 0
                            sp_nar_mar = spis_mk[i][j + 1].split('$')
                            for k in sp_nar_mar:
                                if "Завершен" in k:
                                    n_nar = k.split(' Завершен')[0].strip()
                                    kol = F.naiti_v_spis_1_1(sp_nar,0,n_nar,12)
                                    summ_kol += int(kol)
                            s6[int((j-12-koef*4)/4)][1] = summ_kol

        if rab_c == '':
            return [s,s2,s3,s32,s4,s5,s52,s6] # По маршруту, вне маршрута, s_bez_tar_po_mar, s_bez_tar_vne_mar, Выдано, s_v_tare_po_mar, s_v_tare_vne_mar,s_pererabot_k_perem
        else:
            return int(s[rab_c][1])

    def sost_rabot_po_det_rc_pornom(self,sp,rc,id,kol):
        for i in range(len(sp)):
            if sp[i][6] == id:
                por_nom = 0
                for j in range(12,len(sp[i]),4):
                    if sp[i][j-1] != '':
                        por_nom+=1
                    if sp[0][j-1] == rc and por_nom == kol:
                        if 'Полный' in sp[i][j+1]:
                            return True
                        return False


    def obn_rc_potok(self):
        tabl_rc_potok = self.ui.table_bd_rab_c_potok
        spis = F.otkr_f(F.tcfg('bd_rab_c'),separ='|')
        tabl_det = self.ui.table_sod_mk
        flag = False
        tabl_mk = self.ui.table_bd_mk
        tabl_det = self.ui.table_sod_mk
        spisok = F.spisok_iz_wtabl(tabl_det, '', shapka=True)
        bd_arh_tar = F.otkr_f(F.tcfg('arh_tar'), separ='|')
        if tabl_det.hasFocus() == True:
            s = self.tek_pol_det(spisok[tabl_det.currentRow()+1][11],tabl_mk.item(tabl_mk.currentRow(), 0).text(),'',bd_arh_tar)[1]
            for i in range(len(spis)):
                spis[i].insert(2,s[spis[i][0]])
        F.zapoln_wtabl(self, spis, tabl_rc_potok, 0, 0, '', '', isp_shapka=False, separ='')
        if tabl_rc_potok.columnCount()>3:
            for cell in range(0,tabl_rc_potok.rowCount()):
                if tabl_rc_potok.item(cell,2).text() != '0':
                    F.ust_color_wtab(tabl_rc_potok,cell,2, 254, 115, 0)
                    flag = True
        return flag

    def obn_potok(self):
        tabl_potok = self.ui.table_potok
        tabl_mk = self.ui.table_bd_mk


        if tabl_mk.currentRow() == -1:
            return
        bd_arh = F.otkr_f(F.tcfg('arh_tar'), separ='|')
        mom_mk = tabl_mk.item(tabl_mk.currentRow(), 0).text()

        s= []
        s.append(bd_arh[0])
        for i in bd_arh:
            if i[6] == 'открыта' and i[3] == mom_mk:
                s.append(i)
        F.zapoln_wtabl(self,s,tabl_potok,0,0,"","",200,True,'',30)

    def dost_ostatok_det(self):
        tabl_mk = self.ui.table_bd_mk
        tabl_det = self.ui.table_sod_mk
        spisok = F.spisok_iz_wtabl(tabl_det, '', shapka=True)
        bd_arh_tar = F.otkr_f(F.tcfg('arh_tar'), separ='|')
        for i in range(1,len(spisok)-1):
            s = self.tek_pol_det(spisok[i][11],tabl_mk.item(tabl_mk.currentRow(), 0).text(),'',bd_arh_tar)
            for j in range(16,len(spisok[i])):
                if tabl_det.item(i-1,j).text().count('_'):
                    arr = tabl_det.item(i-1,j).text().split('_')
                    tmp = arr[0]
                else:
                    tmp = tabl_det.item(i-1,j).text()
                tabl_det.item(i-1,j).setText(tmp + '_' + str(s[0][j-16][1]))


    def tek_tara_rc(self,spisok,nom_mk):
        bd_arh_tar = F.otkr_f(F.tcfg('arh_tar'),separ='|')
        for i in bd_arh_tar:
            if i[3] == nom_mk:
                sost = i[6]
                nom = i[0]
                marsh = i[9]
                det_tmp = F.otkr_f(F.scfg('bd_tara') + os.sep + nom + '.txt',separ='|')
                for i in range(0,len(det_tmp)):
                    for j in range(1,len(spisok)):
                        if det_tmp[i][1].strip() == spisok[j][3].strip() and det_tmp[i][2].strip() == spisok[j][4].strip():
                            polojenie = self.tek_rc(marsh)
                            spisok[j][7] = polojenie
                            if sost == 'выдано':
                                spisok[j][6] = 'выдано'
                            if sost == 'открыта':
                                spisok[j][6] = nom
                            if sost == 'закрыта':
                                spisok[j][6] = ''
                            break
        return spisok

    def tek_rc(self,marsh):
        arr = marsh.split('-->')
        tmp = arr[-1].split('$')
        return tmp[-1]

    def zagruz_det_iz_mk(self):
        tabl_mk = self.ui.table_bd_mk
        tabl_det = self.ui.table_sod_mk

        tabl_det.sortByColumn(-1,0)
        if tabl_mk.currentRow() == -1:
            return
        nom = tabl_mk.item(tabl_mk.currentRow(), 0).text()
        s = self.spis_det_s_ih_marsh(nom)

        s[0][0] = 'Факт.кол.'
        s[0][1] = 'Дост.кол.'
        s[0][2] = 'Выдано'
        s[0][6] = 'Тара'
        s[0][7] = 'Тек.РЦ'

        ed_kol = {0}
        F.zapoln_wtabl(self, s, tabl_det, 0, ed_kol,'', '', isp_shapka=True, separ='',ogr_maxshir_kol= 150)

        self.dost_ostatok_det()

        tabl_det.setColumnHidden(6,True)
        tabl_det.setColumnHidden(7,True)
        tabl_det.setColumnHidden(1,True)
        tabl_det.setColumnHidden(2,True)

        tabl_det.setSelectionMode(1)
        self.oform_cveta()
        self.obn_potok()

    def oform_cveta(self):
        tabl_det = self.ui.table_sod_mk
        tabl_mk = self.ui.table_bd_mk
        nom_mk = tabl_mk.item(tabl_mk.currentRow(), 0).text()
        s = F.spisok_iz_wtabl(tabl_det,'',True)
        for i in range(1,len(s)):
            for j in range(1, len(s[i])):
                F.ust_color_wtab(tabl_det,i-1,j,233,233,233)   #все в серый
            #if i > 0 and int(s[i][2]) == int(s[i][5]):
             #   F.dob_color_wtab(tabl_det, i-1, 0, 16, 16, 16) #выдано и кол-во совпадают серым
        spis_mk = F.otkr_f(F.scfg('bd_mk') + os.sep + nom_mk + '.txt', separ="|")
        for i in range(1,len(spis_mk)):
            koef = 0
            for j in range(12,len(spis_mk[i]),4):
                if spis_mk[i][j-1] == '':
                    koef+=1
                if spis_mk[i][j] != '':
                    if '(полный компл.)' in spis_mk[i][j]:
                        F.ust_color_wtab(tabl_det, i - 1, 16 -koef+ (j- 12)/4 , 0, 254, 0)
                        if 'Полный' in spis_mk[i][j+1]:
                            F.ust_color_wtab(tabl_det, i - 1, 17 - koef + (j - 12) / 4, 254, 254, 254)
                    else:
                        F.ust_color_wtab(tabl_det, i - 1, 16 - koef+ (j - 12) / 4, 254, 115, 0)
        tabl_det.setColumnHidden(11,True)



app = QtWidgets.QApplication([])

myappid = 'Powerz.BAG.SustControlWork.0.0.0'  # !!!

QtWin.setCurrentProcessExplicitAppUserModelID(myappid)
app.setWindowIcon(QtGui.QIcon(os.path.join("icons", "icon.png")))

S = F.scfg('Stile').split(",")
app.setStyle(S[1])

application = mywindow()
application.show()

sys.exit(app.exec())