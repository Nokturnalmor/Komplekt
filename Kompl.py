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
        tabl_det.doubleClicked.connect(self.dblclk_tab_det)

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
        tabl_potok.clicked.connect(self.obn_ob_mar)

        tabl_ob_mar = self.ui.table_obsh_marshrut
        #tabl_potok.setSelectionBehavior(1)
        tabl_ob_mar.setSelectionMode(1)
        F.ust_cvet_videl_tab(tabl_ob_mar)
        tabl_ob_mar.verticalHeader().setVisible(False)
        tabl_ob_mar.horizontalHeader().setVisible(False)

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

    def tab_click(self):
        tab = self.ui.tabWidget
        if tab.currentIndex() == 0:
            self.zagruz_det_iz_mk()

    def dblclk_tab_det(self):
        tabl_det = self.ui.table_sod_mk
        r = tabl_det.currentRow()
        k = tabl_det.currentColumn()
        if k == 6:
            self.sost_tary_iz_mk(r,k)
        if k > 14:
            self.spis_op_iz_mk(r,k)

    def spis_op_iz_mk(self,r,c):
        tabl_det = self.ui.table_sod_mk
        tabl_mk = self.ui.table_bd_mk
        id = tabl_det.item(r,11).text().strip()
        naim = tabl_det.item(r,3).text().strip()
        nn = tabl_det.item(r,4).text().strip()
        nom_mk = tabl_mk.item(tabl_mk.currentRow(), 0).text()
        n_rc = tabl_det.item(r,c).text().strip()
        if F.nalich_file(F.scfg('bd_mk') + os.sep + nom_mk + '.txt') == False:
            showDialog(self, 'Не найти файл ' + F.scfg('bd_mk') + os.sep + nom_mk + '.txt')
            return
        spis_mk = F.otkr_f(F.scfg('bd_mk') + os.sep + nom_mk + '.txt', separ="|")
        for i in range(1,len(spis_mk)):
            if spis_mk[i][6] == id:
                for j in range(11, len(spis_mk[0])):
                    if spis_mk[i][j] != '' and spis_mk[0][j] == n_rc:
                        tmp = spis_mk[i][j].split('$')
                        sp_op = tmp[-1].split(';')
                        if F.nalich_file(F.tcfg('BD_dse')) == False:
                            showDialog(self, 'Не найден BD_dse')
                            return
                        sp_dse = F.otkr_f(F.tcfg('BD_dse'), False, '|')

                        for i in range(0, len(sp_dse)):
                            if sp_dse[i][0] == nn and sp_dse[i][1] == naim:
                                nom_tk = sp_dse[i][2]
                                if nom_tk == '':
                                    showDialog(self, 'Не найден номер ТК')
                                    return
                                break
                        if F.nalich_file(F.scfg('add_docs') + os.sep + nom_tk + '_' + nn + '.txt') == False:
                            showDialog(self, 'Не найден файл ТК')
                            return
                        sp_tk = F.otkr_f(F.scfg('add_docs') + os.sep + nom_tk + '_' + nn + '.txt', False, "|")
                        msgg = ''
                        for o1 in sp_op:
                            msgg += str(o1) + ': '
                            for i in range(11, len(sp_tk)):
                                if sp_tk[i][3].startswith('Т1-' + str(o1).strip()) == True:
                                    if sp_tk[i][20] == '1':
                                        msgg += sp_tk[i][0] + '\n' + ' Tп.з.=' + sp_tk[i][6] + ' Tшт.=' + sp_tk[i][
                                            7] + '\n'
                                    else:
                                        msgg += sp_tk[i][0] + '\n'
                            msgg += '\n'
                        showDialog(self, msgg)


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

    def obn_ob_mar(self):
        tabl_ob_mar = self.ui.table_obsh_marshrut
        tabl_potok = self.ui.table_potok
        nom_mk = tabl_potok.item(tabl_potok.currentRow(), 3).text()
        if F.nalich_file(F.scfg('bd_mk') + os.sep + nom_mk + '.txt') == False:
            showDialog(self, 'Не найти файл ' + F.scfg('bd_mk') + os.sep + nom_mk + '.txt')
            return
        spis_mk = F.otkr_f(F.scfg('bd_mk') + os.sep + nom_mk + '.txt', separ="|")
        ss= []
        s=[]
        for i in range(11, len(spis_mk[0]),4):
            ss.append(spis_mk[0][i])
        s.append(ss)
        F.zapoln_wtabl(self,s,tabl_ob_mar,0,0,"","",100,False,"")



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
            s[0].append('')
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
        tabl_ob_mar = self.ui.table_obsh_marshrut
        tabl_potok = self.ui.table_potok
        tabl_rc_potok = self.ui.table_bd_rab_c_potok
        if tabl_potok.currentRow() == None or tabl_potok.currentRow() == -1:
            showDialog(self, 'Не выбрана тара')
            return
        if tabl_ob_mar.currentRow() == None or tabl_ob_mar.currentRow() == -1:
            showDialog(self, 'Не выбран РЦ')
            return
        nom_tar = tabl_potok.item(tabl_potok.currentRow(), 0).text()
        nom_mk = tabl_potok.item(tabl_potok.currentRow(), 3).text()
        rc = tabl_ob_mar.item(0, tabl_ob_mar.currentColumn()).text()

        if self.vse_det_v_odin_rc(nom_mk,nom_tar,rc) == False:
            showDialog(self, 'Не все детали в таре ' + nom_tar + " имеют общий РЦ :" + rc)
            return
        sp_tar = F.otkr_f(F.tcfg('arh_tar'), separ='|')
        for i in range(0, len(sp_tar)):
            if sp_tar[i][0] == nom_tar:
                if rc != self.tek_rc(sp_tar[i][9]):
                    showDialog(self, 'Нельзя выдать в ' + rc + ", т.к. тара находится в " + self.tek_rc(sp_tar[i][9]))
                    return
                sp_tar[i][6] = 'выдано'
                sp_tar[i][7] = str(tabl_ob_mar.currentColumn())
                break
        self.uchet_komplektacii(sp_tar,nom_tar,nom_mk,rc,tabl_ob_mar.currentColumn())
        F.zap_f(F.tcfg('arh_tar'), sp_tar, '|')
        self.obn_potok()
        showDialog(self, 'Выдача записана успешно')
        return

    def uchet_komplektacii(self,spis_arh_tar,nom_tar,nom_mk,rc,nomer_rc):
        spis_mk = F.otkr_f(F.scfg('bd_mk') + os.sep + nom_mk + '.txt', separ="|")
        spis_tar = F.otkr_f(F.scfg('bd_tara') + os.sep + nom_tar + '.txt',separ='|')

        for i in range(len(spis_tar)):
            for j in range(len(spis_mk)):
                if spis_tar[i][1].strip() == spis_mk[j][0].strip() and spis_tar[i][2].strip() == spis_mk[j][1].strip():
                    kol_vyd = 0
                    for k in range(len(spis_arh_tar)):
                        if spis_arh_tar[k][3] == nom_mk and spis_arh_tar[k][6] == 'выдано' and str(nomer_rc) == spis_arh_tar[k][7]:
                            if self.tek_rc(spis_arh_tar[k][9]) == rc:
                                spis_tar_tmp = F.otkr_f(F.scfg('bd_tara') + os.sep + spis_arh_tar[k][0] + '.txt', separ='|')
                                for l in range(len(spis_tar_tmp)):
                                    if spis_tar_tmp[l][3].strip() == spis_mk[j][6].strip():
                                        kol_vyd+= int(spis_tar_tmp[l][0])
                    '''n = nomer_rc
                    nomer_rc_v_mk = -1
                    for k in range(11, len(spis_mk[j]),4):
                        if spis_mk[j][k] != '':
                            n-=1
                        if n == -1:
                            nomer_rc_v_mk = k+1
                            break
                    if nomer_rc_v_mk == -1:
                        
                        showDialog(self,
                                   'Не найден в МК для ДСЕ ' + spis_mk[j][0].strip() + ' ' + spis_mk[j][1].strip() +
                                    ' порядковый номер РЦ ' + str(nomer_rc))
                        return'''
                    nomer_rc_v_mk = 10 + 4 * (nomer_rc + 1) - 2
                    if kol_vyd == int(spis_mk[j][2]):
                        spis_mk[j][nomer_rc_v_mk] = str(kol_vyd) + ' шт. (полный компл.)'
                    elif kol_vyd < int(spis_mk[j][2]):
                        if kol_vyd//int(spis_mk[j][7]) > 1:
                            spis_mk[j][nomer_rc_v_mk] = str(kol_vyd) + ' шт. (' + str(
                                int(kol_vyd//int(spis_mk[j][7]))) + 'компл.)'
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
                sp_tar[i][9] = sp_tar[i][9] + "-->" + dop
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
        rab_c = tabl_rc.item(tabl_rc.currentRow(), 0).text()
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
                if spis[i][7] != '':
                    if rab_c != spis[i][7]:
                        showDialog(self, 'ДСЕ не находится на ' + rab_c)
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

        sp_tar.append([nom,F.tek_polz(),F.date(2),nom_mk,'',vid_tar,'открыта',"",text_prim.text(),F.tek_polz()+"$"+F.now()+"$"+rab_c])
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
                            if det_tmp[i][3].strip() == spisok[j][11].strip():
                                if sost == 'открыта':
                                    spisok[j][1] = int(spisok[j][1]) - int(det_tmp[i][0])
                                if sost == 'выдано':

                                    spisok[j][2] = int(spisok[j][2]) + int(det_tmp[i][0])
                                break
                    for j in range(1, len(spisok)):
                        if int(spisok[j][2])>int(spisok[j][5]):
                            spisok[j][2] = int(spisok[j][2]) % int(spisok[j][5])
        return spisok

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

        nom = tabl_mk.item(tabl_mk.currentRow(),0).text()
        s = self.spis_det_s_ih_marsh(nom)
        s[0][0] = 'Факт.кол.'
        s[0][1] = 'Дост.кол.'
        s[0][2] = 'Выдано'
        s[0][6] = 'Тара'
        s[0][7] = 'Тек.РЦ'
        nom_mk = tabl_mk.item(tabl_mk.currentRow(), 0).text()
        s = self.dost_ostatok_det(s,nom_mk)
        s = self.tek_tara_rc(s,nom_mk)
        ed_kol = {0}

        F.zapoln_wtabl(self, s, tabl_det, 0, ed_kol,'', '', isp_shapka=True, separ='',ogr_maxshir_kol= 150)




        #tabl_det.resizeColumnsToContents()


        tabl_det.setSelectionMode(1)
        for i in range(0,len(s)-1):
            for j in range(1, len(s[i])):
                F.dob_color_wtab(tabl_det,i,j,16,16,16)
            if i > 0 and int(s[i][2]) == int(s[i][5]):
                F.dob_color_wtab(tabl_det, i-1, 0, 16, 16, 16)
        spis_mk = F.otkr_f(F.scfg('bd_mk') + os.sep + nom_mk + '.txt', separ="|")
        for i in range(1,len(spis_mk)):
            koef = 0
            for j in range(12,len(spis_mk[i]),4):
                if spis_mk[i][j-1] == '':
                    koef+=1
                if spis_mk[i][j] != '':
                    if '(полный компл.)' in spis_mk[i][j]:
                        F.ust_color_wtab(tabl_det, i - 1, 16 -koef+ (j- 12)/4 , 0, 254, 0)
                    else:
                        F.ust_color_wtab(tabl_det, i - 1, 16 -koef+ (j - 12) / 4, 254, 115, 0)
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