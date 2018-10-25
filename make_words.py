from PyQt5 import QtCore, QtGui, QtWidgets
from ui.createNew import Ui_CreateNew
from ui.testings import Ui_MainWindow
from ui.printPDF import Ui_printWindow
from words import getUnits, WordsPrint
import os, time
from sqlite_tools import EasySqlite


class WinPrint(QtWidgets.QDialog, Ui_printWindow):
    def __init__(self):
        super(WinPrint, self).__init__()
        self.setupUi(self)
        self.setWindowModality(QtCore.Qt.ApplicationModal)
        self.btnOpenPDF.setEnabled(False)
        self._t_id = -1
        self.__filename = ''
        self.btnCreate.clicked.connect(self.create_pdf)
        self.btnOpenPDF.clicked.connect(self.open_pdf)

    def loadTest(self):
        _id = mainw.listTesting.currentRow()
        self._t_id = mainw._test_id[_id]
        self.btnOpenPDF.setEnabled(False)
        self.show()

    def create_pdf(self):
        _wintitle = self.windowTitle()
        self.setWindowTitle('开始生成PDF文件，请稍候...')
        page_conf = {
            'pagesize': self.cmbPageSize.currentText(),
            'columns': self.spbColumns.value(),
            'col_padding': self.spbColMargin.value(),
            'left': self.spbLeftPadding.value(),
            'top': self.spbTopPadding.value(),
            'right': self.spbRightPadding.value(),
            'bottom': self.spbBottomPadding.value(),
            'line_height': self.lineHeight.value()
        }
        pdf = WordsPrint(page_conf)
        done = pdf.loadTest(self._t_id)
        if not done:
            QtWidgets.QMessageBox.warning(self, '警告', '没有符合条件的练习内容,无法生成PDF练习文件!')
            return
        else:
            _t_unit = done[0]
            _isEng = done[1]
            _onlyError = done[2]
        if _isEng:
            pdf_content_type = 'Eng'
        else:
            pdf_content_type = 'Chs'
        if _onlyError:
            pdf_content_type = pdf_content_type + '_onlyError'
        thisfile_name = pdf_content_type + '_' + _t_unit + '.pdf'
        file_name, _ = QtWidgets.QFileDialog.getSaveFileName(None, '保存PDF', os.getcwd() + '/' + thisfile_name, 'PDF文件(*.PDF)')
        if len(file_name) == 0:
            self.setWindowTitle('PDF文件已生成已取消')
            return

        pdf.printPDF(file_name, _isEng, self.createAnswer.isChecked(), _onlyError, self.d_page.isChecked(),
                     msg_obj=self)
        self.setWindowTitle('练习文件%s已生成，点击下方“打开”按钮立即查看!' % file_name)
        self.__filename = file_name
        self.btnOpenPDF.setEnabled(True)

    def showMessage(self, msg):
        self.setWindowTitle(msg)

    def open_pdf(self):
        done = os.system('evince %s' % self.__filename)
        if done != 0:
            QMessageBox.critical(win, "警告", "%s文件不存在，打开失败！" % self.__filename)

class WinCreate(QtWidgets.QDialog, Ui_CreateNew):
    def __init__(self):
        super(WinCreate, self).__init__()
        self.setupUi(self)
        self.setWindowModality(QtCore.Qt.ApplicationModal)
        self.btnOpenPDF.setDisabled(True)
        self.btnCreate.clicked.connect(self.create_pdf)
        self.btnOpenPDF.clicked.connect(self.open_pdf)
        self.__filename = ''
        __unit_list = getUnits()
        self.startUnit.clear()
        self.endUnit.clear()
        for u in __unit_list:
            self.startUnit.addItem(u)
            self.endUnit.addItem(u)
        self.isSave = False

    def open_pdf(self):
        done = os.system('evince %s' % self.__filename)
        if done != 0:
            QMessageBox.critical(win, "警告", "%s文件不存在，打开失败！" % self.__filename)

    def create_pdf(self):
        _wintitle = self.windowTitle()
        self.setWindowTitle('开始生成PDF文件，请稍候...')
        page_conf = {
            'pagesize': self.cmbPageSize.currentText(),
            'columns': self.spbColumns.value(),
            'col_padding': self.spbColMargin.value(),
            'left': self.spbLeftPadding.value(),
            'top': self.spbTopPadding.value(),
            'right': self.spbRightPadding.value(),
            'bottom': self.spbBottomPadding.value(),
            'line_height': self.lineHeight.value()
        }
        content = []
        if self.selectWords.isChecked():
            content.append("'单词'")
        if self.selectWordGroup.isChecked():
            content.append("'词组'")
        content_str = ", ".join(content)
        content_conf = {
            'list': content_str,
            'isEnglish': self.printEn.isChecked(),
            'answer': self.createAnswer.isChecked(),
            'onlyError': self.onlyError.isChecked()
        }
        s_u = self.startUnit.currentText()
        e_u = self.endUnit.currentText()
        start_u = min(s_u, e_u)
        end_u = max(s_u, e_u)
        pdf = WordsPrint(page_conf)
        pdf.setRange(start_u, end_u, content_conf, is_Rand=self.isRandom.isChecked())
        if self.printEn.isChecked():
            pdf_content_type = 'Eng'
            ch_content_type = '看英文写中文'
        else:
            pdf_content_type = 'Chs'
            ch_content_type = '看中文写英文'
        if self.onlyError.isChecked():
            pdf_content_type = pdf_content_type + '_onlyError'
        thisfile_name = pdf_content_type + '_' + start_u + '-' + end_u + '.pdf'
        ch_this_testing = ch_content_type + '(' + start_u + '-' + end_u + ')'
        ch_content = ','.join(content)
        if self.onlyError.isChecked():
            ch_content = ch_content + '  [仅错题]'
        ch_content = ch_content.strip()
        test_list, test_records = pdf.getWords
        if len(test_list) < 1:
            QtWidgets.QMessageBox.warning(self, '警告', '没有符合条件的练习内容,\r请重新设定条件再生成!')
            return

        file_name, _ = QtWidgets.QFileDialog.getSaveFileName(None, '保存PDF', os.getcwd() + '/' + thisfile_name, 'PDF文件(*.PDF)')
        if len(file_name) == 0:
            self.setWindowTitle('PDF文件已生成已取消')
            return

        ch_test_list = str(test_list)[1:-1]
        db = EasySqlite(dbfile)
        db.execute(sql='insert into testings (卷名, 生成时间, 类型, 内容, 题列表) values (?, ?, ?, ?, ?)',
                   args=[ch_this_testing, time.strftime('%Y-%m-%d %H:%M:%S'), ch_content_type, ch_content, ch_test_list])

        pdf.printPDF(file_name, content_conf['isEnglish'], content_conf['answer'], content_conf['onlyError'], self.d_page.isChecked(),
                     msg_obj=self)
        self.setWindowTitle('练习文件%s已生成，点击下方“打开”按钮立即查看!' % file_name)
        self.__filename = file_name
        self.btnOpenPDF.setEnabled(True)

    def showMessage(self, msg):
        self.setWindowTitle(msg)

    def closeEvent(self, a0: QtGui.QCloseEvent):
        self.isSave = self.btnOpenPDF.isEnabled()
        # mainw.add_new()
        mainw.load_test_list()

    def new_testing(self):
        self.btnOpenPDF.setEnabled(False)
        self.show()


class MainWindow(QtWidgets.QMainWindow, Ui_MainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()
        self.setupUi(self)
        self.setWindowTitle("高中英语单词(词组)练习生成器-v1.21")
        for i in range(0, self.tabWords.columnCount()):
            self.tabWords.horizontalHeader().setSectionResizeMode(i, QtWidgets.QHeaderView.ResizeToContents)
        self.tabWords.horizontalHeader().setSectionResizeMode(3, QtWidgets.QHeaderView.Stretch)
        self._test_id = []
        self.load_test_list()
        self.btnCreate.setEnabled(False)
        self.btnDone.setVisible(False)
        self.btnCorrect.setEnabled(False)
        self.listTesting.itemSelectionChanged.connect(self.load_test)
        self.dispAll.clicked.connect(self.load_test_list)
        # self.tabWords.cellDoubleClicked.connect(self.correct)
        self.btnCorrect.clicked.connect(self.to_correct)
        self.btnDone.clicked.connect(self.to_Done)
        self.btnCancelCorrect.clicked.connect(self.cancel_correct)
        self.tabWords.setToolTip('双击某行进行对错批改')


    def load_test_list(self):
        _list_range = ' where 批改时间 is Null ' if not self.dispAll.isChecked() else ''
        sql = f'select ID, 卷名 from testings {_list_range} order by 生成时间 desc'
        db = EasySqlite(dbfile)
        result = db.execute(sql)
        self.listTesting.clear()
        self._test_id = []
        for row in result:
            item = row['卷名']+'_'+str(row['ID'])
            self._test_id.append(row['ID'])
            self.listTesting.addItem(item)
        if self.listTesting.currentRow() < 0:
            self.btnCorrect.setEnabled(False)

    def load_test(self):
        db = EasySqlite(dbfile)
        sql = 'select * from testings where ID=%d' % self._test_id[self.listTesting.currentRow()]
        t_list_id = db.execute(sql)
        t_list = t_list_id[0]

        self.t_name.setText(t_list['卷名'])
        self.t_CreateTime.setText(t_list['生成时间'])
        self.t_type.setText(t_list['类型'])
        self.t_content.setText(t_list['内容'].replace("'", ""))
        self.t_LastModi.setText(t_list['批改时间'])
        _corrected_str=t_list['已标注'].strip()
        if len(_corrected_str) < 1:
            _corrected_list=[]
        else:
            _corrected_list = _corrected_str.split(',')
        self.corrected = []
        for x in _corrected_list:
            self.corrected.append(int(x))
        _isEnglish = '英文' if t_list['类型'] == '看英文写中文' else '中文'
        sql = 'select * from words where id in (%s)' % t_list['题列表']
        result = db.execute(sql)
        # print(_isEnglish)
        rows = {}
        # print(result)
        for row in result:
            rows[row['ID']] = row
        i = 0
        self.tabWords.setColumnCount(9)
        self.tabWords.setColumnHidden(8, True)
        for s_id in t_list['题列表'].split(','):
            row = rows[int(s_id)]
            if t_list['批改时间'] is None:
                no = row['ID'] in self.corrected
                _red = QtGui.QColor(255, 0, 0, 255)
                self.tabWords.setRowCount(i+1)
                item = QtWidgets.QTableWidgetItem(str(row['ID']))
                if no:
                    item.setForeground(_red)
                self.tabWords.setItem(i, 0, item)
                item = QtWidgets.QTableWidgetItem(row['类型'])
                if no:
                    item.setForeground(_red)
                self.tabWords.setItem(i, 1, item)
                item = QtWidgets.QTableWidgetItem(row['单词'] + ('(' + row['词性'] + ')' if row['类型']=='单词' else ''))
                if no:
                    item.setForeground(_red)
                self.tabWords.setItem(i, 2, item)
                item = QtWidgets.QTableWidgetItem(str(row['解析']))
                if no:
                    item.setForeground(_red)
                self.tabWords.setItem(i, 3, item)
                if row['ID'] in self.corrected:
                    item = QtWidgets.QTableWidgetItem('错')
                    if no:
                        item.setForeground(_red)
                    self.tabWords.setItem(i, 4, item)
                    item = QtWidgets.QTableWidgetItem(str(row[_isEnglish + '错次']+1))
                    if no:
                        item.setForeground(_red)
                    self.tabWords.setItem(i, 5, item)

                    item = QtWidgets.QTableWidgetItem(str(row[_isEnglish + '练次']+1))
                    if no:
                        item.setForeground(_red)
                    self.tabWords.setItem(i, 6, item)
                    item = QtWidgets.QTableWidgetItem(
                        '{:.1f}'.format((row[_isEnglish + '错次']+1) / (row[_isEnglish + '练次']+1) * 100))
                    if no:
                        item.setForeground(_red)
                    self.tabWords.setItem(i, 7, item)
                    self.tabWords.selectRow(i)

                else:
                    item = QtWidgets.QTableWidgetItem('对')
                    self.tabWords.setItem(i, 4, item)
                    item = QtWidgets.QTableWidgetItem(str(row[_isEnglish+'错次']))
                    self.tabWords.setItem(i, 5, item)
                    item = QtWidgets.QTableWidgetItem(str(row[_isEnglish + '练次']+1))
                    self.tabWords.setItem(i, 6, item)
                    item = QtWidgets.QTableWidgetItem('{:.1f}'.format((row[_isEnglish + '错次']) / (row[_isEnglish + '练次']+1) * 100))
                    self.tabWords.setItem(i, 7, item)
                item = QtWidgets.QTableWidgetItem('N')
                self.tabWords.setItem(i, 8, item)
            else:
                self.tabWords.setRowCount(i + 1)
                item = QtWidgets.QTableWidgetItem(str(row['ID']))
                self.tabWords.setItem(i, 0, item)
                item = QtWidgets.QTableWidgetItem(row['类型'])
                self.tabWords.setItem(i, 1, item)
                item = QtWidgets.QTableWidgetItem(row['单词'] + ('(' + row['词性'] + ')' if row['类型'] == '单词' else ''))
                self.tabWords.setItem(i, 2, item)
                item = QtWidgets.QTableWidgetItem(str(row['解析']))
                self.tabWords.setItem(i, 3, item)
                if row['ID'] in self.corrected:
                    item = QtWidgets.QTableWidgetItem('错')
                else:
                    item = QtWidgets.QTableWidgetItem('对')
                self.tabWords.setItem(i, 4, item)
                item = QtWidgets.QTableWidgetItem(str(row[_isEnglish + '错次']))
                self.tabWords.setItem(i, 5, item)
                item = QtWidgets.QTableWidgetItem(str(row[_isEnglish + '练次']))
                self.tabWords.setItem(i, 6, item)
                item = QtWidgets.QTableWidgetItem('{:.1f}'.format((row[_isEnglish + '错次']) / (row[_isEnglish + '练次']) * 100))
                self.tabWords.setItem(i, 7, item)
                item = QtWidgets.QTableWidgetItem('N')
                self.tabWords.setItem(i, 8, item)

            i = i + 1
        self.btnCreate.setEnabled(i > 0)
        if t_list['批改时间'] is None:
            self.btnCorrect.setEnabled(True)
        else:
            self.btnCorrect.setEnabled(False)
            self.btnDone.setVisible(False)
            try:
                self.tabWords.cellDoubleClicked.disconnect(self.correct)
            except:
                pass

    def to_Done(self):
        _ok = QtWidgets.QMessageBox.question(self, '保存提示', '保存将使练习次数增加一次，并重新计算错误率！\r确认已完成本练习批改（此操作不可撤消）？',
                                             buttons=QtWidgets.QMessageBox.Ok | QtWidgets.QMessageBox.No,
                                             defaultButton=QtWidgets.QMessageBox.No)
        if _ok == QtWidgets.QMessageBox.Ok:
            self.btnCorrect.setText('批改')
            self.btnCorrect.setToolTip('进入批改状态')
            self.btnCorrect.setEnabled(False)
            self.listTesting.setEnabled(True)
            self.btnCancelCorrect.setEnabled(False)
            self.tabWords.cellDoubleClicked.disconnect(self.correct)
            self.btnNew.setEnabled(True)
            self.dispAll.setEnabled(True)
            self._save_correct()
            self._correct_done()
            self.load_test_list()

    def to_correct(self):
        if self.btnCorrect.text() == '批改':
            self.btnNew.setEnabled(False)
            self.listTesting.setEnabled(False)
            self.btnCorrect.setText('保存')
            self.btnCorrect.setToolTip('只要不结束批改，保存的批改将在下次批改时自动读入，以继续批改！')
            self.btnCancelCorrect.setEnabled(True)
            self.btnDone.setVisible(True)
            self.dispAll.setEnabled(False)
            self.load_test()
            self.tabWords.cellDoubleClicked.connect(self.correct)
        else:
            self.btnCorrect.setText('批改')
            self.btnCorrect.setToolTip('进入批改状态')
            self.listTesting.setEnabled(True)
            self.btnCancelCorrect.setEnabled(False)
            self.tabWords.cellDoubleClicked.disconnect(self.correct)
            self.btnNew.setEnabled(True)
            self.dispAll.setEnabled(True)
            self.btnDone.setVisible(False)
            self._save_correct()

    def _save_correct(self):
        _clist = str(self.corrected)
        _this_id = self._test_id[self.listTesting.currentRow()]
        sql = f"update testings set 已标注='{_clist[1:-1]}' where ID={_this_id:d}"
        db = EasySqlite(dbfile)
        db.execute(sql)

    def _correct_done(self):
        _isEnglisth = '英文' if self.t_type.text() == '看英文写中文' else '中文'
        db = EasySqlite(dbfile)
        sql = f"update testings set 批改时间=datetime('now') where ID={self._test_id[self.listTesting.currentRow()]:d}"
        db.execute(sql)
        for i in range(0, self.tabWords.rowCount()):
            sql = 'update words set %s错次=%d, %s练次=%d, %s错率=%f where ID=%d' % (
                _isEnglisth, int(self.tabWords.item(i, 5).text()),
                _isEnglisth, int(self.tabWords.item(i, 6).text()),
                _isEnglisth, float(self.tabWords.item(i, 7).text()),
                int(self.tabWords.item(i, 0).text()))
            db.execute(sql, commit=True)
        self.load_test_list()



    def cancel_correct(self):
        self.listTesting.setEnabled(True)
        self.btnCancelCorrect.setEnabled(False)
        self.btnCorrect.setText('批改')
        try:
            self.tabWords.cellDoubleClicked.disconnect(self.correct)
        except:
            pass
        self.btnNew.setEnabled(True)
        self.dispAll.setEnabled(True)
        self.btnDone.setVisible(False)
        self.load_test()



    def correct(self):
        this_row = self.tabWords.currentRow()
        result = self.tabWords.item(this_row, 4)
        if result.text() == '对':
            result.setText('错')
            this_color =QtGui.QColor(255, 0, 0, 255)
            error_count = int(self.tabWords.item(this_row, 5).text()) + 1
            self.tabWords.item(this_row, 5).setText(str(error_count))
            self.tabWords.item(this_row, 8).setText('Y')
            self.corrected.append(int(self.tabWords.item(this_row, 0).text()))
        else:
            result.setText('对')
            this_color = QtGui.QColor(0, 0, 0, 255)
            error_count = int(self.tabWords.item(this_row, 5).text()) - 1
            self.tabWords.item(this_row, 5).setText(str(error_count))
            self.tabWords.item(this_row, 8).setText('N')
            self.corrected.remove(int(self.tabWords.item(this_row, 0).text()))
        self.tabWords.item(this_row, 7).setText('{:.1f}'.format(int(self.tabWords.item(this_row, 5).text()) / (int(self.tabWords.item(this_row, 6).text())) * 100))
        for c in range(self.tabWords.columnCount()):
            self.tabWords.item(this_row, c).setForeground(this_color)




if __name__=="__main__":
    import sys, tarfile
    s_dbfile = r'/opt/high-school-words/high_school_words_data.tar.gz'
    # dbfile = os.getcwd() + '/high_school_words_data/words.db'
    home_path = os.getenv('~')  # 这种方式获取用户主目录可以兼容windows
    dbfile = os.path.expanduser('~') +'/high_school_words_data/words.db'
    if not os.path.exists(dbfile):
        t = tarfile.open(s_dbfile)
        t.extractall(path=os.path.expanduser('~'))

    app = QtWidgets.QApplication(sys.argv)
    mainw = MainWindow()
    if not os.path.exists(dbfile):
        QtWidgets.QMessageBox.critical(mainw, "警告", "%s数据文件不存在，无法继续运行！" % dbfile)
        exit(1)
    creatw = WinCreate()
    printw = WinPrint()
    mainw.btnNew.clicked.connect(creatw.new_testing)
    mainw.btnCreate.clicked.connect(printw.loadTest)
    x = (app.desktop().width() - mainw.width()) / 2
    y = (app.desktop().height() - mainw.height()) / 2
    mainw.move(x, y)

    mainw.show()
    sys.exit(app.exec_())
