from DataStorage import DataStorage as DS
import Editor as ed
import Executor
import time
from PySide6 import QtWidgets, QtGui
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (QPushButton,
                               QLabel,
                               QGridLayout,
                               QLineEdit,
                               QFileDialog,
                               QComboBox)


class Window(QtWidgets.QWidget):

    def __init__(self):

        super().__init__()

        self.ds = DS()
        self.setWindowTitle('Inserter')
        self.setFixedSize(360, 250)
        self.setWindowIcon(QtGui.QIcon("C:\\Users\\EDN\\Downloads\\images2.ico"))

        # region Лейблы
        self.label_First = QLabel('№ первого изделия: ')
        self.label_Amount = QLabel('Общее количество изделий: ')
        self.label_PerMSL = QLabel('Количество на МСЛ: ')
        self.label_DeviceName = QLabel('Наименование изделия: ')
        self.label_Contract = QLabel('Номер контракта: ')
        self.label_MasterName = QLabel('Мастер: ')
        # endregion

        # region Текстбоксы
        self.textbox_First = QLineEdit('100')
        self.textbox_Amount = QLineEdit('100')

        self.textbox_PathToTable = QLineEdit("C:\\Users\\EDN\\PycharmProjects\\AutoInsert\\Templates\\table.xlsx")

        # self.textbox_PathToTable = QLineEdit('C:\\Users\\fmm\\Desktop\\Контроль участков\\Новое')

        # self.textbox_PathToTable = (
        #     QLineEdit('\\\\192.168.2.10\\xchg\\Производство\\Работа с МСЛ\\Регистрация МСЛ.xlsx'))

        self.textbox_PerMSL = QLineEdit('10')
        # self.textbox_DeviceName = QLineEdit()
        self.textbox_Contract = QLineEdit("G1021")
        # self.textbox_MasterName = QLineEdit('Женя')

        self.listbox_Masters = QComboBox()
        self.listbox_Masters.addItems(self.ds.masters)

        self.listbox_Devices = QComboBox()
        self.listbox_Devices.addItems(self.ds.devices)

        self.textbox_Status = QLineEdit('Статус: ')
        self.textbox_Status.setReadOnly(True)

        # endregion

        # region Кнопки
        self.button_Insert = QPushButton('Вставить номера')
        self.button_Insert.clicked.connect(self.button_insert_clicked)

        self.button_SelectPathToTable = QPushButton('Путь к таблице')
        self.button_SelectPathToTable.clicked.connect(self.button_PathToTable)

        self.button_Clear = QPushButton('Очистить')
        self.button_Clear.clicked.connect(self.button_Clear_pressed)

        # endregion

        self.UI()

    def UI(self):

        grid = QGridLayout()
        self.setLayout(grid)

        # Организация сетки окна

        # region Лейблы
        grid.addWidget(self.label_First, 0, 0, alignment=Qt.AlignCenter)
        grid.addWidget(self.label_Amount, 1, 0, alignment=Qt.AlignCenter)
        grid.addWidget(self.label_PerMSL, 2, 0, alignment=Qt.AlignCenter)
        grid.addWidget(self.label_DeviceName, 3, 0, alignment=Qt.AlignCenter)
        grid.addWidget(self.label_Contract, 4, 0, alignment=Qt.AlignCenter)
        grid.addWidget(self.label_MasterName, 5, 0, alignment=Qt.AlignCenter)

        # endregion

        # region Текстбоксы
        grid.addWidget(self.textbox_First, 0, 1, alignment=Qt.AlignCenter)
        grid.addWidget(self.textbox_Amount, 1, 1, alignment=Qt.AlignCenter)
        grid.addWidget(self.textbox_PerMSL, 2, 1, alignment=Qt.AlignCenter)
        # grid.addWidget(self.textbox_PathToTable, 9, 0, 1, 2)
        grid.addWidget(self.listbox_Devices, 3, 1, alignment=Qt.AlignCenter)
        grid.addWidget(self.textbox_Contract, 4, 1, alignment=Qt.AlignCenter)
        grid.addWidget(self.listbox_Masters, 5, 1, alignment=Qt.AlignCenter)
        grid.addWidget(self.textbox_Status, 12, 0, 1, 2)

        # endregion

        # region Кнопки
        grid.addWidget(self.button_Insert, 11, 0, 1, 2)
        # grid.addWidget(self.button_SelectPathToTable, 10, 0)
        # grid.addWidget(self.button_Clear, 10, 1)
        # endregion

    def parse_textboxes(self):
        """
        Забирает из текстбоксов ввведенные значения и
        складывает данные в соответствующие поля в классе DataStorage
        """
        self.ds.first = int(self.textbox_First.text())
        self.ds.amount = int(self.textbox_Amount.text())
        self.ds.per_one_msl = int(self.textbox_PerMSL.text())
        self.ds.contract = self.textbox_Contract.text()
        self.ds.device_name = self.listbox_Devices.currentText()
        self.ds.master_name = self.listbox_Masters.currentText()

    def button_PathToTable(self):
        """
        Открывает диалоговое окно с файловой системой для выбора таблицы
        """
        # self.textbox_PathToTable.clear()
        selected_table: str = ''
        try:
            selected_table = QFileDialog.getOpenFileNames(self, filter='*.xlsx')[0][0]
        except:
            pass

        selected_table = ed.path_for_win(selected_table)
        self.ds.table = selected_table
        self.textbox_PathToTable.setText(selected_table)

    def button_Clear_pressed(self):
        """Очищает текстбокс PathToTable"""
        self.textbox_PathToTable.clear()

    def button_insert_clicked(self):

        try:
            self.parse_textboxes()
        except ValueError:
            self.textbox_Status.setStyleSheet('QLineEdit {color: red;}')
            self.textbox_Status.setText('Статус: ' + 'заполните все поля данных!')

        try:
            begin = time.time()

            Executor.execute(self.ds, self.ds.table_N)

            end = time.time()
            # self.parse_textboxes()
            # Executor.execute(self.ds, self.ds.table_M)
            # self.parse_textboxes()
            # Executor.execute(self.ds, self.ds.table_S)

            self.textbox_Status.setStyleSheet('QLineEdit {color: green;}')
            self.textbox_Status.setText('Статус: ' + 'выполнено успешно')


            print("Время выполнения:", end - begin)

        except PermissionError:
            self.textbox_Status.setStyleSheet('QLineEdit {color: red;}')
            self.textbox_Status.setText('Статус: ' + 'закройте таблицу или проверьте к ней доступ ')


