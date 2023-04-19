import os
import sys
import xlsxwriter

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import *


def read_directory(directory):
    file_list = []

    for file in os.listdir(directory):
        file_list.append(file)

    return file_list


def write_to_excel(workbook, worksheet, file_list, column):
    # Create a cell format with red text color
    for (index, file) in enumerate(file_list):
        location = f"{column}{index + 1}"
        if file.background() == Qt.red:
            red_format = workbook.add_format({'color': 'red'})
            worksheet.write(location, file.text(), red_format)
        elif file.background() == Qt.green:
            green_format = workbook.add_format({'color': 'green'})
            worksheet.write(location, file.text(), green_format)
        else:
            worksheet.write(location, file.text())


def alert(title, message, icon=QMessageBox.Information):
    msg_box = QMessageBox()
    msg_box.setWindowTitle(title)
    msg_box.setText(message)
    msg_box.setIcon(icon)
    msg_box.setStandardButtons(QMessageBox.Ok)
    msg_box.exec_()


def success_alert(message, title="Success"):
    alert(title, message, QMessageBox.Information)


class DirectoryCompare(QMainWindow):
    def __init__(self):
        super().__init__()

        self.file_list_widget_1 = QListWidget()
        self.file_list_widget_2 = QListWidget()

        self.is_case_sensitive = False

        self.summary_text_1 = QLineEdit()
        self.summary_text_1.setReadOnly(True)
        self.summary_text_2 = QLineEdit()
        self.summary_text_2.setReadOnly(True)

        self.init_ui()

    def center_on_screen(self):
        screen = QDesktopWidget().screenGeometry()

        size = self.geometry()

        x = int((screen.width() - size.width()) / 2)
        y = int((screen.height() - size.height()) / 2)

        self.move(x, y)

    def init_ui(self):
        self.setWindowTitle('File Compare')
        self.setGeometry(100, 100, 800, 600)

        btn_open_directory_1 = QPushButton('Open Directory 1', self)
        btn_open_directory_2 = QPushButton('Open Directory 2', self)
        btn_open_directory_1.clicked.connect(self.open_directory_1)
        btn_open_directory_2.clicked.connect(self.open_directory_2)

        # Create a central widget
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)

        # Options
        option_layout = QHBoxLayout()
        case_sensitive = QCheckBox('Case Sensitive')
        case_sensitive.setChecked(self.is_case_sensitive)
        case_sensitive.stateChanged.connect(self.on_case_sensitive_changed)

        option_layout.addWidget(case_sensitive)

        # Compare button
        btn_compare = QPushButton('Compare', self)
        btn_compare.clicked.connect(self.compare)

        # Export button
        btn_export = QPushButton('Export', self)
        btn_export.clicked.connect(self.export_to_xlsx)

        # Layout

        vbox = QVBoxLayout()
        grid = QGridLayout()
        grid.addWidget(btn_open_directory_1, 0, 0)
        grid.addWidget(btn_open_directory_2, 0, 1)
        grid.addWidget(self.file_list_widget_1, 1, 0)
        grid.addWidget(self.file_list_widget_2, 1, 1)
        grid.addWidget(self.summary_text_1, 2, 0)
        grid.addWidget(self.summary_text_2, 2, 1)
        vbox.addLayout(grid)
        vbox.addLayout(option_layout)
        vbox.addWidget(btn_compare)
        vbox.addWidget(btn_export)

        central_widget.setLayout(vbox)

        self.center_on_screen()

        self.show()

    # compare two directories and coloring each file show the result
    # if file is in both directories, color it green
    # if file is in only one directory, color it red
    def compare(self):
        # get file list of directory 1
        file_list_1 = []
        for index in range(self.file_list_widget_1.count()):
            file_list_1.append(self.file_list_widget_1.item(index).text())

        # get file list of directory 2
        file_list_2 = []
        for index in range(self.file_list_widget_2.count()):
            file_list_2.append(self.file_list_widget_2.item(index).text())

        # compare two file lists
        dir_1_missing, dir_2_missing, find = 0, 0, 0
        if self.is_case_sensitive:
            # case sensitive
            for file in file_list_1:
                if file in file_list_2:
                    # file is in both directories
                    # color it green
                    index = file_list_1.index(file)
                    item = self.file_list_widget_1.item(index)
                    item.setBackground(Qt.green)
                    index = file_list_2.index(file)
                    item = self.file_list_widget_2.item(index)
                    item.setBackground(Qt.green)
                    find += 1
                else:
                    # file is only in directory 1
                    # color it red
                    index = file_list_1.index(file)
                    item = self.file_list_widget_1.item(index)
                    item.setBackground(Qt.red)
                    dir_1_missing += 1
            for file in file_list_2:
                if file not in file_list_1:
                    # file is only in directory 2
                    # color it red
                    index = file_list_2.index(file)
                    item = self.file_list_widget_2.item(index)
                    item.setBackground(Qt.red)
                    dir_2_missing += 1
        else:
            # case insensitive
            for file in file_list_1:
                if file.lower() in [f.lower() for f in file_list_2]:
                    # file is in both directories
                    # color it green
                    index = file_list_1.index(file)
                    item = self.file_list_widget_1.item(index)
                    item.setBackground(Qt.green)
                    index = file_list_2.index(file)
                    item = self.file_list_widget_2.item(index)
                    item.setBackground(Qt.green)
                    find += 1
                else:
                    # file is only in directory 1
                    # color it red
                    index = file_list_1.index(file)
                    item = self.file_list_widget_1.item(index)
                    item.setBackground(Qt.red)
                    dir_1_missing += 1
            for file in file_list_2:
                if file.lower() not in [f.lower() for f in file_list_1]:
                    # file is only in directory
                    # color it red
                    index = file_list_2.index(file)
                    item = self.file_list_widget_2.item(index)
                    item.setBackground(Qt.red)
                    dir_2_missing += 1

        self.summary_text_1.setText('Summary: {} files found, {} files missing'.format(find, dir_1_missing))
        self.summary_text_2.setText('Summary: {} files found, {} files missing'.format(find, dir_2_missing))

        success_alert('Compare successfully', title='Compare')

    def open_directory_1(self):
        directory = QFileDialog.getExistingDirectory(self, 'Select Directory')

        if directory:
            # clean list widget
            self.file_list_widget_1.clear()
            # read file list of directory
            file_list = read_directory(directory)
            # add file_list to list widget
            self.file_list_widget_1.addItems(file_list)

    def open_directory_2(self):
        directory = QFileDialog.getExistingDirectory(self, 'Select Directory')

        if directory:
            self.file_list_widget_2.clear()
            # read file list of directory
            file_list = read_directory(directory)
            # add file_list to list widget
            self.file_list_widget_2.addItems(file_list)

    def on_case_sensitive_changed(self, state):
        if state == Qt.Checked:
            self.is_case_sensitive = True
        else:
            self.is_case_sensitive = False

    # export compare result to excel file
    def export_to_xlsx(self):
        # get file list of directory 1
        file_list_1 = []
        for index in range(self.file_list_widget_1.count()):
            file_list_1.append(self.file_list_widget_1.item(index))

        # get file list of directory 2
        file_list_2 = []
        for index in range(self.file_list_widget_2.count()):
            file_list_2.append(self.file_list_widget_2.item(index))

        # Show the save file dialog
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        save_path, _ = QFileDialog.getSaveFileName(self, 'Save File', '', 'Xlsx Files (*.xlsx)',
                                                   options=options)

        if save_path == '':
            return

        workbook = xlsxwriter.Workbook(save_path if save_path.endswith('.xlsx') else save_path + '.xlsx')
        worksheet = workbook.add_worksheet()

        write_to_excel(workbook, worksheet, file_list_1, 'A')
        write_to_excel(workbook, worksheet, file_list_2, 'B')

        workbook.close()

        success_alert('Export to excel file successfully', title='Export')


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = DirectoryCompare()
    sys.exit(app.exec_())
