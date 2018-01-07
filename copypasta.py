#!/usr/bin/env python
# -*- coding: utf-8 -*-
import os
import sys
import webbrowser

import pandas as pd
import pyperclip
from PyQt5 import QtCore
from PyQt5.QtWidgets import (
    QApplication,
    QGridLayout,
    QTextEdit,
    QLabel,
    QWidget,
    QPushButton,
    QLineEdit,
    QMessageBox)
from pandas import ExcelWriter

DATA_PATH = "./copypasta.xlsx"
TABLE_NAME = "copypasta data"


class MainWindow(QWidget):
    def __init__(self, parent=None):
        super(MainWindow, self).__init__(parent)

        self.label_reference = QLabel("Reference:")
        self.edit_reference = QLineEdit()
        self.edit_reference.setText(pyperclip.paste())

        self.label_excerpt = QLabel("Excerpt:")
        self.edit_excerpt = QTextEdit()
        self.edit_excerpt.setAcceptRichText(False)

        self.button_save = QPushButton("Save")
        self.button_save.clicked.connect(self.save)

        self.timer = QtCore.QTimer(self)
        self.timer.timeout.connect(self.on_timer)
        self.timer.start(100)

        self.mainLayout = QGridLayout()
        # row, col, rowspan, colspan
        self.mainLayout.addWidget(self.label_reference, 0, 0, 1, 1)
        self.mainLayout.addWidget(self.edit_reference, 1, 0, 1, 1)
        self.mainLayout.addWidget(self.label_excerpt, 2, 0, 1, 1)
        self.mainLayout.addWidget(self.edit_excerpt, 3, 0, 1, 1)
        self.mainLayout.addWidget(self.button_save, 4, 0, 1, 1)
        self.setLayout(self.mainLayout)

    def save(self):
        new_df = pd.DataFrame(
            {
                'reference': [self.edit_reference.text()],
                'excerpt': [self.edit_excerpt.toPlainText()],
            }
        )

        if os.path.isfile(DATA_PATH):
            df = pd.read_excel(DATA_PATH, sheet_name=TABLE_NAME)
            df = df.append(new_df)
        else:
            df = new_df

        writer = ExcelWriter(DATA_PATH)
        df.to_excel(writer, TABLE_NAME)
        writer.save()

        buttonReply = QMessageBox.question(
            self,
            "Paste Saved",
            "Paste Saved to %s\nDo you want to open it now?" % DATA_PATH,
            QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel, QMessageBox.Cancel
        )
        if buttonReply == QMessageBox.Yes:
            webbrowser.open(DATA_PATH)

    def on_timer(self):
        if not pyperclip.paste() == self.edit_reference.text():
            self.edit_excerpt.setText(pyperclip.paste())


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
