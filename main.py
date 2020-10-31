import sys
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import os
from PyQt5 import QtWidgets, QtCore, QtGui
from os.path import basename

import PyPDF2
import smtplib

from PyQt5.QtCore import pyqtSlot, QDate
from PyQt5.QtWidgets import QFileDialog

from ui_email import Ui_EmailWindow


def create_pdf():
    pdfFile = open('H:\\Prova.pdf', 'rb')
    # Create reader and writer object
    pdfReader = PyPDF2.PdfFileReader(pdfFile)
    pdfWriter = PyPDF2.PdfFileWriter()
    # Add all pages to writer (accepted answer results into blank pages)
    for pageNum in range(pdfReader.numPages):
        pdfWriter.addPage(pdfReader.getPage(pageNum))
    # Encrypt with your password
    pdfWriter.encrypt('password')
    # Write it to an output file. (you can delete unencrypted version now)
    resultPdf = open('H:\\Prova2.pdf', 'wb')
    pdfWriter.write(resultPdf)
    resultPdf.close()
    #pyuic5 -x ui\email.ui -o ui_email.py


def send_email():
    mes = MIMEMultipart()
    mes["From"] = "breaPython@gmail.com"
    mes["To"] = "fabio.brea@gmail.com"
    mes["Subject"] = "TopKek"

    body = MIMEText("Yolo swag", "plain")
    mes.attach(body)

    files = ['H:\\Prova.pdf']

    for f in files or []:
        with open(f, "rb") as fil:
            part = MIMEApplication(
                fil.read(),
                Name=basename(f)
            )
        # After the file is closed
        part['Content-Disposition'] = 'attachment; filename="%s"' % basename(f)
        mes.attach(part)

    try:
        mail = smtplib.SMTP("smtp.gmail.com", 587)
        mail.ehlo()
        mail.starttls()
        mail.login("breaPython@gmail.com", "python92")
        mail.sendmail(mes["from"], mes["To"], mes.as_string())
        mail.close()
    except:
        sys.stderr.write("Failed....")
        sys.stderr.flush()


class EmailSender(QtWidgets.QMainWindow, Ui_EmailWindow):

    def __init__(self) -> None:
        super(EmailSender, self).__init__()
        self.setupUi(self)
        self.born_date_edit.setMaximumDate(QDate.currentDate())

    @pyqtSlot()
    def on_see_password_check_box_clicked(self):
        if self.see_password_check_box.isChecked():
            self.sender_password_line_edit.setEchoMode(QtWidgets.QLineEdit.Normal)
        else:
            self.sender_password_line_edit.setEchoMode(QtWidgets.QLineEdit.Password)

    @pyqtSlot()
    def on_edit_report_button_clicked(self) -> None:
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        filename, _ = QFileDialog.getOpenFileName(self, 'Seleziona il referto',
                                                  os.path.normpath(os.path.expanduser('~/Desktop')), 'File pdf (*.pdf)',
                                                  options=options)
        if filename:
            self.report_line_edit.setText(filename)

    @pyqtSlot()
    def on_edit_bill_button_clicked(self) -> None:
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        filename, _ = QFileDialog.getOpenFileName(self, 'Seleziona la fattura',
                                                  os.path.normpath(os.path.expanduser('~/Desktop')), 'File pdf (*.pdf)',
                                                  options=options)
        if filename:
            self.bill_line_edit.setText(filename)

    @pyqtSlot()
    def on_send_button_clicked(self) -> None:
        self.create_password()

    @pyqtSlot(str)
    def on_surname_line_edit_textChanged(self, text: str) -> None:
        self.create_password()

    @pyqtSlot(str)
    def on_name_line_edit_textChanged(self, text: str) -> None:
        self.create_password()

    @pyqtSlot(QDate)
    def on_born_date_edit_dateChanged(self, date: QDate) -> None:
        self.create_password()

    def create_password(self) -> str:
        password = ''
        if len(self.surname_line_edit.text()) > 0:
            password += self.surname_line_edit.text()[0].upper()
        password += '.'
        if len(self.name_line_edit.text()) > 0:
            password += self.name_line_edit.text()[0].lower()
        date = self.born_date_edit.date()
        password += str(date.day()).zfill(2) + str(date.month()).zfill(2) + str(date.year())
        self.second_email_line_edit.setText(password)
        return password


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    window = EmailSender()
    window.show()
    app.exec_()
