import json
import sys
import re

from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import os
from PyQt5 import QtWidgets, QtGui
from os.path import basename

import PyPDF2
import smtplib

from PyQt5.QtCore import pyqtSlot, QDate, QSettings, pyqtSignal, Qt, QThread
from PyQt5.QtWidgets import QFileDialog, QLineEdit, QMessageBox, QDialog

from progress_window_ui import Ui_dialog
from ui_email import Ui_EmailWindow


class EmailSender(QtWidgets.QMainWindow, Ui_EmailWindow):

    def __init__(self) -> None:
        super(EmailSender, self).__init__()
        self.setupUi(self)
        self.born_date_edit.setMaximumDate(QDate.currentDate())
        self.regex = '^[a-zA-Z0-9]+[\\._]?[a-zA-Z0-9]+[@]\\w+[.]\\w+$'

    @pyqtSlot()
    def on_see_password_check_box_clicked(self) -> None:
        if self.see_password_check_box.isChecked():
            self.sender_password_line_edit.setEchoMode(QtWidgets.QLineEdit.Normal)
        else:
            self.sender_password_line_edit.setEchoMode(QtWidgets.QLineEdit.Password)

    @pyqtSlot()
    def on_edit_report_button_clicked(self) -> None:
        settings = QSettings()
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        filename, _ = QFileDialog.getOpenFileName(self, 'Seleziona il referto',
                                                  settings.value('default_dir_key'), 'File pdf (*.pdf)',
                                                  options=options)
        if filename:
            self.report_line_edit.setText(filename)
            settings.setValue('default_dir_key', filename)

    @pyqtSlot()
    def on_edit_attachment_button_clicked(self) -> None:
        settings = QSettings()
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        file_list, _ = QFileDialog.getOpenFileNames(self, 'Seleziona la fattura', settings.value('default_dir_key'),
                                                    'File pdf (*.pdf)', options=options)
        if len(file_list) > 0:
            attachment_text = ''
            for file in file_list:
                attachment_text += file + ', '
            attachment_text = attachment_text[:-2]

            self.attachment_line_edit.setText(attachment_text)

    @pyqtSlot(str)
    def on_sender_email_line_edit_textChanged(self) -> None:
        self.check_email(self.sender_email_line_edit)

    @pyqtSlot(str)
    def on_email_line_edit_textChanged(self) -> None:
        self.check_email(self.email_line_edit)

    @pyqtSlot()
    def on_send_button_clicked(self) -> None:
        if self.can_send_email():
            self.send_email()
            self.clear_patient_data()

    def can_send_email(self) -> bool:
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Critical)

        msg.setText('Ci sono dei campi che non sono stati correttamente riempiti.')
        msg.setWindowTitle('Errore nell\'invio delle email')
        msg.setStandardButtons(QMessageBox.Ok)

        missing_fields = ''
        if len(self.sender_email_line_edit.text()) == 0:
            missing_fields += 'Manca l\'email del mittente\n'
        else:
            if not self.check_email(self.sender_email_line_edit):
                missing_fields += 'L\'email del mittente non è nel formato corretto\n'

        if len(self.sender_password_line_edit.text()) == 0:
            missing_fields += 'Manca la password della mail del mittente\n'

        if len(self.surname_line_edit.text()) == 0:
            missing_fields += 'Manca il cognome del paziente\n'

        if len(self.name_line_edit.text()) == 0:
            missing_fields += 'Manca il nome del paziente\n'

        if len(self.report_line_edit.text()) == 0:
            missing_fields += 'Manca il referto del paziente\n'

        if len(self.email_line_edit.text()) == 0:
            missing_fields += 'Manca l\'email del paziente\n'
        else:
            if not self.check_email(self.email_line_edit):
                missing_fields += 'L\'email del paziente non è nel formato corretto\n'

        if len(self.first_mail_subject.text()) == 0:
            missing_fields += 'Manca l\'oggetto della prima email\n'

        if len(self.first_email_text.toPlainText()) == 0:
            missing_fields += 'Manca il corpo della prima email\n'

        if len(self.second_mail_subject.text()) == 0:
            missing_fields += 'Manca l\'oggetto della seconda email\n'

        if len(self.second_email_text.toPlainText()) == 0:
            missing_fields += 'Manca il corpo della seconda email\n'

        if len(missing_fields) > 0:
            msg.setDetailedText(missing_fields)
            msg.exec_()
            return False

        return True

    def clear_patient_data(self) -> None:
        self.surname_line_edit.clear()
        self.name_line_edit.clear()
        self.report_line_edit.clear()
        self.attachment_line_edit.clear()
        self.email_line_edit.clear()
        default_date = QDate(2000, 1, 1)
        self.born_date_edit.setDate(default_date)

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

    def check_email(self, line_edit: QLineEdit) -> bool:
        if len(line_edit.text()) == 0:
            line_edit.setStyleSheet('')
            return False
        else:
            if re.search(self.regex, line_edit.text()):
                line_edit.setStyleSheet('border: 1.5px solid green')
                return True
            else:
                line_edit.setStyleSheet('border: 1.5px solid red')
                return False

    def send_email(self) -> None:
        progress_dialog = ProgressDownloader(self)
        progress_dialog.show()
        date = str(self.born_date_edit.date().day()) + '/' + str(self.born_date_edit.date().month()) + '/' + str(
            self.born_date_edit.date().year())
        attachment = self.attachment_line_edit.text().split(', ')
        sender_thread = SenderThread(self,
                                     sender_email=self.sender_email_line_edit.text(),
                                     sender_pwd=self.sender_password_line_edit.text(),
                                     receiver_email=self.email_line_edit.text(),
                                     receiver_surname=self.surname_line_edit.text(),
                                     receiver_name=self.name_line_edit.text(),
                                     receiver_date=date,
                                     first_subject=self.first_mail_subject.text(),
                                     first_body=self.first_email_text.toPlainText(),
                                     second_subject=self.second_mail_subject.text(),
                                     second_body=self.second_email_text.toPlainText(),
                                     report=self.report_line_edit.text(),
                                     encrypted_password=self.second_email_line_edit.text(),
                                     attachment=attachment)
        sender_thread.sending_progress.connect(progress_dialog.set_progress)
        sender_thread.start()

    @pyqtSlot()
    def on_save_action_triggered(self) -> None:
        settings = QSettings()
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        filename, _ = QFileDialog.getSaveFileName(self, 'Salva la configurazione',
                                                  settings.value('default_dir_key'), 'Json files (*.json)',
                                                  options=options)
        if filename:
            if not filename.endswith('.json'):
                filename += '.json'

            settings.setValue('default_dir_key', filename)
            data = {
                'mail': self.sender_email_line_edit.text(),
                'first_object': self.first_mail_subject.text(),
                'first_text': self.first_email_text.toPlainText(),
                'second_object': self.second_mail_subject.text(),
                'second_text': self.second_email_text.toPlainText()
            }

            with open(filename, 'w') as output_file:
                json.dump(data, output_file, indent=4)

    @pyqtSlot()
    def on_load_action_triggered(self) -> None:
        settings = QSettings()
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        filename, _ = QFileDialog.getOpenFileName(self, 'Carica la configurazione',
                                                  settings.value('default_dir_key'), 'Json files (*.json)',
                                                  options=options)
        if filename:
            with open(filename) as json_file:
                data = json.load(json_file)
                if 'mail' in data.keys():
                    self.sender_email_line_edit.setText(data['mail'])
                if 'first_object' in data.keys():
                    self.first_mail_subject.setText(data['first_object'])
                if 'first_text' in data.keys():
                    self.first_email_text.setPlainText(data['first_text'])
                if 'second_object' in data.keys():
                    self.second_mail_subject.setText(data['second_object'])
                if 'second_text' in data.keys():
                    self.second_email_text.setPlainText(data['second_text'])


class SenderThread(QThread):
    sending_progress = pyqtSignal(str, int)

    def __init__(self, parent, **kwargs) -> None:
        QThread.__init__(self, parent)
        self.sender_email = kwargs['sender_email']
        self.sender_pwd = kwargs['sender_pwd']
        self.receiver_email = kwargs['receiver_email']
        self.receiver_surname = kwargs['receiver_surname']
        self.receiver_name = kwargs['receiver_name']
        self.receiver_date = kwargs['receiver_date']
        self.first_subject = kwargs['first_subject']
        self.first_body = kwargs['first_body']
        self.second_subject = kwargs['second_subject']
        self.second_body = kwargs['second_body']
        self.report = kwargs['report']
        self.encrypted_password = kwargs['encrypted_password']
        self.attachment = kwargs['attachment']

    def run(self) -> None:
        self.sending_progress.emit('Creo il pdf con password...', 0)
        first_message = MIMEMultipart()
        first_message["From"] = self.sender_email
        first_message["To"] = self.receiver_email
        first_message["Subject"] = self.first_subject

        first_body_formatted = self.format_text(self.first_body)
        body = MIMEText(first_body_formatted, "plain")
        first_message.attach(body)

        files = [self.create_pdf_with_password()]
        self.sending_progress.emit('Invio la prima email', 30)

        if len(self.attachment) > 0:
            files.extend(self.attachment)

        for f in files or []:
            with open(f, "rb") as fil:
                part = MIMEApplication(fil.read(), Name=basename(f))
            part['Content-Disposition'] = 'attachment; filename="%s"' % basename(f)
            first_message.attach(part)

        try:
            mail = smtplib.SMTP("smtp.gmail.com", 587)
            mail.ehlo()
            mail.starttls()
            mail.login(self.sender_email, self.sender_pwd)
            mail.sendmail(first_message["from"], first_message["To"], first_message.as_string())
            mail.close()
        except:
            sys.stderr.write("Failed....")
            sys.stderr.flush()

        self.sending_progress.emit('Invio la seconda email', 60)
        second_message = MIMEMultipart()
        second_message["From"] = self.sender_email
        second_message["To"] = self.receiver_email
        second_message["Subject"] = self.second_subject

        second_body_formatted = self.format_text(self.second_body)
        if self.encrypted_password not in second_body_formatted:
            second_body_formatted += '\n' + self.encrypted_password
        body = MIMEText(second_body_formatted, "plain")
        second_message.attach(body)
        try:
            mail = smtplib.SMTP("smtp.gmail.com", 587)
            mail.ehlo()
            mail.starttls()
            mail.login(self.sender_email, self.sender_pwd)
            mail.sendmail(second_message["from"], second_message["To"], second_message.as_string())
            mail.close()
        except:
            sys.stderr.write("Failed....")
            sys.stderr.flush()
        self.sending_progress.emit('Email inviate con successo', 100)

    def create_pdf_with_password(self) -> str:
        file_name = os.path.splitext(self.report)[0]
        pdf_file = open(self.report, 'rb')
        pdf_reader = PyPDF2.PdfFileReader(pdf_file)
        pdf_writer = PyPDF2.PdfFileWriter()

        for page_num in range(pdf_reader.numPages):
            pdf_writer.addPage(pdf_reader.getPage(page_num))

        password = self.encrypted_password
        password_file = open(file_name + '_password.txt', 'w')
        password_file.write(password)
        password_file.close()

        pdf_writer.encrypt(password)

        pdf_with_password = open(file_name + '_encrypted.pdf', 'wb')
        pdf_writer.write(pdf_with_password)
        pdf_with_password.close()
        return file_name + '_encrypted.pdf'

    def format_text(self, text: str) -> str:
        text = text.replace('{NOME}', self.receiver_name)
        text = text.replace('{COGNOME}', self.receiver_surname)
        text = text.replace('{DATA}', self.receiver_date)
        text = text.replace('{PASSWORD}', self.encrypted_password)
        return text


class ProgressDownloader(QDialog, Ui_dialog):

    def __init__(self, parent) -> None:
        super(ProgressDownloader, self).__init__(parent)
        self.setupUi(self)
        self.setWindowTitle('Inviando le email')
        self.setWindowFlags(Qt.Dialog | Qt.WindowTitleHint)
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(100)

    def closeEvent(self, event: QtGui.QCloseEvent) -> None:
        QMessageBox.information(self, "Email inviate", "Hai inviato le email con successo.", QMessageBox.Ok,
                                QMessageBox.Ok)

    @pyqtSlot(str, int)
    def set_progress(self, message: str, progress: int) -> None:
        self.message_label.setText(message)
        self.progress_bar.setValue(progress)
        if progress == 100:
            self.close()


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    app.setApplicationName('Email_Project')
    app.setOrganizationName('Fabio Brea')
    app.setOrganizationDomain('https://fabiobreo.github.io/')
    window = EmailSender()
    window.show()
    app.exec_()
