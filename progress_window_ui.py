# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'ui\progress_window.ui'
#
# Created by: PyQt5 UI code generator 5.15.1
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_dialog(object):
    def setupUi(self, dialog):
        dialog.setObjectName("dialog")
        dialog.setWindowModality(QtCore.Qt.WindowModal)
        dialog.resize(400, 58)
        dialog.setModal(True)
        self.verticalLayout = QtWidgets.QVBoxLayout(dialog)
        self.verticalLayout.setObjectName("verticalLayout")
        self.message_label = QtWidgets.QLabel(dialog)
        self.message_label.setObjectName("message_label")
        self.verticalLayout.addWidget(self.message_label)
        self.progress_bar = QtWidgets.QProgressBar(dialog)
        self.progress_bar.setMaximum(1)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setOrientation(QtCore.Qt.Horizontal)
        self.progress_bar.setInvertedAppearance(False)
        self.progress_bar.setTextDirection(QtWidgets.QProgressBar.TopToBottom)
        self.progress_bar.setObjectName("progress_bar")
        self.verticalLayout.addWidget(self.progress_bar)

        self.retranslateUi(dialog)
        QtCore.QMetaObject.connectSlotsByName(dialog)

    def retranslateUi(self, dialog):
        _translate = QtCore.QCoreApplication.translate
        dialog.setWindowTitle(_translate("dialog", "Downloading"))
        self.message_label.setText(_translate("dialog", "Message"))


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    dialog = QtWidgets.QDialog()
    ui = Ui_dialog()
    ui.setupUi(dialog)
    dialog.show()
    sys.exit(app.exec_())
