# main_window.py

import sys
import os
from PyQt5.QtWidgets import QVBoxLayout, QPushButton, QLabel, QApplication, QWidget
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon
from cnb_transfer_je import CNBTransferJEWindow
from toast_reconcile_window import ToastReconcileWindow
from retro_style import RetroWindow, create_retro_central_widget
from due_to_from_window import DueToFromWindow
from ap_process import APWindow
from ap_reconcile import APReconcileWindow
from tips_reconcile import TipsReconcileWindow
from px_processor import PXGiftCardsWindow
from grubhub_window import GrubHubWindow
from doordash_window import DoorDashWindow
from ubereats_window import UberEatsWindow
from royalties_window import RoyaltiesWindow
from payroll_window import PayrollWindow  # Import the new PayrollWindow

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class AdditionalFunctionsWindow(RetroWindow):
    def __init__(self, parent=None):
        super().__init__()
        self.parent = parent
        self.initUI()

        icon_path = resource_path(os.path.join('assets', 'icon.png'))
        self.setWindowIcon(QIcon(icon_path))

    def initUI(self):
        self.setWindowTitle("Additional Functions")
        self.resize(1000, 738)  # Initial size but resizable
        self.center()

        central_widget = create_retro_central_widget(self)
        layout = QVBoxLayout(central_widget)

        # Create welcome label
        welcome_text = "Additional Functions"
        self.welcome_label = QLabel("", self)
        self.welcome_label.setAlignment(Qt.AlignCenter)
        self.welcome_label.setObjectName("welcome_label")
        layout.addWidget(self.welcome_label)

        # Create typing animation
        self.create_typing_animation(self.welcome_label, welcome_text)

        # Add spacing after welcome message
        layout.addSpacing(20)

        # Create buttons
        self.create_button("CNB TRANSFER JE IMPORT", self.open_cnb_transfer_je, layout)
        self.create_button("DUE TO/FROM ANALYSIS", self.open_due_to_from, layout)
        self.create_button("TOAST NET SALES RECONCILIATION", self.open_toast_reconcile, layout)
        self.create_button("AP RECONCILIATION", self.open_ap_reconcile, layout)
        self.create_button("TIPS RECONCILIATION", self.open_tips_reconcile, layout)
        self.create_button("PX GIFT CARDS", self.open_px_giftcards, layout)
        self.create_button("GRUBHUB JE IMPORT", self.open_grubhub_process, layout)
        self.create_button("DOORDASH JE IMPORT", self.open_doordash_process, layout)
        self.create_button("UBEREATS JE IMPORT", self.open_ubereats_process, layout)
        self.create_button("ROYALTIES PROCESSOR", self.open_royalties_process, layout)
        self.create_button("PAYROLL AUTOMATION", self.open_payroll_automation, layout)  # Added new button

    def create_button(self, text, function, layout):
        btn = QPushButton(text)
        btn.clicked.connect(function)
        layout.addWidget(btn)

    def open_cnb_transfer_je(self):
        self.cnb_transfer_je_window = CNBTransferJEWindow()
        self.cnb_transfer_je_window.show()

    def open_toast_reconcile(self):
        self.toast_reconcile_window = ToastReconcileWindow()
        self.toast_reconcile_window.show()

    def open_due_to_from(self):
        self.due_to_from_window = DueToFromWindow()
        self.due_to_from_window.show()

    def open_ap_reconcile(self):
        self.ap_reconcile_window = APReconcileWindow()
        self.ap_reconcile_window.show()

    def open_tips_reconcile(self):
        self.tips_reconcile_window = TipsReconcileWindow()
        self.tips_reconcile_window.show()

    def open_px_giftcards(self):
        self.px_giftcards_window = PXGiftCardsWindow()
        self.px_giftcards_window.show()

    def open_grubhub_process(self):
        self.grubhub_window = GrubHubWindow()
        self.grubhub_window.show()

    def open_doordash_process(self):
        self.doordash_window = DoorDashWindow()
        self.doordash_window.show()

    def open_ubereats_process(self):
        self.ubereats_window = UberEatsWindow()
        self.ubereats_window.show()

    def open_royalties_process(self):
        self.royalties_window = RoyaltiesWindow()
        self.royalties_window.show()

    def open_payroll_automation(self):  # Added new function
        self.payroll_window = PayrollWindow()
        self.payroll_window.show()

    def center(self):
        frame_geometry = self.frameGeometry()
        center_point = QApplication.desktop().availableGeometry().center()
        frame_geometry.moveCenter(center_point)
        self.move(frame_geometry.topLeft())


class MainWindow(RetroWindow):
    def __init__(self, username=None, available_functions=None):
        super().__init__()
        self.username = username
        self.available_functions = available_functions or []
        self.initUI()

        icon_path = resource_path(os.path.join('assets', 'icon.png'))
        self.setWindowIcon(QIcon(icon_path))

    def initUI(self):
        self.setWindowTitle("Overnight Codes")
        self.resize(1000, 738)  # Initial size but resizable
        self.center()

        central_widget = create_retro_central_widget(self)
        layout = QVBoxLayout(central_widget)

        # Create welcome label
        welcome_text = "Welcome to Overnight Codes!"
        self.welcome_label = QLabel("", self)
        self.welcome_label.setAlignment(Qt.AlignCenter)
        self.welcome_label.setObjectName("welcome_label")
        layout.addWidget(self.welcome_label)

        # Create typing animation
        self.create_typing_animation(self.welcome_label, welcome_text)

        # Add spacing after welcome message
        layout.addSpacing(50)

        # Create AP Payments button
        ap_btn = QPushButton('AP PAYMENTS')
        ap_btn.clicked.connect(self.open_ap_process)
        layout.addWidget(ap_btn)

        # Add more spacing
        layout.addSpacing(10)

        # Create Additional Functions button
        additional_btn = QPushButton('ADDITIONAL FUNCTIONS')
        additional_btn.clicked.connect(self.open_additional_functions)
        layout.addWidget(additional_btn)

    def center(self):
        frame_geometry = self.frameGeometry()
        center_point = QApplication.desktop().availableGeometry().center()
        frame_geometry.moveCenter(center_point)
        self.move(frame_geometry.topLeft())

    def open_ap_process(self):
        self.ap_process_window = APWindow()
        self.ap_process_window.show()

    def open_additional_functions(self):
        self.additional_functions_window = AdditionalFunctionsWindow(self)
        self.additional_functions_window.show()
