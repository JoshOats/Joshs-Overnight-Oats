from PyQt5.QtWidgets import QDialog, QVBoxLayout, QLabel, QProgressBar
from PyQt5.QtCore import Qt
import logging

class UpdateWindow(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowFlags(Qt.Window | Qt.WindowStaysOnTopHint | Qt.CustomizeWindowHint)
        self.setFixedSize(400, 200)

        layout = QVBoxLayout()

        self.title = QLabel("Josh's Overnight Oats")
        self.status = QLabel("Checking for Updates...")
        self.progress = QProgressBar()

        self.title.setAlignment(Qt.AlignCenter)
        self.status.setAlignment(Qt.AlignCenter)

        layout.addWidget(self.title)
        layout.addWidget(self.status)
        layout.addWidget(self.progress)

        self.setLayout(layout)

    def update_status(self, text, progress=None):
        try:
            self.status.setText(text)
            if progress is not None:
                self.progress.setValue(progress)
        except Exception as e:
            logging.error(f"Update status error: {e}")
