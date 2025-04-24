from PyQt5.QtWidgets import QWidget, QMainWindow
from PyQt5.QtGui import QFont, QPalette, QColor
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtWidgets import QLabel, QDialog
from PyQt5.QtGui import QPainter, QPixmap, QTransform, QPen, QRadialGradient
from PyQt5.QtCore import Qt, QRect
import logging
from PyQt5.QtWidgets import QSplashScreen
import os
from PyQt5.QtGui import QLinearGradient

def create_pixel_border(widget):
    top_border = QLabel(widget)
    bottom_border = QLabel(widget)
    left_border = QLabel(widget)
    right_border = QLabel(widget)

    # Create pixelated borders for all sides
    border_size = 4
    for border in [top_border, bottom_border, left_border, right_border]:
        border_pixmap = QPixmap(widget.width(), border_size)
        border_pixmap.fill(QColor(255, 255, 255))  # White border

        with QPainter(border_pixmap) as painter:
            painter.setPen(Qt.transparent)
            # Create pixelated effect
            for i in range(0, widget.width(), 4):
                painter.fillRect(i, 0, 2, 2, Qt.transparent)

        border.setPixmap(border_pixmap)

    # Position borders
    top_border.move(0, 0)
    bottom_border.move(0, widget.height() - border_size)
    left_border.setPixmap(border_pixmap.transformed(QTransform().rotate(90)))
    right_border.setPixmap(border_pixmap.transformed(QTransform().rotate(90)))
    left_border.move(0, 0)
    right_border.move(widget.width() - border_size, 0)

class RetroSplash(QSplashScreen):
    def __init__(self):
        # Create a custom pixmap for the splash screen
        width = 600
        height = 400
        self._pixmap = QPixmap(width, height)
        painter = QPainter(self._pixmap)

        # Fill background with gradient
        gradient = QLinearGradient(0, 0, 0, height)
        gradient.setColorAt(0, QColor("#330833"))
        gradient.setColorAt(1, QColor("#230823"))
        painter.fillRect(0, 0, width, height, gradient)

        # Draw border
        pen = QPen(QColor("#8A2BE2"))  # Purple border from theme
        pen.setWidth(4)
        painter.setPen(pen)
        painter.drawRect(2, 2, width-4, height-4)

        # Draw title with shadow
        title_font = QFont('Courier', 24, QFont.Bold)
        painter.setFont(title_font)

        # Draw shadow text
        painter.setPen(QColor(0, 0, 0, 180))  # Semi-transparent black
        painter.drawText(3, 53, width, 50, Qt.AlignCenter, "Overnight Codes")

        # Draw main text
        painter.setPen(QColor("#FFFFFF"))  # White text
        painter.drawText(0, 50, width, 50, Qt.AlignCenter, "Overnight Codes")

        # Load and draw icon
        icon_size = 128
        icon = QPixmap(os.path.join('assets', 'icon.png'))
        if not icon.isNull():
            scaled_icon = icon.scaled(icon_size, icon_size, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            painter.drawPixmap((width - icon_size) // 2, 120, scaled_icon)

        # Add scanline effect
        scanline_pen = QPen(QColor(0, 0, 0, 75))  # Semi-transparent black
        scanline_pen.setWidth(2)
        painter.setPen(scanline_pen)
        for y in range(0, height, 5):
            painter.drawLine(0, y, width, y)

        # Add vignette effect
        gradient = QRadialGradient(width/2, height/2, max(width/1.5, height/1.5))
        gradient.setColorAt(0.0, QColor(0, 0, 0, 0))
        gradient.setColorAt(0.5, QColor(0, 0, 0, 0))
        gradient.setColorAt(1.0, QColor(0, 0, 0, 130))
        painter.setBrush(gradient)
        painter.setPen(Qt.NoPen)
        painter.drawRect(0, 0, width, height)

        # Set up status area
        self.status_rect = QRect(0, height - 100, width, 30)

        painter.end()
        super().__init__(self._pixmap)
        self.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.SplashScreen)

    def drawContents(self, painter):
        """Override drawContents to handle message drawing"""
        if self.message():
            # Draw message with shadow
            painter.setFont(QFont('Courier', 12))

            # First draw message shadow
            painter.setPen(QColor(0, 0, 0, 180))
            shadow_rect = self.status_rect.adjusted(2, 2, 2, 2)
            painter.drawText(shadow_rect, Qt.AlignCenter, self.message())

            # Then draw actual message
            painter.setPen(QColor("#FFFFFF"))
            painter.drawText(self.status_rect, Qt.AlignCenter, self.message())

    def mousePressEvent(self, event):
        """Prevent closing on mouse click"""
        pass

class ScanlineEffect(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAttribute(Qt.WA_TransparentForMouseEvents)
        self.setAttribute(Qt.WA_NoSystemBackground)
        self.setStyleSheet("background: transparent;")
        self.setAutoFillBackground(False)
        self.show()

    def paintEvent(self, event):
        painter = QPainter(self)

        # Draw scanlines more visibly
        width = self.width()
        height = self.height()

        # Make scanlines more visible
        scanline_pen = QPen(QColor(0, 0, 0, 75))  # Increased opacity
        scanline_pen.setWidth(2)  # Ensure line width is 1 pixel
        painter.setPen(scanline_pen)

        # Draw scanlines closer together
        for y in range(0, height, 5):  # Changed from 2 to 3 pixels spacing
            painter.drawLine(0, y, width, y)

        # Optional: Add a very subtle vertical scanline effect
        # vertical_pen = QPen(QColor(255, 255, 255, 15))  # Very subtle white lines
        # vertical_pen.setWidth(1)
        # painter.setPen(vertical_pen)
        # for x in range(0, width, 4):
            # painter.drawLine(x, 0, x, height)
        # Draw vignette effect
        gradient = QRadialGradient(width/2, height/2,
                                 max(width/1.5, height/1.5))
        gradient.setColorAt(0.0, QColor(0, 0, 0, 0))    # Center is transparent
        gradient.setColorAt(0.5, QColor(0, 0, 0, 0))    # Start darkening halfway
        gradient.setColorAt(1.0, QColor(0, 0, 0, 130))  # Edges are dark

        painter.setBrush(gradient)
        painter.setPen(Qt.NoPen)
        painter.drawRect(0, 0, width, height)

class RetroDialog(QDialog):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.apply_8bit_style()

        # Add scanlines and ensure they cover the whole dialog
        self.scanlines = ScanlineEffect(self)
        self.scanlines.setGeometry(0, 0, self.width(), self.height())
        self.scanlines.raise_()  # Try raising instead of lowering


    def resizeEvent(self, event):
        super().resizeEvent(event)
        if hasattr(self, 'scanlines'):
            self.scanlines.setGeometry(0, 0, self.width(), self.height())
            self.scanlines.raise_() #to make it visisible over all other elements

    def apply_8bit_style(self):
        # Use the same method as RetroWindow
        font = QFont("Courier", 10)
        font.setBold(True)
        self.setFont(font)

        palette = QPalette()
        palette.setColor(QPalette.Window, QColor(42, 10, 41))
        palette.setColor(QPalette.WindowText, QColor(255, 255, 255))
        palette.setColor(QPalette.Button, QColor(138, 43, 226))
        palette.setColor(QPalette.ButtonText, QColor(255, 255, 255))
        palette.setColor(QPalette.Highlight, QColor(255, 0, 255))
        palette.setColor(QPalette.Base, QColor(30, 0, 30))
        palette.setColor(QPalette.Text, QColor(255, 255, 255))
        self.setPalette(palette)


class RetroWindow(QMainWindow):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.apply_8bit_style()

        # Add scanlines first
        self.scanlines = ScanlineEffect(self)
        self.scanlines.setGeometry(0, 0, self.width(), self.height())
        self.scanlines.raise_() #to make visible over all other elements

        # Then add other elements
        if hasattr(self, 'console_output'):
            self.setup_blinking_cursor()

        # Make sure scanlines are on top again after decorations
        if hasattr(self, 'scanlines'):
            self.scanlines.raise_()

    def resizeEvent(self, event):
        # Handle the resize event first
        super().resizeEvent(event)

        # Resize scanlines to match new window size
        if hasattr(self, 'scanlines'):
            self.scanlines.setGeometry(0, 0, self.width(), self.height())
            self.scanlines.raise_()  # Make sure scanlines stay visible

        # Recreate pixel border and corner decorations after resize


    def create_typing_animation(self, label, text, typing_speed=20, cursor_blink_speed=100, shadow_color=QColor(0, 0, 0), shadow_offset_x=3, shadow_offset_y=3):
        """
        Enhanced method to create a typing animation with shadow effect using QPainter

        Args:
            label: The QLabel to animate
            text: The text to display
            typing_speed: Milliseconds between each character (default: 100)
            cursor_blink_speed: Milliseconds for cursor blink interval (default: 530)
            shadow_color: QColor for shadow (default: black)
            shadow_offset_x: Shadow X offset in pixels (default: 2)
            shadow_offset_y: Shadow Y offset in pixels (default: 2)
        """
        self.animation_text = text
        self.animation_index = 0
        self.animation_label = label
        self.animation_label.setText("")

        # Store shadow parameters
        self.shadow_color = shadow_color
        self.shadow_offset_x = shadow_offset_x
        self.shadow_offset_y = shadow_offset_y

        # Enable custom painting on the label
        self.animation_label.paintEvent = self._paint_label

        # Create and start typing timer with configurable speed
        self.typing_timer = QTimer(self)
        self.typing_timer.timeout.connect(self._type_next_char)
        self.typing_timer.start(typing_speed)

        # Create and start cursor blink timer with configurable speed
        self.cursor_visible = True
        self.cursor_timer = QTimer(self)
        self.cursor_timer.timeout.connect(self._blink_cursor)
        self.cursor_timer.start(cursor_blink_speed)

    def _paint_label(self, event):
        """Custom paint event for the label to draw centered text with shadow"""
        painter = QPainter(self.animation_label)
        painter.setRenderHint(QPainter.Antialiasing)

        # Get current text
        current_text = self.animation_text[:self.animation_index]
        if hasattr(self, 'cursor_visible') and self.cursor_visible and self.animation_index < len(self.animation_text):
            current_text += "█"

        # Get label dimensions
        rect = self.animation_label.rect()

        # Set font
        font = self.animation_label.font()
        painter.setFont(font)

        # Draw shadow
        painter.setPen(self.shadow_color)
        shadow_rect = QRect(
            rect.x() + self.shadow_offset_x,
            rect.y() + self.shadow_offset_y,
            rect.width(),
            rect.height()
        )
        painter.drawText(shadow_rect, Qt.AlignCenter, current_text)  # Changed to AlignCenter

        # Draw main text
        painter.setPen(QColor(255, 255, 255))  # White text
        painter.drawText(rect, Qt.AlignCenter, current_text)  # Changed to AlignCenter

        painter.end()

    def _type_next_char(self):
        if self.animation_index < len(self.animation_text):
            self.animation_index += 1
            self.animation_label.update()  # Trigger repaint
        else:
            self.typing_timer.stop()
            self.cursor_timer.stop()  # Stop cursor blinking when done
            self.animation_label.update()  # Final repaint

    def _blink_cursor(self):
        if hasattr(self, 'animation_label') and self.animation_index < len(self.animation_text):
            self.cursor_visible = not self.cursor_visible
            self.animation_label.update()  # Trigger repaint

    def apply_8bit_style(self):
        self.setStyleSheet("""
            QDialog {
                background: qlineargradient(
                    x1: 0, y1: 0, x2: 0, y2: 1,
                    stop: 0 #330833,
                    stop: 1 #230823
                );
            }
            QMainWindow, QWidget {
                background: qlineargradient(
                    x1: 0, y1: 0, x2: 0, y2: 1,
                    stop: 0 #330833,
                    stop: 1 #230823
                );
            }
            QLabel {
                color: #FFFFFF;
                font-family: 'Courier';
                font-size: 20px;
                font-weight: none;
                border: none;
                background: transparent;
            }
            QLabel#welcome_label, QLabel#login_label {
                color: #FFFFFF;
                font-family: 'Courier';
                font-size: 28px;
                font-weight: bold;
                padding: 20px;
                letter-spacing: 2px;
                background: transparent;
                text-shadow: 3px 3px 6px rgba(0, 0, 0, 0.8),
                            -1px -1px 4px rgba(0, 0, 0, 0.4);
            }
            QPushButton {
                background: qlineargradient(
                    x1: 0, y1: 0, x2: 0, y2: 1,
                    stop: 0 #4B0082,
                    stop: 1 #2A0058
                ) !important;
                color: #FFFFFF;
                border: 2px solid #8A2BE2;
                border-style: outset;
                border-radius: 0px;
                padding: 6px;
                margin: 5px;
                font-family: 'Courier';
                font-size: 20px;
                text-transform: uppercase;
            }

            QPushButton:hover {
                background: qlineargradient(
                    x1: 0, y1: 0, x2: 0, y2: 1,
                    stop: 0 #6B0082,
                    stop: 1 #4B0082
                ) !important;
                border: 2px solid #8A2BE2;
                border-style: outset;
            }

            QPushButton:pressed {
                background: qlineargradient(
                    x1: 0, y1: 0, x2: 0, y2: 1,
                    stop: 0 #2A0058,
                    stop: 1 #1A0038
                ) !important;
                border-style: inset;
                padding: 12px 8px 8px 12px;
            }
            QTextEdit, QListWidget {
                background-color: #2A0A29;
                color: #FFFFFF;
                border: 2px solid #4B0082;
                font-family: 'Courier';
                font-size: 18px;
            }
            QLineEdit {
                background-color: #2A0A29;
                color: #FFFFFF;
                border: 2px solid #4B0082;
                padding: 5px;
                font-family: 'Courier';
                font-size: 18px;
            }
                    QLabel#title_label {
                color: #FFFFFF;
                font-family: 'Courier';
                font-size: 24px;
                font-weight: bold;
                border: none;
                background: transparent;
            }
        """)



    def update_console(self, message):
        if hasattr(self, 'console_output'):
            current_text = self.console_output.toPlainText()
            if current_text.endswith("█") or current_text.endswith(" "):
                current_text = current_text[:-1]
            self.console_output.setPlainText(current_text + "\n" + message + "█")

def create_retro_central_widget(window):
    central_widget = QWidget(window)
    window.setCentralWidget(central_widget)
    return central_widget
