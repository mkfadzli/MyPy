import sys
from PyQt6.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout,
                             QPushButton, QLabel, QTimeEdit, QFontDialog, QColorDialog)
from PyQt6.QtCore import QTimer, QTime, Qt
from PyQt6.QtGui import QFont, QColor, QPalette, QLinearGradient, QGradient

class FuturisticButton(QPushButton):
    def __init__(self, text):
        super().__init__(text)
        self.setFixedSize(150, 50)
        self.setStyleSheet("""
            QPushButton {
                background-color: #2C3E50;
                color: #ECF0F1;
                border: 2px solid #3498DB;
                border-radius: 5px;
                font-weight: bold;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #34495E;
                border-color: #2980B9;
            }
            QPushButton:pressed {
                background-color: #2980B9;
            }
        """)

class FuturisticTimeEdit(QTimeEdit):
    def __init__(self):
        super().__init__()
        self.setFixedSize(150, 50)
        self.setStyleSheet("""
            QTimeEdit {
                background-color: #2C3E50;
                color: #ECF0F1;
                border: 2px solid #3498DB;
                border-radius: 5px;
                padding: 5px;
                font-weight: bold;
                font-size: 16px;
            }
            QTimeEdit::up-button, QTimeEdit::down-button {
                width: 20px;
                background-color: #34495E;
            }
            QTimeEdit QAbstractItemView {
                selection-background-color: #2980B9;
                selection-color: #ECF0F1;
                background-color: #34495E;
                color: #ECF0F1;
            }
        """)

class TimerAlarmClock(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Compact Timer & Alarm')
        self.setFixedSize(340, 340)
        self.setStyleSheet("""
            QWidget {
                background-color: #1A2530;
                color: #ECF0F1;
            }
        """)

        main_layout = QVBoxLayout()
        main_layout.setSpacing(10)
        main_layout.setContentsMargins(10, 10, 10, 10)

        # Time display
        self.time_label = QLabel('00:00:00')
        self.time_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.time_label.setFont(QFont('Arial', 36, QFont.Weight.Bold))
        self.time_label.setStyleSheet("""
            QLabel {
                background-color: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #2C3E50, stop:1 #34495E);
                border: 2px solid #3498DB;
                border-radius: 10px;
                padding: 10px;
            }
        """)
        main_layout.addWidget(self.time_label)

        # Timer and Alarm Edits
        edit_layout = QHBoxLayout()
        self.timer_edit = FuturisticTimeEdit()
        self.timer_edit.setDisplayFormat('HH:mm:ss')
        edit_layout.addWidget(self.timer_edit)

        self.alarm_edit = FuturisticTimeEdit()
        self.alarm_edit.setDisplayFormat('HH:mm:ss')
        edit_layout.addWidget(self.alarm_edit)

        main_layout.addLayout(edit_layout)

        # Start Timer and Set Alarm Buttons
        button_layout = QHBoxLayout()
        self.start_timer_btn = FuturisticButton('Start Timer')
        self.start_timer_btn.clicked.connect(self.start_timer)
        button_layout.addWidget(self.start_timer_btn)

        self.set_alarm_btn = FuturisticButton('Set Alarm')
        self.set_alarm_btn.clicked.connect(self.set_alarm)
        button_layout.addWidget(self.set_alarm_btn)

        main_layout.addLayout(button_layout)

        # Settings Button
        self.settings_btn = FuturisticButton('Settings')
        self.settings_btn.clicked.connect(self.open_settings)
        main_layout.addWidget(self.settings_btn, alignment=Qt.AlignmentFlag.AlignCenter)

        self.setLayout(main_layout)

        # Timer setup
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_time)
        self.timer.start(1000)

        self.countdown_timer = QTimer(self)
        self.countdown_timer.timeout.connect(self.update_countdown)

        self.alarm_time = None

    def update_time(self):
        current_time = QTime.currentTime()
        time_text = current_time.toString('hh:mm:ss')
        self.time_label.setText(time_text)

        if self.alarm_time and current_time >= self.alarm_time:
            self.trigger_alarm()

    def start_timer(self):
        timer_time = self.timer_edit.time()
        self.remaining_time = timer_time.hour() * 3600 + timer_time.minute() * 60 + timer_time.second()
        self.countdown_timer.start(1000)
        self.start_timer_btn.setEnabled(False)

    def update_countdown(self):
        if self.remaining_time > 0:
            self.remaining_time -= 1
            hours, remainder = divmod(self.remaining_time, 3600)
            minutes, seconds = divmod(remainder, 60)
            self.time_label.setText(f"{hours:02d}:{minutes:02d}:{seconds:02d}")
        else:
            self.countdown_timer.stop()
            self.trigger_alarm()
            self.start_timer_btn.setEnabled(True)

    def set_alarm(self):
        self.alarm_time = self.alarm_edit.time()

    def trigger_alarm(self):
        self.time_label.setText("ALARM!")
        self.time_label.setStyleSheet("""
            QLabel {
                background-color: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #C0392B, stop:1 #E74C3C);
                border: 2px solid #ECF0F1;
                border-radius: 10px;
                padding: 10px;
            }
        """)
        # You can add sound or other alarm actions here

    def open_settings(self):
        font, ok = QFontDialog.getFont()
        if ok:
            self.time_label.setFont(font)

        color = QColorDialog.getColor()
        if color.isValid():
            gradient = QLinearGradient(0, 0, 1, 1)
            gradient.setColorAt(0, color.darker(150))
            gradient.setColorAt(1, color)
            gradient.setCoordinateMode(QGradient.CoordinateMode.ObjectBoundingMode)
            
            palette = self.time_label.palette()
            palette.setBrush(QPalette.ColorRole.Window, gradient)
            self.time_label.setPalette(palette)
            self.time_label.setStyleSheet(f"""
                QLabel {{
                    color: {color.name()};
                    border: 2px solid {color.lighter(150).name()};
                    border-radius: 10px;
                    padding: 10px;
                }}
            """)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    timer_alarm_clock = TimerAlarmClock()
    timer_alarm_clock.show()
    sys.exit(app.exec())