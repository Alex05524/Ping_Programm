import sys
import logging
import datetime
import pandas as pd
from datetime import datetime
import matplotlib.pyplot as plt
from ping3 import ping
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Font, PatternFill
from concurrent.futures import ThreadPoolExecutor, CancelledError
from PyQt5.QtWidgets import (
    QLabel, QApplication, QWidget, QVBoxLayout, QHBoxLayout, QTableWidget,
    QTableWidgetItem, QPushButton, QFileDialog, QHeaderView, QMessageBox, QDialog,
    QLineEdit, QDialogButtonBox, QSizePolicy, QInputDialog, QListWidget, QListWidgetItem
)
from PyQt5.QtCore import QThread, QThreadPool, pyqtSignal, QRunnable, QObject, QSettings
from PyQt5.QtGui import QIcon, QPainter, QColor
from matplotlib.figure import Figure
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.backends.backend_qt import NavigationToolbar2QT

class PingWorkerSignals(QObject):
    result = pyqtSignal(str, str, str)

class PingWorker(QRunnable):
    def __init__(self, ip, ping_count, timeout):
        super(PingWorker, self).__init__()
        self.ip = ip
        self.ping_count = ping_count
        self.timeout = timeout
        self.signals = PingWorkerSignals()

    def run(self):
        rtt_sum = 0
        status = 'Не отвечает'

        for _ in range(self.ping_count):
            if self.signals.cancelled:
                return

            try:
                start_time = datetime.datetime.now()
                rtt = ping(self.ip, timeout=self.timeout)
                end_time = datetime.datetime.now()

                if rtt is not None:
                    status = 'Да'
                    rtt_sum += rtt

                log_message = f"Pinging {self.ip}: Status: {status}, RTT: {rtt}, Time: {end_time - start_time}"
                logging.debug(log_message)

            except Exception as e:
                logging.error(f"An error occurred while pinging {self.ip}: {e}")

        if status == 'Да':
            rtt_average = rtt_sum / self.ping_count
        else:
            rtt_average = ''

        self.signals.result.emit(self.ip, status, str(rtt_average))

class RoundedCornersWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.radius = 15

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setBrush(QColor("#f0f0f0"))  # Background color
        painter.setPen(QColor("#f0f0f0"))
        painter.drawRoundedRect(self.rect(), self.radius, self.radius)

class SettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Настройки')
        layout = QVBoxLayout(self)

        self.ping_count_label = QLabel('Количество пингов:', self)
        self.ping_count_edit = QLineEdit(self)

        self.timeout_label = QLabel('Таймаут (сек):', self)
        self.timeout_edit = QLineEdit(self)

        layout.addWidget(self.ping_count_label)
        layout.addWidget(self.ping_count_edit)
        layout.addWidget(self.timeout_label)
        layout.addWidget(self.timeout_edit)

        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, self)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)

class MplCanvas(FigureCanvas):
    def __init__(self, parent=None, width=5, height=4, dpi=100):
        fig = plt.Figure(figsize=(width, height), dpi=dpi)
        self.ax = fig.add_subplot(111)
        FigureCanvas.__init__(self, fig)
        self.setParent(parent)
        FigureCanvas.setSizePolicy(self, QSizePolicy.Expanding, QSizePolicy.Expanding)
        FigureCanvas.updateGeometry(self)


class PingThread(QThread):
    update_table = pyqtSignal(str, str, str)

    def __init__(self, ip_addresses, ping_count=4, timeout=1.0):
        super().__init__()
        self.ip_addresses = ip_addresses
        self.ping_count = ping_count
        self.timeout = timeout
        self.cancelled = False

    def run(self):
        try:
            for ip in self.ip_addresses:
                rtt_sum = 0
                status = 'Не отвечает'

                for _ in range(self.ping_count):
                    if self.cancelled:
                        return

                    try:
                        start_time = datetime.datetime.now()
                        rtt = ping(ip, timeout=self.timeout)
                        end_time = datetime.datetime.now()

                        if rtt is not None:
                            status = 'Да'
                            rtt_sum += rtt

                        log_message = f"Pinging {ip}: Status: {status}, RTT: {rtt}, Time: {end_time - start_time}"
                        logging.debug(log_message)

                    except Exception as e:
                        logging.error(f"An error occurred while pinging {ip}: {e}")

                if status == 'Да':
                    rtt_average = rtt_sum / self.ping_count
                else:
                    rtt_average = ''

                self.update_table.emit(ip, status, str(rtt_average))

        except Exception as e:
            logging.error(f"An error occurred in the ping thread: {e}")

    def cancel_ping(self):
        self.cancelled = True

class App(RoundedCornersWidget):
    def __init__(self):
        super().__init__()
        self.thread = None
        self.initUI()

    def initUI(self):
        self.setGeometry(100, 100, 1200, 800)
        self.setWindowTitle('Программа Ping')
        self.setup_styles()

        # Виджеты и кнопки для IP-адресов
        self.add_ip_button = self.create_button('Добавить IP', self.show_ip_dialog)
        self.add_ip_button.setIcon(QIcon('C:\\Users\\Alex\\Desktop\\practice\\Ping_Programm\\Icons\\Add.png'))
        self.remove_ip_button = self.create_button('Удалить IP', self.remove_selected_ip)
        self.remove_ip_button.setIcon(QIcon('C:\\Users\\Alex\\Desktop\\practice\\Ping_Programm\\Icons\\Delete.png'))
        self.ip_list_widget = QListWidget(self)

        # Виджеты и кнопки для мониторинга
        self.start_button = self.create_button('Старт мониторинг', self.start_monitoring)
        self.stop_button = self.create_button('Стоп мониторинг', self.stop_monitoring)
        self.save_button = self.create_button('Сохранить в Excel', self.save_to_excel)

        # Иконки для кнопок
        self.start_button.setIcon(QIcon('C:\\Users\\Alex\\Desktop\\practice\\Ping_Programm\\Icons\\Start.png'))
        self.stop_button.setIcon(QIcon('C:\\Users\\Alex\\Desktop\\practice\\Ping_Programm\\Icons\\Stop.png'))
        self.save_button.setIcon(QIcon('C:\\Users\\Alex\\Desktop\\practice\\Ping_Programm\\Icons\\Export.png'))

        # Дополнительные кнопки и настройки
        self.exit_button = self.create_button('Выход', self.close)
        self.exit_button.setIcon(QIcon('C:\\Users\\Alex\\Desktop\\practice\\Ping_Programm\\Icons\\Exit.png'))
        
        self.settings_button = self.create_button('Настройки', self.show_settings_dialog)
        self.settings_button.setIcon(QIcon('C:\\Users\\Alex\\Desktop\\practice\\Ping_Programm\\Icons\\Settings.png'))

        # Виджеты для отображения данных
        self.status_label = QLabel('Status: Ready', self)
        self.table = QTableWidget(self)

        # График
        self.plot_widget = MplCanvas(self)
        self.plot_toolbar = NavigationToolbar2QT(self.plot_widget, self)

        # Отображение виджетов в макете
        self.setup_layout()

        # Начальные значения для мониторинга
        self.ping_count = 4
        self.timeout = 1.0

        # Загрузка настроек
        self.load_settings()

        self.show()

    def create_button(self, text, slot):
        button = QPushButton(text, self)
        button.clicked.connect(slot)
        return button

    def load_settings(self):
        settings = QSettings("YourOrganizationName", "YourAppName")
        ip_addresses_str = settings.value("ip_addresses", defaultValue="192.168.30.3,192.168.54.3,192.168.101.3,192.168.103.3,192.168.104.3,192.168.105.3,192.168.106.3,192.168.108.3,192.168.109.3")
        ip_addresses = ip_addresses_str.split(",")
        self.ping_count = settings.value("ping_count", defaultValue=4, type=int)
        self.timeout = settings.value("timeout", defaultValue=1.0, type=float)

        # Обновление соответствующих элементов интерфейса
        self.update_settings_ui(ip_addresses)

    def save_settings(self):
        settings = QSettings("YourOrganizationName", "YourAppName")
        ip_addresses = ",".join(self.get_ip_addresses())  # Преобразование списка в строку с разделителями
        settings.setValue("ip_addresses", ip_addresses)
        settings.setValue("ping_count", self.ping_count)
        settings.setValue("timeout", self.timeout)

    def update_settings_ui(self, ip_addresses):
        # Очистка списка IP-адресов
        self.ip_list_widget.clear()
        
        # Добавление IP-адресов в список
        for ip in ip_addresses:
            item = QListWidgetItem(ip)
            self.ip_list_widget.addItem(item)

    def show_ip_dialog(self):
        # Диалог для добавления нового IP-адреса
        ip, ok_pressed = QInputDialog.getText(self, "Добавить IP-адрес", "Введите IP-адрес:")
        if ok_pressed and ip:
            # Добавить новый IP-адрес
            self.add_ip_address(ip)

    def add_ip_address(self, ip):
        # Добавление нового IP-адреса в список и в QSettings
        ip_addresses = self.get_ip_addresses()
        ip_addresses.append(ip)
        self.update_settings_ui(ip_addresses)
        self.save_settings()

    def closeEvent(self, event):
        # Сохранение настроек при закрытии приложения
        self.save_settings()
        event.accept()

    def remove_selected_ip(self):
        selected_item = self.ip_list_widget.currentItem()
        if selected_item is not None:
            ip_addresses = self.get_ip_addresses()
            ip_to_remove = selected_item.text()

        # Проверка наличия IP-адреса в списке перед удалением
            if ip_to_remove in ip_addresses:
                ip_addresses.remove(ip_to_remove)
                self.update_settings_ui(ip_addresses)
                self.save_settings()
            else:
                QMessageBox.warning(self, 'Warning', 'Selected IP address not found in the list', QMessageBox.Ok)


    def get_ip_addresses(self):    
        # Получение списка IP-адресов из QSettings
        settings = QSettings("YourOrganizationName", "YourAppName")
        ip_addresses = settings.value("ip_addresses", defaultValue=[
        "192.168.30.3", "192.168.54.3", "192.168.101.3",
        "192.168.103.3", "192.168.104.3", "192.168.105.3",
        "192.168.106.3", "192.168.108.3", "192.168.109.3"
    ])

        if isinstance(ip_addresses, str):
            ip_addresses = ip_addresses.split(",")
        elif not isinstance(ip_addresses, list):
            ip_addresses = []

        return ip_addresses

    def setup_styles(self):
        stylesheet = """
            RoundedCornersWidget {
                background-color: #f0f0f0;
            }
            QPushButton {
                background-color: #2f80ed;
                color: #ffffff;
                padding: 10px;
                border: none;
                border-radius: 5px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #2a66b5;
            }
            QTableWidget {
                background-color: #ffffff;
                border-radius: 10px;
            }
            QTableWidget::item {
                padding: 5px;
            }
        """
        self.setStyleSheet(stylesheet)

    def setup_layout(self):
        layout = QVBoxLayout(self)

        # Верхняя часть интерфейса
        top_layout = QHBoxLayout()
        top_layout.addWidget(self.add_ip_button)
        top_layout.addWidget(self.remove_ip_button)
        top_layout.addWidget(self.ip_list_widget)
        layout.addLayout(top_layout)

        # Центральная часть интерфейса
        middle_layout = QVBoxLayout()
        middle_layout.addWidget(self.status_label)
        middle_layout.addWidget(self.table)
        layout.addLayout(middle_layout)

        # Нижняя часть интерфейса
        bottom_layout = QHBoxLayout()
        bottom_layout.addWidget(self.start_button)
        bottom_layout.addWidget(self.stop_button)
        bottom_layout.addWidget(self.save_button)
        bottom_layout.addWidget(self.exit_button)
        bottom_layout.addWidget(self.settings_button)
        layout.addLayout(bottom_layout)

        # Добавление графика
        layout.addWidget(self.plot_toolbar)
        layout.addWidget(self.plot_widget)

        # Настройки таблицы
        self.table.setColumnCount(2)
        self.table.setHorizontalHeaderLabels(['IP Address', 'Status'])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

    def show_settings_dialog(self):
        dialog = SettingsDialog(self)
        dialog.ping_count_edit.setText(str(self.ping_count))
        dialog.timeout_edit.setText(str(self.timeout))

        result = dialog.exec_()
        if result == QDialog.Accepted:
            try:
                self.ping_count = int(dialog.ping_count_edit.text())
                self.timeout = float(dialog.timeout_edit.text())
            except ValueError:
                QMessageBox.warning(self, 'Warning', 'Invalid input format for settings', QMessageBox.Ok)

    def start_monitoring(self):
        ip_addresses = self.get_ip_addresses()
        self.thread = PingThread(ip_addresses, ping_count=self.ping_count, timeout=self.timeout)
        self.thread.update_table.connect(self.update_table)
        self.thread.start()
        self.status_label.setText('Мониторинг запущен')
        self.start_button.setEnabled(False)
        self.stop_button.setEnabled(True)
        self.save_button.setEnabled(False)

    def stop_monitoring(self):
        if self.thread is not None:
            self.thread.cancel_ping()
            self.thread.wait()
            self.thread = None
        self.status_label.setText('Мониторинг остановлен')
        self.start_button.setEnabled(True)
        self.stop_button.setEnabled(False)
        self.save_button.setEnabled(True)

    def update_table(self, ip, status, rtt):
        row_position = self.table.rowCount()
        self.table.insertRow(row_position)
        self.table.setItem(row_position, 0, QTableWidgetItem(ip))
        self.table.setItem(row_position, 1, QTableWidgetItem(status))

        # Update the plot
        self.update_plot()

    def update_plot(self):
        self.plot_widget.ax.clear()

        ips = [self.table.item(row, 0).text() for row in range(self.table.rowCount())]
        statuses = [self.table.item(row, 1).text() for row in range(self.table.rowCount())]
        rtts = [float(self.table.item(row, 2).text()) if self.table.item(row, 2) is not None else 0 for row in range(self.table.rowCount())]

        # Выбираем цвета столбцов в зависимости от статуса
        colors = ['green' if status == 'Да' else 'red' for status in statuses]

        # Создаем столбчатую диаграмму
        self.plot_widget.ax.bar(ips, rtts, color=colors)

        # Настройка осей и заголовка
        self.plot_widget.ax.set_xlabel('IP Address')
        self.plot_widget.ax.set_ylabel('Response Time (ms)')
        self.plot_widget.ax.set_title('Latency Graph')

        # Добавляем легенду
        self.plot_widget.ax.legend(['Success' if status == 'Да' else 'Failure' for status in statuses])

        self.plot_widget.draw()

    def update_text(self, text):
        self.textbox.appendPlainText(text)

    def save_to_excel(self):
        filename, _ = QFileDialog.getSaveFileName(self, 'Сохранить в Excel', '', 'Excel Files (*.xlsx)')
        if filename:
            data = [(self.table.item(row, 0).text(), self.table.item(row, 1).text())
                    for row in range(self.table.rowCount())]
            df = pd.DataFrame(data, columns=['IP Address', 'Status'])

            try:
                with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name='Sheet1', index=False, header=True)

                QMessageBox.information(self, 'Success', 'File saved successfully!', QMessageBox.Ok)

            except Exception as e:
                QMessageBox.critical(self, 'Error', f'An error occurred while saving to Excel: {str(e)}', QMessageBox.Ok)

def main():
    logging.basicConfig(filename='ping_app.log', level=logging.ERROR)
    app = QApplication(sys.argv)
    ex = App()
    try:
        sys.exit(app.exec_())
    except KeyboardInterrupt:
        print("KeyboardInterrupt: Программа завершена пользователем")
        if ex.thread is not None:
            ex.thread.terminate()
            ex.thread.wait()

if __name__ == '__main__':
    main()
