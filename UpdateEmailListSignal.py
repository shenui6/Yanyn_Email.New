from PyQt6.QtCore import pyqtSignal, QObject


class UpdateEmailListSignal(QObject):
    # 声明信号
    signal = pyqtSignal()  # 注意这里是类属性

    def emit_signal(self):
        # 添加一个方法来发射信号
        self.signal.emit()