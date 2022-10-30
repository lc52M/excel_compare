import os
import sys

import pandas as pd

from excel_compare import Ui_MainWindow

from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox
from PyQt5.QtCore import QStringListModel


class Main(QMainWindow, Ui_MainWindow):
    def __init__(self):
        QMainWindow.__init__(self)
        Ui_MainWindow.__init__(self)
        self.setupUi(self)

        # 主要关联列
        self.major_relation_column = ''

        # 设置待比较列数据
        self.to_compare_columns_data = []
        self.to_compare_columns_data_model = QStringListModel()
        self.to_compare_columns_data_model.setStringList(self.to_compare_columns_data)

        # 设置比较列数据
        self.compare_columns_data = []
        self.compare_columns_data_model = QStringListModel()
        self.compare_columns_data_model.setStringList(self.compare_columns_data)

        # 设置差异列数据
        self.difference_columns_data = []
        self.difference_columns_data_model = QStringListModel()
        self.difference_columns_data_model.setStringList(self.difference_columns_data)

        # 设置按钮事件
        self.compare_file_button.clicked.connect(lambda: self.choose_excel_file(self.compare_file, True))
        self.to_compare_file_button.clicked.connect(lambda: self.choose_excel_file(self.to_compare_file, False))
        self.save_file_path_button.clicked.connect(lambda: self.open_dir(self.save_file_path))
        self.compare_columns_view.clicked.connect(
            lambda: self.operation_item(self.compare_columns_view.currentIndex(), False, self.compare_columns_view,
                                        self.compare_columns_data_model, self.compare_columns_data,
                                        self.to_compare_columns_view, self.to_compare_columns_data_model,
                                        self.to_compare_columns_data))

        self.to_compare_columns_view.clicked.connect(
            lambda: self.operation_item(self.to_compare_columns_view.currentIndex(), True, self.compare_columns_view,
                                        self.compare_columns_data_model, self.compare_columns_data,
                                        self.to_compare_columns_view, self.to_compare_columns_data_model,
                                        self.to_compare_columns_data))
        self.to_compare_columns_view.doubleClicked.connect(
            lambda: self.choose_compare_major_column(self.to_compare_columns_data,
                                                     self.to_compare_columns_view.currentIndex(), self.major_column))

        self.start_compare_button.clicked.connect(lambda: self.start_compare(
            self.compare_file.text(), self.to_compare_file.text(), self.save_file_path.text(),
            self.compare_columns_data, self.major_column.text()))

    def set_major_relation_column(self, major_relation_column):
        self.major_relation_column = major_relation_column

    def choose_excel_file(self, file_input, major_file_flag):
        """
        选择 excel 文件
        :param file_input: 文件输入框
        :param major_file_flag: 主文件标识
        :return:
        """
        excel_file, _ = QFileDialog.getOpenFileName(self, '选择excel文件', './', 'excel(*.xls *.xlsx)')
        excel_data_columns = self.read_excel_file_columns(excel_file)
        if excel_data_columns:
            if len([column for column in excel_data_columns if 'Unnamed' in column]) > 0:
                QMessageBox.information(None, "excel 数据列表不合法",
                                        "excel 数据列表有未知列,建议不要有表头或者补充未知列")
                return
            # 设置文件输入框数据
            file_input.setText(excel_file)
            # 设置差异列表数据
            self.difference_columns_data = set(excel_data_columns).difference(set(self.difference_columns_data))
            self.view_fill_data(self.difference_columns_view, self.difference_columns_data_model,
                                self.difference_columns_data)
            if major_file_flag:
                # 设置比较列表数据
                self.to_compare_columns_data = excel_data_columns
                self.view_fill_data(self.to_compare_columns_view, self.to_compare_columns_data_model,
                                    self.to_compare_columns_data)
        else:
            QMessageBox.information(None, "excel 文件不合法", "请选择正确的 excel 文件")

    def view_fill_data(self, view, model, data):
        """
        试图填充数据
        :param view:  试图
        :param model: 模型
        :param data: 数据
        :return:
        """
        model.setStringList(data)
        view.setModel(model)

    def read_excel_file_columns(self, excel_file):
        """
        读取 excel 文件
        :param excel_file: excel 文件
        :return:
        """
        try:
            excel_data = pd.read_excel(excel_file)
            return excel_data.columns.values.tolist()
        except Exception as ex:
            QMessageBox.information(None, "excel 文件解析出错了", ex)
            return None

    def open_dir(self, file_path):
        """
        打开目录
        :return:
        """
        file_path.setText(QFileDialog.getExistingDirectory(self, '选择文件存储路径', './'))

    def operation_item(self, index, operation_flag, compare_columns_view, compare_columns_data_model,
                       compare_columns_data,
                       to_compare_columns_view, to_compare_columns_data_model, to_compare_columns_data):
        """
        操作项
        :param index: 选择项索引
        :param operation_flag: 操作标识【True:比较视图添加项】【False:待比较视图添加项】
        :param compare_columns_view: 比较列视图
        :param compare_columns_data_model: 比较列数据模型
        :param compare_columns_data: 比较列数据
        :param to_compare_columns_view: 待比较列视图
        :param to_compare_columns_data_model: 待比较列数据模型
        :param to_compare_columns_data: 待比较列数据
        :return:
        """
        if operation_flag:
            item = to_compare_columns_data[index.row()]
            to_compare_columns_data.remove(item)
            compare_columns_data.append(item)
        else:
            item = compare_columns_data[index.row()]
            compare_columns_data.remove(item)
            to_compare_columns_data.append(item)

        self.view_fill_data(compare_columns_view, compare_columns_data_model, compare_columns_data)
        self.view_fill_data(to_compare_columns_view, to_compare_columns_data_model,
                            to_compare_columns_data)

    def choose_compare_major_column(self, compare_columns_data, index, major_column):
        """
        选择主要对比列
        :param compare_columns_data: 对比列数据
        :param index: 选择项索引
        :param major_column: 主要列
        :return:
        """
        item = compare_columns_data[index.row()]
        major_column.setText(item)
        self.set_major_relation_column(item)

    def generate_file_name(self, compare_file_path, to_compare_file_path):
        """
        生成文件名
        :param compare_file_path: 比较文件路径
        :param to_compare_file_path: 待比较文件路径
        :return:
        """
        compare_file_name, compare_file_name_ext = os.path.splitext(compare_file_path)
        to_compare_file_name, to_compare_file_name_ext = os.path.splitext(to_compare_file_path)
        return self.get_file_name(compare_file_name) + "对比" + self.get_file_name(
            to_compare_file_name) + "结果" + compare_file_name_ext

    def get_file_name(self, file_path):
        """
        获取文件名
        :param file_path: 文件路径
        :param ext: 文件扩展名
        :return:
        """
        return file_path[file_path.find(file_path.split('/')[-1]):]

    def start_compare(self, compare_excel_file, to_compare_excel_file, save_file_path, compare_columns_data,
                      major_column):
        """
        开始比对
        :param compare_excel_file: 对比 excel 文件
        :param to_compare_excel_file: 待对比 excel 文件
        :param save_file_path: 保存文件路径
        :param compare_columns_data: 比较列数据
        :param major_column: 关键列
        :return:
        """
        try:
            if all([compare_excel_file, to_compare_excel_file, save_file_path, compare_columns_data, major_column]):
                compare_excel_data = pd.read_excel(compare_excel_file)
                to_compare_excel_data = pd.read_excel(to_compare_excel_file)
                to_compare_excel_copy_data = to_compare_excel_data.copy()
                for column in compare_columns_data:
                    for item in range(len(to_compare_excel_data[major_column])):
                        column_data = compare_excel_data[
                            compare_excel_data[major_column] == to_compare_excel_data[major_column][item]][column]
                        result = self.compare_result(to_compare_excel_data[column][item], column_data.iloc[0])
                        to_compare_excel_copy_data.loc[item, column] = result

                save_file = save_file_path + "/" + self.generate_file_name(compare_excel_file, to_compare_excel_file)
                to_compare_excel_copy_data.to_excel(save_file, index=None)
                QMessageBox.information(None, "对比完成",
                                        "请查看对比文件结果 <a href=" + save_file + ">" + save_file + "</a>")
            else:
                QMessageBox.information(None, "入参不正确", "请确认红色字体参数是否正确填写")
        except:
            QMessageBox.information(None, "比对出错了", "换个文件再试一次")

    def compare_result(self, current_score, compare_score):
        """
        对比结果
        :param current_score: 当前分数
        :param compare_score: 比较分数
        :return:
        """
        return str(current_score) + "(" + str(current_score - compare_score) + ")"


if __name__ == '__main__':
    app = QApplication(sys.argv)
    win = Main()
    win.show()
    sys.exit(app.exec())
