import sys
from PyQt5.QtWidgets import QApplication, QWidget, QComboBox, QSpinBox, QVBoxLayout, QPushButton, QMainWindow, \
    QLabel, QFormLayout, QDoubleSpinBox, QFrame, QGroupBox, QHBoxLayout, QFileDialog, QTableView, QDialog, QHeaderView
from PyQt5.QtSql import QSqlDatabase, QSqlTableModel, QSqlQuery
from class_calculations import SoilLoadCalculator

import xlwt

data = {
    'песок гравелистый плотный': [24, 50, 3, 8.075],
    'песок гравелистый ср': [16, 25, 3, 8.075],
    'песок гравелистый рыхлый': [10, 16, 3, 8.075],
    'песок крупный плотный': [16, 22, 3, 9.69],
    'песок крупный ср': [7, 18, 3, 9.69],
    'песок крупный рыхлый': [1, 5, 3, 9.69],
    'песок средний плотный': [10, 20, 4, 11.305],
    'песок средний ср': [6, 15, 5, 11.305],
    'песок средний рыхлый': [2.5, 5, 6, 11.305],
    'песок мелкий плотный': [5, 12, 7, 12.92],
    'песок мелкий ср': [4, 10, 7.5, 12.92],
    'песок мелкий рыхлый': [1.1, 4, 8, 12.92],
    'песок пылеватый плотный': [10, 12, 15, 20.1875],
    'песок пылеватый ср': [3, 10, 25.5, 24.225],
    'песок пылеватый рыхлый': [1, 3, 30, 29.07],
    'суглинок моренный твердый': [4, 6, 10, 40.375],
    'суглинок моренный полутвердый': [2.5, 4.7, 15, 48.45],
    'суглинок моренный тугопластичный': [2.25, 4.2, 20, 56.525],
    'суглинок моренный мягкопластичный': [1.375, 3.8, 25, 64.6],
    'суглинок моренный текучепластичный': [0.5, 0.8, 30, 72.675],
    'суглинки тяжелые, моренные полутвердые и тугопластичные': [3, 4.9, 10, 62.7],
    'суглинки тяжелые, моренные мягкопластичные': [1, 1.5, 15, 42.75],
    'пылеватые, аллювиальные, слабозаторфованые': [0.8, 1.5, 22, 52],
    'глины полутвердые и твердые': [2.5, 3.5, 50, 80],
    'глины твердые': [4, 6.5, 20, 40],
    'глины мягкопластичные': [1.5, 3.75, 30, 60],
    'юрские суглинки пластичные': [2.5, 3.5, 30, 50],
    'покровный суглинок мягкопластичный': [0.7, 2.9, 50, 90],
    'суглинок флювиогляциальный тугопл': [1, 3.3, 20, 60],
    'суглинок техногенн': [0.7, 5.5, 20, 56.525],
    'супесь пастичная сильнопесчанистая': [0.7, 5.5, 10, 40]
}


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.db = QSqlDatabase.addDatabase("QSQLITE")
        self.db.setDatabaseName("data1.db")
        self.layer_window = None
        self.data_from_db = None
        self.setWindowTitle("Исследование скважин")
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        self.layout = QVBoxLayout()
        self.central_widget.setLayout(self.layout)

        self.num_wells = QSpinBox()
        self.num_wells.setMinimum(1)
        self.num_wells.setMaximum(100)
        self.layout.addWidget(QLabel("Сколько скважин всего?"))
        self.layout.addWidget(self.num_wells)

        self.submit_button = QPushButton("Отправить")
        self.submit_button.clicked.connect(self.showWellData)
        self.layout.addWidget(self.submit_button)

        self.db_button = QPushButton("Редактирование таблицы")
        self.db_button.clicked.connect(self.open_edit_table_dialog)
        self.layout.addWidget(self.db_button)
        # self.data_from_db = self.fetch_data_from_db()
        self.well_groupbox = QGroupBox()
        self.well_layout = QVBoxLayout()  # Use QVBoxLayout for adding QLabel and QSpinBox
        self.well_groupbox.setLayout(self.well_layout)
        self.layout.addWidget(self.well_groupbox)

        self.well_values = []
        self.well_fields = []

    def showWellData(self):
        num_wells = self.num_wells.value()
        self.hide()
        self.num_wells.setDisabled(True)
        for i in range(num_wells):
            well_label = QLabel(f"Скважина {i + 1}")
            num_layers = QSpinBox()
            num_layers.setMinimum(1)
            num_layers.setMaximum(100)

            self.well_fields.append(num_layers)
            self.well_layout.addWidget(well_label)
            self.well_layout.addWidget(num_layers)

        self.submit_button.clicked.disconnect()
        self.submit_button.setText("Отправить данные")
        self.data_from_db = self.fetch_data_from_db()
        self.submit_button.clicked.connect(self.saveWellData)
        self.show()

    def saveWellData(self):
        self.well_values = [spinbox.value() for spinbox in self.well_fields]
        self.showLayerWindow()

    def showLayerWindow(self):
        self.layer_window = LayerWindow(self.well_fields, self.well_values, self.data_from_db)
        self.layer_window.show()
        self.hide()

    def open_edit_table_dialog(self):
        edit_table_dialog = EditTableDialog(parent=self)
        edit_table_dialog.exec_()

    def fetch_data_from_db(self):
        data = {}
        db = QSqlDatabase.addDatabase("QSQLITE")
        db.setDatabaseName("data1.db")
        query = QSqlQuery(db)
        if db.open():
            if query.exec_("SELECT * FROM soil_data_new"):
                while query.next():
                    soil_type = query.value(0)
                    values = [query.value(i) for i in range(1, query.record().count())]
                    data[soil_type] = values
            else:
                print("Ошибка выполнения запроса:", query.lastError().text())

        return data


class EditTableDialog(QDialog):
    def __init__(self, parent=MainWindow):
        super().__init__(parent)
        self.setWindowTitle("Редактирование таблицы")
        self.db = QSqlDatabase.database()
        layout = QVBoxLayout(self)

        self.table_view = QTableView(self)
        layout.addWidget(self.table_view)

        self.model = QSqlTableModel(self)
        self.model.setTable('soil_data_new')
        self.model.select()
        self.table_view.setModel(self.model)
        # Added these two lines
        self.table_view.resizeColumnsToContents()
        table_width = self.table_view.horizontalHeader().length() + self.table_view.verticalHeader().width() + 100
        table_height = self.table_view.verticalHeader().length() + self.table_view.horizontalHeader().height()

        # Установите размер диалогового окна в соответствии с размерами таблицы
        self.resize(table_width, table_height)



        # ... ваш код ...

        # Установите горизонтальное растяжение для о

        button_layout = QVBoxLayout()
        save_button = QPushButton("Сохранить")
        save_button.clicked.connect(self.save_changes)
        button_layout.addWidget(save_button)

        delete_button = QPushButton("Удалить")
        delete_button.clicked.connect(self.delete_row)
        button_layout.addWidget(delete_button)

        add_button = QPushButton("Добавить")
        add_button.clicked.connect(self.add_row)
        button_layout.addWidget(add_button)

        exit_button = QPushButton("Выход")
        exit_button.clicked.connect(self.accept)
        button_layout.addWidget(exit_button)

        layout.addLayout(button_layout)

        # Создание и настройка соединения с базой данных

    def save_changes(self):
        self.model.submitAll()

    def delete_row(self):
        selected_rows = self.table_view.selectionModel().selectedRows()
        for row in selected_rows:
            self.model.removeRow(row.row())
        self.model.submitAll()

    def add_row(self):
        self.model.insertRow(self.model.rowCount())

    def closeEvent(self, event):
        self.save_changes()
        self.db.close()


class LayerWindow(QWidget):
    def __init__(self, well_fields, well_values, data_from_db):
        super().__init__()
        self.data_from_db = data_from_db
        self.soil_types = None
        self.layer_depths = None
        self.setWindowTitle("Информация о слоях")
        self.current_well = 0
        self.well_fields = well_fields
        self.well_values = well_values
        self.current_well_window = None
        self.finish_window = None
        self.ugv_depths = None
        self.ugv_values_data = []
        self.well_data = {}
        for well_number, num_layers in enumerate(well_fields,
                                                 start=0):  # Измените значения грунта и глубины слоя по своему усмотрению
            self.well_data[well_number] = []
        for _ in enumerate(well_fields, start=0):
            self.ugv_values_data.append(0)
        self.layout = QFormLayout()
        self.setLayout(self.layout)

        self.next_button = QPushButton("Далее")
        self.back_button = QPushButton("Назад")
        self.finish_button = QPushButton("Завершить")

        self.next_button.clicked.connect(self.showNextWellWindow)
        self.back_button.clicked.connect(self.showPreviousWellWindow)
        self.finish_button.clicked.connect(self.showFinishWindow)

        self.buttons_layout = QHBoxLayout()
        self.buttons_layout.addWidget(self.back_button)
        self.buttons_layout.addWidget(self.next_button)

        self.layout.addRow(self.buttons_layout)

        self.createCurrentWellWindow()

    def createCurrentWellWindow(self):
        if self.current_well_window:
            self.layout.removeWidget(self.current_well_window)
            self.current_well_window.deleteLater()

        num_wells = len(self.well_fields)
        if self.current_well < 0 or self.current_well >= num_wells:
            return

        well_number = self.current_well + 1
        num_layers = self.well_fields[self.current_well].value()

        self.current_well_window = QWidget()
        current_well_layout = QFormLayout()
        self.current_well_window.setLayout(current_well_layout)

        well_label = QLabel(f"Скважина {well_number}")
        current_well_layout.addRow(well_label)

        # Create lists to store the combo boxes and spin boxes
        self.soil_types = []
        self.layer_depths = []
        self.ugv_depths = []
        for i in range(num_layers):
            layer_label = QLabel(f"Слой {i + 1}:")
            soil_type = QComboBox()
            soil_type.addItems(self.data_from_db.keys())
            layer_depth = QDoubleSpinBox()
            layer_depth.setMinimum(0)
            layer_depth.setMaximum(100)
            layer_depth.setDecimals(1)
            current_well_layout.addRow(layer_label, soil_type)
            current_well_layout.addRow(QLabel("Тип грунта:"), soil_type)
            current_well_layout.addRow(QLabel("Глубина слоя:"), layer_depth)

            # Append combo boxes and spin boxes to the lists
            self.soil_types.append(soil_type)
            self.layer_depths.append(layer_depth)
        # Append combo boxes and spin boxes to the lists
        ugv_depth_spin = QDoubleSpinBox()
        ugv_depth_spin.setMinimum(0)
        ugv_depth_spin.setMaximum(100)
        ugv_depth_spin.setDecimals(1)
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        current_well_layout.addRow(line)
        current_well_layout.addRow(ugv_depth_spin)
        current_well_layout.addRow(QLabel("УГВ:"), ugv_depth_spin)
        self.ugv_depths.append(ugv_depth_spin)
        self.layout.insertRow(0, self.current_well_window)

        if self.current_well == num_wells - 1:
            self.buttons_layout.addWidget(self.finish_button)

        else:
            self.buttons_layout.addWidget(self.next_button)

        if self.current_well == 0:
            self.back_button.setVisible(False)
        else:
            self.back_button.setVisible(True)

        if self.current_well == num_wells - 1:
            self.next_button.setVisible(False)
        else:
            self.next_button.setVisible(True)

        if self.current_well == num_wells - 1:
            self.finish_button.setVisible(True)
        else:
            self.finish_button.setVisible(False)
        self.openCurrentWellData()

    def openCurrentWellData(self):
        if self.current_well in self.well_data:
            well_data = self.well_data[self.current_well]
            num_layers = self.well_fields[self.current_well].value()

            for i in range(num_layers):
                if i < len(well_data):
                    soil_type, layer_depth = well_data[i]
                    self.soil_types[i].setCurrentText(soil_type)
                    self.layer_depths[i].setValue(layer_depth)
            self.ugv_depths[0].setValue(self.ugv_values_data[self.current_well])

    def showNextWellWindow(self):
        self.saveCurrentWellData()
        self.current_well += 1
        self.createCurrentWellWindow()

    def showPreviousWellWindow(self):
        self.saveCurrentWellData()
        self.current_well -= 1
        self.createCurrentWellWindow()

    def saveCurrentWellData(self):
        well_number = self.current_well
        num_layers = self.well_fields[well_number].value()
        self.well_data[well_number].clear()
        ugv_depth = self.ugv_depths[0].value()
        for i in range(num_layers):
            soil_type = self.soil_types[i].currentText()
            layer_depth = self.layer_depths[i].value()

            if well_number not in self.well_data:
                self.well_data[well_number] = []

            self.well_data[well_number].append((soil_type, layer_depth))

        self.ugv_values_data[well_number] = ugv_depth

        print(f"Well {well_number + 1} data:", self.well_data[well_number])
        print(f"Well {well_number + 1} data:", self.ugv_values_data)

    def showFinishWindow(self):
        print(self.well_data)
        self.saveCurrentWellData()
        self.finish_window = FinishWindow(self.well_data, self.ugv_values_data, self.data_from_db)
        self.finish_window.show()
        self.hide()


class FinishWindow(QWidget):
    def __init__(self, well_data, ugv_values_data, data_from_db):
        super().__init__()
        self.depthWindow = None
        self.dataframes = None
        self.setWindowTitle("Завершение")
        self.layout = QVBoxLayout()
        self.setLayout(self.layout)

        self.layout.addWidget(QLabel("Исследование скважин завершено."))
        self.finish_button = QPushButton("Завершить")
        self.layout.addWidget(self.finish_button)
        self.finish_button.clicked.connect(self.showDepthRangeWindow)
        soil_calculator = SoilLoadCalculator()
        for borehole_num in range(len(well_data)):
            soil_calculator.generate_data(borehole_num, data_from_db, well_data)
            soil_calculator.process_data(ugv_values_data, borehole_num)
        self.dataframes = soil_calculator.df_list

    def showDepthRangeWindow(self):
        self.depthWindow = DepthRangeWindow(self.dataframes)
        self.depthWindow.show()
        self.hide()


class DepthRangeWindow(QMainWindow):
    def __init__(self, dataframes):
        super().__init__()

        self.folder_path = ''
        self.dataframes = dataframes
        self.range_start_spinboxes = []
        self.range_end_spinboxes = []
        self.setWindowTitle("Выбор диапазона глубины")

        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        for i, dataframe in enumerate(dataframes):
            well_label = QLabel(f"Скважина: {i + 1}", self)
            layout.addWidget(well_label)

            start_spinbox = QDoubleSpinBox(self)
            end_spinbox = QDoubleSpinBox(self)
            start_spinbox.setDecimals(1)
            end_spinbox.setDecimals(1)
            start_row = QWidget()
            start_layout = QHBoxLayout(start_row)
            start_layout.addWidget(QLabel("Начало:"))
            start_layout.addWidget(start_spinbox)
            layout.addWidget(start_row)

            end_row = QWidget()
            end_layout = QHBoxLayout(end_row)
            end_layout.addWidget(QLabel("Конец:"))
            end_layout.addWidget(end_spinbox)
            layout.addWidget(end_row)

            self.range_start_spinboxes.append(start_spinbox)
            self.range_end_spinboxes.append(end_spinbox)

        load_button = QPushButton("Загрузить данные", self)
        layout.addWidget(load_button)
        load_button.clicked.connect(self.load_data)

        folder_button = QPushButton("Выбрать папку", self)
        layout.addWidget(folder_button)
        folder_button.clicked.connect(self.select_folder)

    def load_data(self):
        for i, df in enumerate(self.dataframes):
            range_start = self.range_start_spinboxes[i].value()
            range_end = self.range_end_spinboxes[i].value()

            filtered_data = df[
                (df["Глубина"] >= range_start) &
                (df["Глубина"] <= range_end)
                ]
            column_mapping = {
                'Глубина': 'Глубина',
                'Округленное_Среднее_Гаусса': 'СР1',
                'Округленный_Средний_Нов.Бок.Лоб': 'СР2'
            }
            if self.folder_path == '':
                self.save_to_xls(filtered_data, f"skvazhina_{i + 1}.xls", column_mapping)
            else:
                self.save_to_xls(filtered_data, f"{self.folder_path}/skvazhina_{i + 1}.xls",
                                 column_mapping)  # Сохраняем
            # отфильтрованный датафрейм

    def save_to_xls(self, df, filename, column_mapping=None):
        if column_mapping is None:
            column_mapping = df.columns.tolist()

        workbook = xlwt.Workbook()
        worksheet = workbook.add_sheet('Sheet1')

        # Записываем новые заголовки
        for col_idx, (old_name, new_name) in enumerate(column_mapping.items()):
            worksheet.write(0, col_idx, new_name)

        # Записываем данные
        for row_idx, (_, row) in enumerate(df.iterrows(), start=1):
            for col_idx, (old_name, _) in enumerate(column_mapping.items()):
                worksheet.write(row_idx, col_idx, row[old_name])

        workbook.save(filename)

    def select_folder(self):
        options = QFileDialog.Options()
        self.folder_path = QFileDialog.getExistingDirectory(self, "Выберите папку", options=options)
        if self.folder_path:
            print("Выбрана папка:", self.folder_path)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()
    sys.exit(app.exec_())
