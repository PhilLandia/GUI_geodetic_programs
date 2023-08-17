import sys
from PyQt5.QtWidgets import QDialog, QApplication, QMainWindow, QLabel, QVBoxLayout, QWidget, QPushButton, QFileDialog, \
    QLineEdit, QSpinBox
from well_layer_sampling import WellLayerSampling


class LayerSamplesDialog(QDialog):
    def __init__(self, layers):
        super().__init__()
        self.setWindowTitle("Input Number of Samples")
        self.layout = QVBoxLayout()
        self.setLayout(self.layout)
        self.value_first_sample = int()
        self.num_samples = {}
        self.spin = QSpinBox(self)
        for layer in layers:
            label = QLabel(f"Слой {layer}")
            spin_edit = QSpinBox()
            spin_edit.setValue(10)  # Default value
            self.layout.addWidget(label)
            self.layout.addWidget(spin_edit)

            self.num_samples[layer] = spin_edit
        label_first_sample = QLabel("Номер первой пробы:")
        self.spin.setRange(1000,9999)
        self.spin.setValue(1000)
        self.layout.addWidget(label_first_sample)
        self.layout.addWidget(self.spin)
        self.value_first_sample = self.spin.value
        button = QPushButton("OK")
        button.clicked.connect(self.accept)
        self.layout.addWidget(button)

    def get_num_samples(self):
        return {layer: int(line_edit.text()) for layer, line_edit in self.num_samples.items()}

    def get_first_sample(self):
        return self.spin.value()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Well Layer Sampling")
        self.layout = QVBoxLayout()
        self.central_widget = QWidget()
        self.central_widget.setLayout(self.layout)
        self.setCentralWidget(self.central_widget)

        self.file_label = QLabel("Excel File:")
        self.layout.addWidget(self.file_label)

        self.file_button = QPushButton("Select File")
        self.file_button.clicked.connect(self.select_file)
        self.layout.addWidget(self.file_button)

        self.template_label = QLabel("Template File:")
        self.layout.addWidget(self.template_label)

        self.template_button = QPushButton("Select Template")
        self.template_button.clicked.connect(self.select_template)
        self.layout.addWidget(self.template_button)

        self.output_label = QLabel("Output File:")
        self.layout.addWidget(self.output_label)

        self.output_button = QPushButton("Select Output")
        self.output_button.clicked.connect(self.select_output)
        self.layout.addWidget(self.output_button)

        self.generate_button = QPushButton("Generate Samples")
        self.generate_button.clicked.connect(self.generate_samples)
        self.layout.addWidget(self.generate_button)

        self.sampling = None
        self.excel_file = None
        self.template_file = None
        self.output_file = None

    def select_file(self):
        file_dialog = QFileDialog(self)
        file_dialog.setNameFilter("Excel Files (*.xls *.xlsx)")
        if file_dialog.exec_():
            self.excel_file = file_dialog.selectedFiles()[0]

    def select_template(self):
        file_dialog = QFileDialog(self)
        file_dialog.setNameFilter("Excel Files (*.xls *.xlsx)")
        if file_dialog.exec_():
            self.template_file = file_dialog.selectedFiles()[0]

    def select_output(self):
        file_dialog = QFileDialog(self)
        file_dialog.setNameFilter("Excel Files (*.xls *.xlsx)")
        file_dialog.setDefaultSuffix("xls")
        file_dialog.setAcceptMode(QFileDialog.AcceptSave)
        if file_dialog.exec_():
            self.output_file = file_dialog.selectedFiles()[0]

    def generate_samples(self):
        if self.excel_file and self.template_file and self.output_file:
            self.sampling = WellLayerSampling(self.excel_file)
            self.sampling.load_data()
            self.sampling.find_layer_boundaries()
            self.sampling.calculate_total_depths()
            layers = self.sampling.count_unique_layers()
            dialog = LayerSamplesDialog(layers)
            if dialog.exec_() == QDialog.Accepted:
                num_samples = dialog.get_num_samples()
                number_samples = dialog.get_first_sample()
                self.sampling.num_samples = num_samples
                self.sampling.calculate_layer_depths()
                self.sampling.calculate_layer_residuals()

                self.sampling.generate_samples(number_samples)

                self.sampling.save_samples_to_excel(self.template_file, self.output_file)
                sys.exit(0)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())