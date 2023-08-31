import win32com.client
import pandas as pd
import random
import xlrd
from xlutils.copy import copy


class WellLayerSampling:
    def __init__(self, excel_file):
        self.excel_file = excel_file
        self.df = pd.read_excel(self.excel_file)
        self.layer_boundaries = {}
        self.total_depths = {}
        self.num_samples = {}
        self.samples = pd.DataFrame(columns=['Номер пробы', 'Номер скважины', 'Глубина отбора пробы ОТ',
                                             'Глубина отбора пробы ДО', 'Номер слоя'])
        self.layer_residuals = {}
        self.id_sample = []

    def load_data(self):
        self.df = pd.read_excel(self.excel_file)

    def find_layer_boundaries(self):
        for index, row in self.df.iterrows():
            well_number = int(row[0])
            layer_number = int(row[1])
            sole_depth = row[2]

            if well_number not in self.layer_boundaries:
                self.layer_boundaries[well_number] = []

            self.layer_boundaries[well_number].append({'layer_number': layer_number, 'top': None, 'bottom': sole_depth})

    def calculate_total_depths(self):
        for well, layers in self.layer_boundaries.items():
            # Сортировка слоев по глубине подошвы
            layers = sorted(layers, key=lambda x: x['bottom'])
            previous_bottom = 0

            for ind, layer_info in enumerate(layers):
                if ind > 0:
                    # Устанавливаем границу верхней поверхности текущего слоя на основе границы подошвы предыдущего слоя
                    layer_info['top'] = previous_bottom
                else:
                    # Если это первый слой, устанавливаем границу верхней поверхности на 0
                    layer_info['top'] = 0
                previous_bottom = layer_info['bottom']

    def count_unique_layers(self):
        count_layers = self.df.iloc[:, 1].unique().tolist()
        return count_layers

    def input_num_samples(self):
        count_layers = self.count_unique_layers()
        for i in count_layers:
            print(f"Введите кол-во для слоя {i} по умолчанию 10")
            k_test = 20
            self.num_samples[i] = k_test

    @staticmethod
    def round_dictionary_values(dictionary):
        rounded_dict = {}
        for key, val in dictionary.items():
            rounded_value = round(val)
            rounded_dict[key] = rounded_value
        return rounded_dict

    def calculate_layer_depths(self):
        zero_top_count = {}
        wells_per_layer = {}
        for well, layers in self.layer_boundaries.items():
            for layer_info in layers:
                layer_depth = layer_info['bottom'] - layer_info['top']

                layer_number = layer_info['layer_number']
                if layer_number not in self.total_depths:
                    self.total_depths[layer_number] = 0
                if layer_number not in wells_per_layer:
                    wells_per_layer[layer_number] = 0
                wells_per_layer[layer_number] += 1

                self.total_depths[layer_number] += layer_depth
        for well, layers in self.layer_boundaries.items():
            for layer_info in layers:
                layer_number = layer_info['layer_number']
                bottom = layer_info['bottom'] - 0.2
                top = layer_info['top']
                layer_depth = bottom - top
                if layer_number not in zero_top_count:
                    zero_top_count[layer_number] = 0
                if top == 0:
                    zero_top_count[layer_number] += 1

                denominator = self.total_depths[layer_number] - (wells_per_layer[layer_number] * 0.2)
                num_samples_layer = round(self.num_samples[layer_number] * (layer_depth / denominator), 1)
                num_samples_layer_rounded = max(round(num_samples_layer), 1)
                layer_info['sample_per_interval'] = num_samples_layer_rounded
        print(self.layer_boundaries)
        # self.total_depths = self.round_dictionary_values(self.total_depths)

    def calculate_samples_info(self):
        samples_per_layer = {}

        for well, layers in self.layer_boundaries.items():
            for layer_info in layers:
                layer_number = layer_info['layer_number']
                samples_per_interval = layer_info['sample_per_interval']

                if layer_number not in samples_per_layer:
                    samples_per_layer[layer_number] = 0

                samples_per_layer[layer_number] += samples_per_interval

        # Проходим по слоям и вычитаем 1 из интервала с максимальным количеством проб на слое
        for well, layers in self.layer_boundaries.items():
            for layer_info in layers:
                layer_number = layer_info['layer_number']

                if samples_per_layer[layer_number] > self.num_samples[layer_number]:
                    max_interval = max([layer for layers in self.layer_boundaries.values() for layer in layers if
                                        layer['layer_number'] == layer_number], key=lambda x: x['sample_per_interval'])
                    max_interval['sample_per_interval'] -= 1
                    samples_per_layer[layer_number] -= 1
                elif samples_per_layer[layer_number] < self.num_samples[layer_number]:
                    min_interval = min([layer for layers in self.layer_boundaries.values() for layer in layers if
                                        layer['layer_number'] == layer_number], key=lambda x: x['sample_per_interval'])
                    min_interval['sample_per_interval'] += 1
                    samples_per_layer[layer_number] += 1

        return self.layer_boundaries

    @staticmethod
    def round_list_values(lst, num_decimals):
        rounded_list = [round(value, num_decimals) for value in lst]
        return rounded_list

    def generate_incremental_random_numbers(self, n, min_value, max_value):
        if n <= 0 or min_value >= max_value:
            return []

        step = (max_value - min_value) / n
        min_distance = 0.2
        if step < 0.2:
            min_distance = 0.1
        steps = [min_value+ step * i for i in range(1, n + 1)]
        numbers = []
        steps.insert(0, min_value)
        for i in range(len(steps) - 1):
            new_number = random.uniform(steps[i], steps[i + 1])
            rounded_number = round(new_number * 10) / 10
            if numbers:
                while rounded_number - numbers[-1] < min_distance - 0.0000000001:
                    new_number = random.uniform(steps[i], steps[i + 1])
                    rounded_number = round(new_number * 10) / 10
            numbers.append(rounded_number)

        return numbers

    def generate_samples(self, sample_number):
        self.id_sample = list(range(sample_number, sample_number + sum(self.num_samples.values())))
        random.shuffle(self.id_sample)
        ind = 0
        min_distance = 0.2
        sample_info = self.calculate_samples_info()
        print(sample_info)

        for well, layers in self.layer_boundaries.items():
            for layer_info in layers:
                layer_number = layer_info['layer_number']
                boundaries = layer_info
                top = boundaries['top']
                bottom = boundaries['bottom']
                num_samples_layer = layer_info['sample_per_interval']
                bottom = bottom - 0.2
                depths = []

                if top < 0.2:
                    top = 0.2

                if num_samples_layer != 1:
                    depths = self.generate_incremental_random_numbers(num_samples_layer, top, bottom)
                else:
                    depths.append(random.uniform(top, bottom))

                depths = self.round_list_values(depths, 1)

                for i, depth in enumerate(depths):
                    if i >= num_samples_layer or ind >= len(self.id_sample):
                        break
                    sample_data = {'Номер пробы': self.id_sample[ind], 'Номер скважины': well,
                                   'Глубина отбора пробы ОТ': depth, 'Глубина отбора пробы ДО': depth + 0.2,
                                   'Номер слоя': layer_number}
                    self.samples = pd.concat([self.samples, pd.DataFrame(sample_data, index=[0])], ignore_index=True)
                    ind += 1

        return self.samples

    # Пример использования  # Return the generated samples

    def save_samples_to_excel(self, template_file, output_file):
        table = self.samples.sort_values(by=['Номер слоя', 'Номер пробы', 'Глубина отбора пробы ОТ'])
        book = xlrd.open_workbook(template_file, formatting_info=True)
        new_book = copy(book)
        new_sheet = new_book.get_sheet(0)
        start_row = 10
        start_column = 0

        for i, row in enumerate(table.itertuples(), start_row):
            for j, value in enumerate(row[1:], start_column):
                new_sheet.write(i, j, value)
        start_row = 7
        new_sheet = new_book.get_sheet(1)
        for i, row in enumerate(table.itertuples(), start_row):
            new_sheet.write(i, start_column, row[1])  # Замените на имя столбца
            new_sheet.write(i, start_column + 1, row[2])  # Замените на имя столбца
            new_sheet.write(i, start_column + 2, row[5])
        # Replace with the actual named range formula
        new_book.save(output_file)
        xlApp = win32com.client.Dispatch("Excel.Application")
        xlApp.Visible = False  # Не показывать Excel

        workbook = xlApp.Workbooks.Open(output_file)
        worksheet = workbook.Worksheets(1)  # Первый лист

        name = "info"
        refers_to = "=Лист1!$A$10:$BM$1463"  # Измените имя листа
        workbook.Names.Add(Name=name, RefersTo=refers_to)

        workbook.Save()
        xlApp.Quit()


# sampling = WellLayerSampling('H:\My_Projects\Fed_project\Wells_calculations_for_tables\output_Объект.xlsx')
# sampling.load_data()
# sampling.find_layer_boundaries()
# sampling.calculate_total_depths()
# sampling.input_num_samples()
# sampling.calculate_layer_depths()
# number_samples = 1111
# sampling.generate_samples(number_samples)
# sampling.save_samples_to_excel('template.xls', 'H:\My_Projects\Fed_project\Wells_calculations_for_tables\kolskaya1.xls')
