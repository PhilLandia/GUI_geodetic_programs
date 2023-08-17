import pandas as pd
import numpy as np
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
                self.layer_boundaries[well_number] = {}

            if layer_number not in self.layer_boundaries[well_number]:
                self.layer_boundaries[well_number][layer_number] = {'top': None, 'bottom': sole_depth}
            else:
                self.layer_boundaries[well_number][layer_number]['bottom'] = sole_depth

    def calculate_total_depths(self):
        for well in self.layer_boundaries:
            layers = self.layer_boundaries[well]

            # Сортировка словаря слоев по значению (глубине подошвы)
            layers = dict(sorted(layers.items(), key=lambda x: x[1]['bottom']))
            ind = 0
            previous_bottom = 0
            for layer in layers:
                ind += 1
                boundaries = layers[layer]

                if ind > 1:
                    # Устанавливаем границу верхней поверхности текущего слоя на основе границы подошвы предыдущего слоя
                    boundaries['top'] = previous_bottom
                else:
                    # Если это первый слой, устанавливаем границу верхней поверхности на 0
                    boundaries['top'] = 0
                previous_bottom = boundaries['bottom']

    def count_unique_layers(self):
        count_layers = self.df.iloc[:, 1].unique().tolist()
        return count_layers

    def input_num_samples(self):
        count_layers = self.count_unique_layers()
        for i in count_layers:
            print(f"Введите кол-во для слоя {i} по умолчанию 10")
            k_test = int(input())
            self.num_samples[i] = k_test

    @staticmethod
    def round_dictionary_values(dictionary):
        rounded_dict = {}
        for key, val in dictionary.items():
            rounded_value = round(val)
            rounded_dict[key] = rounded_value
        return rounded_dict

    def calculate_layer_depths(self):
        for well in self.layer_boundaries:
            layers = self.layer_boundaries[well]

            for layer in layers:
                boundaries = layers[layer]
                layer_depth = boundaries['bottom'] - boundaries['top']

                if layer not in self.total_depths:
                    self.total_depths[layer] = 0

                self.total_depths[layer] += layer_depth

        self.total_depths = self.round_dictionary_values(self.total_depths)

    def calculate_layer_residuals(self):
        for well in self.layer_boundaries:
            layers = self.layer_boundaries[well]

            for layer in layers:
                boundaries = layers[layer]
                layer_depth = boundaries['bottom'] - boundaries['top']

                if layer not in self.layer_residuals:
                    self.layer_residuals[layer] = 0

                num_samples_layer = self.num_samples[layer] * (layer_depth / self.total_depths[layer])
                num_samples_layer_rounded = round(self.num_samples[layer] * (layer_depth / self.total_depths[layer]))
                residual = num_samples_layer - num_samples_layer_rounded

                self.layer_residuals[layer] += residual

        self.layer_residuals = self.round_dictionary_values(self.layer_residuals)

    def generate_samples(self, sample_number):
        self.id_sample = list(range(sample_number, sample_number + sum(self.num_samples.values())))
        random.shuffle(self.id_sample)
        ind = 0

        for well in self.layer_boundaries:
            layers = self.layer_boundaries[well]

            for layer in layers:
                boundaries = layers[layer]
                top = boundaries['top']
                bottom = boundaries['bottom']
                layer_depth = bottom - top
                num_samples_layer_rounded = round(self.num_samples[layer] * (layer_depth / self.total_depths[layer]))

                if self.layer_residuals[layer] > 0:
                    num_samples_layer_rounded += 1
                    self.layer_residuals[layer] -= 1
                if self.layer_residuals[layer] < 0:
                    num_samples_layer_rounded -= 1
                    self.layer_residuals[layer] += 1

                depths = []
                min_distance = 0.2

                while len(depths) < num_samples_layer_rounded:
                    depth = np.random.uniform(low=top, high=bottom - min_distance)
                    depth = round(depth, 1)

                    is_valid = True
                    for existing_depth in depths:
                        if abs(depth - existing_depth) < min_distance - 0.01:
                            is_valid = False
                            break

                    if is_valid:
                        depths.append(depth)

                for i, depth in enumerate(depths):
                    if i >= self.num_samples[layer] or ind >= len(self.id_sample):
                        break
                    sample_data = {'Номер пробы': self.id_sample[ind], 'Номер скважины': well,
                                   'Глубина отбора пробы ОТ': depth, 'Глубина отбора пробы ДО': depth + min_distance,
                                   'Номер слоя': layer}
                    self.samples = pd.concat([self.samples, pd.DataFrame(sample_data, index=[0])], ignore_index=True)
                    ind += 1

                    if ind >= len(self.id_sample):
                        break

            if ind >= len(self.id_sample):
                break

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

        new_book.save(output_file)


# sampling = WellLayerSampling('file2.xls')
# sampling.load_data()
# sampling.find_layer_boundaries()
# sampling.calculate_total_depths()
# sampling.input_num_samples()
# sampling.calculate_layer_depths()
# sampling.calculate_layer_residuals()
# number_samples = int(input("Введите номер начальной пробы: \n"))
# sampling.generate_samples(number_samples)
# sampling.save_samples_to_excel('template.xls', 'kolskaya1.xls')