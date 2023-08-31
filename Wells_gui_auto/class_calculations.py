import random as rand
import pandas as pd
from scipy.stats import norm


class SoilLoadCalculator:
    def __init__(self):
        self.df_list = []  # List to store generated DataFrames
        self.mean_values = None
        self.std_values = None

    def friction_coefficient(self, depth):
        if depth <= 2:
            return 0.5 + (depth / 2) * 0.5
        elif 2 < depth <= 10:
            return 1
        elif 10 < depth <= 60:
            return 1 + (depth - 10) / 50
        else:
            return 2

    def generate_data(self, borehole_num, data, samples):
        data_list = []
        nach_lob = []
        bok = []

        j = 0
        for i in samples[borehole_num]:
            while round(j, 2) <= i[1]:
                key = i[0]

                p = data[key]
                layer_name = p[4]
                nach_lob.append((p[1] - p[0]) * rand.random() + p[0])
                bok.append((p[3] - p[2]) * rand.random() + p[2])
                row = {'Слой': key,'Имя_слоя': layer_name, 'Глубина': round(j, 2), 'Нач.лоб': nach_lob[-1],
                       'Бок.лоб': bok[-1], 'Коэф_Глубины': self.friction_coefficient(j)}
                data_list.append(row)
                j += 0.05
        nach_lob.clear()

        borehole_df = pd.DataFrame(data_list)
        borehole_df['Скважина'] = borehole_num
        self.df_list.append(borehole_df)  # Append the generated DataFrame to the list

    def calculate_statistics(self, number):
        grouped = self.df_list[number].groupby('Слой')  # Use 'Слой' instead of 'Cлой'
        self.mean_values = grouped['Нач.лоб'].mean()
        self.std_values = grouped['Нач.лоб'].std()

    def calculate_probability(self, column):
        layer_mean = self.mean_values[column['Слой']]  # Use 'Слой' instead of 'Cлой'
        layer_std = self.std_values[column['Слой']]  # Use 'Слой' instead of 'Cлой'
        return norm.ppf(rand.random(), loc=layer_mean, scale=layer_std) * column['Коэф_Глубины']

    def calculate_probability_bok(self, column):
        return column['Гаусса'] * column['Бок.лоб']

    def calculate_soil_load(self, number: int) -> None:
        df_hole = self.df_list[number]
        df_hole['Гаусса'] = df_hole.apply(lambda row: self.calculate_probability(row), axis=1)
        df_hole['Нов.Бок.Лоб'] = df_hole.apply(self.calculate_probability_bok, axis=1)
        df_hole['Среднее_Гаусса'] = df_hole['Гаусса'].rolling(window=3, min_periods=1).mean().shift(-1)
        df_hole['Средний_Нов.Бок.Лоб'] = df_hole['Нов.Бок.Лоб'].rolling(window=3, min_periods=1).mean().shift(-1)
        df_hole.loc[df_hole.index[-1], 'Среднее_Гаусса'] = (df_hole['Гаусса'].iloc[-1] + df_hole['Гаусса'].iloc[-2]) / 2
        df_hole.loc[df_hole.index[-1], 'Средний_Нов.Бок.Лоб'] = (df_hole['Нов.Бок.Лоб'].iloc[-1] +
                                                                 df_hole['Нов.Бок.Лоб'].iloc[-2]) / 2

        df_hole['Среднее_Гаусса_Тройка'] = df_hole['Среднее_Гаусса'].rolling(window=3, min_periods=1).mean().shift(-1)
        df_hole.loc[df_hole.index[-1], 'Среднее_Гаусса_Тройка'] = (
                                                                          df_hole['Среднее_Гаусса'].iloc[-1] +
                                                                          df_hole['Среднее_Гаусса'].iloc[-2]
                                                                  ) / 2

        df_hole['Средний_Нов.Бок.Лоб_Тройка'] = df_hole['Средний_Нов.Бок.Лоб'].rolling(window=3,
                                                                                       min_periods=1).mean().shift(-1)
        df_hole.loc[df_hole.index[-1], 'Средний_Нов.Бок.Лоб_Тройка'] = (
                                                                               df_hole['Средний_Нов.Бок.Лоб'].iloc[-1] +
                                                                               df_hole['Средний_Нов.Бок.Лоб'].iloc[-2]
                                                                       ) / 2

    def round_values(self, row, column_name, number, ugv_data):
        value = row[column_name]
        if row['Глубина'] > ugv_data[number] and 'песок' in row['Имя_слоя']:
            return round(value * 1.1, 2)
        return round(value, 2)

    def process_data(self, ugv_data, number: int) -> None:
        self.calculate_statistics(number)
        self.calculate_soil_load(number)
        df_new = self.df_list[number]
        df_new['Округленный_Средний_Нов.Бок.Лоб'] = df_new.apply(
            lambda row: self.round_values(row, 'Средний_Нов.Бок.Лоб_Тройка', number, ugv_data), axis=1
        )
        df_new['Округленное_Среднее_Гаусса'] = df_new.apply(
            lambda row: self.round_values(row, 'Среднее_Гаусса_Тройка', number, ugv_data), axis=1
        )

# if __name__ == "__main__":
#     soil_calculator = SoilLoadCalculator()
#     ugv_data = [0, 0.2, 0.1]
#     data = {
#         'песок гравелистый плотный': [24, 50, 3, 8.075],
#         'песок гравелистый ср': [16, 25, 3, 8.075],
#         'песок гравелистый рыхлый': [10, 16, 3, 8.075],
#         'песок крупный плотный': [16, 22, 3, 9.69],
#         'песок крупный ср': [7, 18, 3, 9.69],
#         'песок крупный рыхлый': [1, 5, 3, 9.69],
#         'песок средний плотный': [10, 20, 4, 11.305],
#         'песок средний ср': [6, 15, 5, 11.305],
#         'песок средний рыхлый': [2.5, 5, 6, 11.305],
#         'песок мелкий плотный': [5, 12, 7, 12.92],
#         'песок мелкий ср': [4, 10, 7.5, 12.92],
#         'песок мелкий рыхлый': [1.1, 4, 8, 12.92],
#         'песок пылеватый плотный': [10, 12, 15, 20.1875],
#         'песок пылеватый ср': [3, 10, 25.5, 24.225],
#         'песок пылеватый рыхлый': [1, 3, 30, 29.07],
#         'суглинок моренный твердый': [4, 6, 10, 40.375],
#         'суглинок моренный полутвердый': [2.5, 4.7, 15, 48.45],
#         'суглинок моренный тугопластичный': [2.25, 4.2, 20, 56.525],
#         'суглинок моренный мягкопластичный': [1.375, 3.8, 25, 64.6],
#         'суглинок моренный текучепластичный': [0.5, 0.8, 30, 72.675],
#         'суглинки тяжелые, моренные полутвердые и тугопластичные': [3, 4.9, 10, 62.7],
#         'суглинки тяжелые, моренные мягкопластичные': [1, 1.5, 15, 42.75],
#         'пылеватые, аллювиальные, слабозаторфованые': [0.8, 1.5, 22, 52],
#         'глины полутвердые и твердые': [2.5, 3.5, 50, 80],
#         'глины твердые': [4, 6.5, 20, 40],
#         'глины мягкопластичные': [1.5, 3.75, 30, 60],
#         'юрские суглинки пластичные': [2.5, 3.5, 30, 50],
#         'покровный суглинок мягкопластичный': [0.7, 2.9, 50, 90],
#         'суглинок флювиогляциальный тугопл': [1, 3.3, 20, 60],
#         'суглинок техногенн': [0.7, 5.5, 20, 56.525],
#         'супесь пастичная сильнопесчанистая': [0.7, 5.5, 10, 40]
#     }
#     samples = {
#         0: [('песок гравелистый рыхлый', 0.1), ('песок средний ср', 0.2), ('песок средний рыхлый', 0.3),
#             ('песок гравелистый плотный', 0.4)],
#         1: [('песок крупный ср', 0.1), ('песок гравелистый рыхлый', 0.3),
#             ('пылеватые, аллювиальные, слабозаторфованые', 0.4)],
#         2: [('песок мелкий ср', 0.1), ('песок гравелистый плотный', 0.3), ('песок гравелистый плотный', 0.5)]
#     }
#
#     # Create a list to store generated DataFrames
#     soil_calculator.df_list = []
#
#     # Generate and process DataFrames
#     for borehole_num in range(len(samples)):
#         soil_calculator.generate_data(borehole_num, data, samples)
#         soil_calculator.process_data(ugv_data, borehole_num)
#     # You can access the generated DataFrames using soil_calculator.df_list
#     for idx, df in enumerate(soil_calculator.df_list):
#         print(f"Borehole {idx} DataFrame:")
#         print(df)
