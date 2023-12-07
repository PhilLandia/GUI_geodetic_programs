import numpy as np

def generate_random_array_with_ranges(n, ranges):
    # Генерация массива случайных значений в заданных диапазонах
    random_array = np.array([np.random.uniform(low, high) for low, high in ranges])

    # Нормировка массива, чтобы сумма была равна n
    random_array = random_array / np.sum(random_array) * n

    return random_array

# Пример использования
n = 100  # Сумма массива
ranges = [(0, 0), (50, 100), (20, 80), (6, 80)]  # Границы для каждого значения

result_array = generate_random_array_with_ranges(n, ranges)
print(result_array)
print("Сумма массива:", np.sum(result_array))