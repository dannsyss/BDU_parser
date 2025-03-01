import pandas as pd
from deep_translator import GoogleTranslator


# Функция для перевода текста
def translate_text(text, src_lang='en', dest_lang='ru'):
    try:
        # Используем синхронный метод translate
        translated = GoogleTranslator(source=src_lang, target=dest_lang).translate(text)
        return translated
    except Exception as e:
        print(f"Ошибка при переводе текста: {e}")
        return text  # Возвращаем исходный текст в случае ошибки


# Основной скрипт
def main():
    # Чтение данных из исходного файла
    input_file = 'capec_results.xlsx'
    output_file = 'capec_results_translated.xlsx'

    df = pd.read_excel(input_file)

    # Перевод столбца "Наименование атаки"
    df['Наименование атаки'] = df['Наименование атаки'].apply(translate_text)

    # Сохранение результатов в новый файл
    df.to_excel(output_file, index=False)
    print(f"Переведенные данные сохранены в файл '{output_file}'")


if __name__ == "__main__":
    main()