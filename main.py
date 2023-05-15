import pandas as pd
import xlwings as xw

reviews = pd.read_csv('data/reviews_sample.csv', index_col=0)
reviews = reviews.rename(columns={'Unnamed: 0' : 'id'})
recipes = pd.read_csv('data/recipes_sample.csv', index_col=0, usecols=['id', 'name', 'minutes', 'submitted', 'description', 'n_ingredients'])

# Случайным образом выбираем 5% строк из каждой таблицы
reviews_sample = reviews.sample(frac=0.05)
recipes_sample = recipes.sample(frac=0.05)

# Создаем объект ExcelWriter
writer = pd.ExcelWriter('recipes.xlsx', engine='xlsxwriter')

# Сохраняем таблицы в разные листы
recipes_sample.to_excel(writer, sheet_name='Рецепты', index=False)
reviews_sample.to_excel(writer, sheet_name='Отзывы', index=False)

# Получаем объект Workbook и лист Рецепты
wb = xw.Book('recipes.xlsx')
sht = wb.sheets['Рецепты']

# Добавляем столбец seconds_assign, показывающий время выполнения рецепта в секундах
seconds_assign = recipes_sample['minutes'] * 60
sht.range('G1').value = 'seconds_assign'
sht.range('G2').options(transpose=True).value = seconds_assign

# Добавляем столбец seconds_formula, показывающий время выполнения рецепта в секундах
sht.range('H1').value = 'seconds_formula'
sht.range('H2').formula = '=E2*60'

# Делаем названия всех добавленных столбцов полужирными и выравниваем по центру ячейки
sht.range('G1:H1').api.Font.Bold = True
sht.range('G1:H1').api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter

# Раскрашиваем ячейки столбца minutes
for cell in sht.range('D2:D{}'.format(len(recipes_sample)+1)):
    if cell.value < 5:
        cell.color = (0, 255, 0)  # зеленый
    elif cell.value < 10:
        cell.color = (255, 255, 0)  # желтый
    else:
        cell.color = (255, 0, 0)  # красный

# Добавляем столбец n_reviews, содержащий кол-во отзывов для этого рецепта
sht.range('I1').value = 'n_reviews'
sht.range('I2').formula = '=COUNTIF(Отзывы!A:A, A2)'

# Сохраняем и закрываем файл
wb.save()
wb.close()