import pandas as pd
from pathlib import Path
import xlwings as xw

##### THIS IS ONLY FOR SAME SHAPED DATA FRAMES I.E. SAME NUMBER OF ROWS AND COLUMS #####

# original = Path.cwd() / "original.xlsx"
# to_compare = Path.cwd() / "to_compare.xlsx"

df_original = pd.read_csv("original.csv")
df_to_compare = pd.read_csv("to_compare.csv")

if df_to_compare.shape == df_original.shape:
    print('Same shaped')
else:
    print('Not same shaped')

# difference_columns = df_to_compare.compare(df_original, align_axis=1)
# difference_rows = df_to_compare.compare(df_original, align_axis=0)
# print(difference)


##### THIS IS FOR DIFFERENTLY SHAPED DATA FRAMES I.E. NOT THE SAME NUMBER OF ROWS AND COLUMS #####

df_original_updated = df_original.reset_index()
print(df_original_updated.head())

difference = pd.merge(df_to_compare, df_original_updated, how='outer', indicator=True)
print(difference)

highlight = difference.query(" _merge != 'both'")
highlight_rows = highlight['index'].tolist()

print(highlight_rows)
row_offset_in_excel = 2

highlight_rows = [row + row_offset_in_excel for row in highlight_rows]

print(highlight_rows)

with xw.App(visible=False) as app:
    updated_wb = app.books.open('original.csv')
    updated_ws = updated_wb.sheets(1)
    rng = updated_ws.used_range

    print(f"Used range: {rng.address}")

    # Highlight the rows
    for row in rng.rows:
        if row.row in highlight_rows:
            row.color = (255, 71, 76) # red

    updated_wb.save('difference_highlighted.xlsx')
