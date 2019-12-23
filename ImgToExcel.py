import os
import tkinter as tk
from tkinter import filedialog
import pandas as pd
import xlsxwriter
from PIL import Image
from numpy import shape, array, hstack


def imgpath():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename()
    return Image.open(f'{file_path}')


def resize(x, img):
    ratio = shape(img)[0] / shape(img)[1]
    y = int(x * ratio)
    return [img.resize((x, y), resample=Image.BILINEAR), y]


def rgb_data(pix, x, y):
    pix = array(pix)
    r, g, b = pix[:, :, 0], pix[:, :, 1], pix[:, :, 2]
    data = hstack((r.flatten()[:, None], g.flatten()[:, None], b.flatten()[:, None]))
    df = pd.DataFrame(data, columns=['r', 'g', 'b'])
    color = []
    for i in range(x * y):
        arr = df.iloc[i].values
        color.append(arr)
    df['color'] = color

    def my_test(a, b, c): return '#{:02x}{:02x}{:02x}'.format(a, b, c)

    df['hex'] = df['color'].apply(lambda row: my_test(row[0], row[1], row[2]))
    im = df['hex'].values
    return im.reshape(y, x)


def create_excel(im, path):
    workbook = xlsxwriter.Workbook(os.path.join(path, "imgtoxlsx.xlsx"))
    worksheet = workbook.add_worksheet()
    for row in range(im.shape[0]):
        for col in range(im.shape[1]):
            colore = im[row, col]
            worksheet.set_column(row, col, 1.5)
            cell_format = workbook.add_format()
            cell_format.set_bg_color(colore)
            worksheet.write(row, col, '', cell_format)
    workbook.close()


def main():
    img = imgpath()
    x = int(input("Value in pixel of the X-axis:"))
    img_resized, y = resize(x, img)
    color_value = rgb_data(img_resized, x, y)
    root = tk.Tk()
    root.withdraw()
    path = filedialog.askdirectory(initialdir='.')
    create_excel(color_value, path)


if __name__ == "__main__":
    main()
