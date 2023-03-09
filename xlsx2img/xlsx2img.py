from PIL import ImageGrab
import xlwings as xw
import time

def xlsx_to_img(excel_path: str, sheet_name: str, img_path: str) -> None:
    """
    Parameters
    ----------
    excel_path: the xlsx file path
    sheet_name: the sheet of xlsx file
    img_path: the output png file path
    """
    app = xw.App(visible=False, add_book=False)

    wb = app.books.open(excel_path)
    sheet = wb.sheets(sheet_name)

    all = sheet.used_range
    all.api.CopyPicture()

    sheet.api.Paste()

    pic = sheet.pictures[-1]
    pic.api.Copy()

    time.sleep(3)

    img = ImageGrab.grabclipboard()
    img.save(img_path)

    pic.delete()
    wb.save(excel_path)
    wb.close()
    app.kill()
