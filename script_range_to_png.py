import uno
from com.sun.star.beans import PropertyValue
import sys
import os

def log(message):
    print(f"[DEBUG] {message}")

def export_range_to_image(input_path, output_path, sheet_index, cell_range):
    try:
        pixel_width = 1
        pixel_height = 1
        # Подключение к LibreOffice
        local_context = uno.getComponentContext()
        resolver = local_context.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local_context)
        ctx = resolver.resolve("uno:socket,host=localhost,port=2099;urp;StarOffice.ComponentContext")

        # Открытие документа
        desktop = ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", ctx)
        input_url = uno.systemPathToFileUrl(os.path.abspath(input_path))
        doc = desktop.loadComponentFromURL(input_url, "_blank", 0, ())

        # Получение листа и диапазона
        sheet = doc.Sheets.getByIndex(sheet_index)
        range_obj = sheet.getCellRangeByName(cell_range)
        range_size = range_obj.Size

        # Установка области печати
        sheet.setPrintAreas((range_obj.RangeAddress,))

        #page_width = int((range_size.Width / 96)* 2540)
        #page_height = int((range_size.Height / 96)* 2540)
        # Установка масштабирования страницы
        page_styles = doc.StyleFamilies.getByName("PageStyles")
        page_style = page_styles.getByName(sheet.PageStyle)

        page_style.Width = range_size.Width
        page_style.Height = range_size.Height
        page_style.HeaderIsOn = False  # Масштабировать по высоте
        page_style.FooterIsOn = False  # Масштабировать по высоте

        page_style.ScaleToPagesX = 1  # Масштабировать по ширине
        page_style.ScaleToPagesY = 1  # Масштабировать по высоте

        # Вычисление коэффициентов масштабирования
        page_style.TopMargin = 0
        page_style.BottomMargin = 0
        page_style.LeftMargin = 0
        page_style.RightMargin = 0
        


        # Параметры экспорта
        filter_props = (
            PropertyValue(Name="FilterName", Value="calc_png_Export"),
            PropertyValue(Name="Selection", Value=range_obj),
            PropertyValue(Name="IsExportSelection", Value=True),
            PropertyValue(Name="Resolution", Value="dpi"),

        )

        # Экспорт в изображение
        output_url = uno.systemPathToFileUrl(os.path.abspath(output_path))
        doc.storeToURL(output_url, filter_props)
        
        doc.close(True)
        return True

    except Exception as e:
        log(f"Ошибка: {str(e)}")
        return False

if __name__ == "__main__":
    if len(sys.argv) != 5:
        print("Использование: python export_range_to_image.py input.xlsx output.jpg sheet_num range pixel_width pixel_height")
        sys.exit(1)
    
    if export_range_to_image(sys.argv[1], sys.argv[2], int(sys.argv[3]), sys.argv[4]):
        print("Изображение успешно создано!")
    else:
        print("Ошибка экспорта")
        sys.exit(1)
