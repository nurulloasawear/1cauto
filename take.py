import pyautogui as boshqarish
import time as kutish
import pandas as pd
import openpyxl
from datetime import datetime

start_button_image_path = r"C:\code\1c\buttons\start\start.jpg"
ok_button_image_path = r"C:\code\1c\buttons\ok\ok.png"
podbor_button_image_path = r"C:\code\1c\buttons\podbor\podbor.jpg"
dot_button_image_path = r"C:\code\1c\buttons\dot\dot.jpg"
x_button_image_path = r"C:\code\1c\buttons\x\x.jpg"
first_button_image_path = r"C:\code\1c\buttons\first\first.jpg"
second_button_image_path = r"C:\code\1c\buttons\second\second.jpg"
oks_button_image_path = r"C:\code\1c\buttons\oks\oks.png"

ot_input_image_path = r"C:\code\1c\input\ot\ot.jpg"
comment_input_image_path = r"C:\code\1c\input\comment\comment.jpg"
sku_add_input_image_path = r"C:\code\1c\input\sku_add\sku_add.jpg"
price_input_image_path  = r"C:\code\1c\input\price\price.jpg"

excel_path = r"D:\test\1c экзель.XLSX"

class HelperFunctions:
    @staticmethod
    def format_date(value):
        if value is None or pd.isna(value):
            return ""
        if isinstance(value,(pd.Timestamp,datetime)):
            return value.strftime("%d.%m.%Y")
        try:
            dt = pd.to_datetime(value)
            return dt.strftime("%d.%m.%Y")
        except Exception:
            return str(value)

class ImageFinder:
    def __init__(self, confidence=0.8):
        self.confidence = confidence

    def find(self, image, region=None, confidence=None,timeout=5, interval=0.3):
        end = kutish.time() + timeout
        while kutish.time() < end:
            try:
                print(confidence)
                pos = boshqarish.locateCenterOnScreen(
                        image,
                        confidence= confidence if  confidence is not None else self.confidence,
                        region=region
                    )
                if pos:
                    return pos
            except Exception as e:
                print(e)
            kutish.sleep(interval)
        return None
class OneCButton:
    def __init__(self):
        self.start = start_button_image_path
        self.ok = ok_button_image_path
        self.podbor = podbor_button_image_path
        self.dot = dot_button_image_path
        self.x = x_button_image_path
        self.first = first_button_image_path
        self.second = second_button_image_path
        self.oks = oks_button_image_path
class OneCInput:
    def __init__(self):
        self.ot = ot_input_image_path
        self.comment = comment_input_image_path
        self.sku_add = sku_add_input_image_path
        self.prices = price_input_image_path

class EXEL:
    helper = HelperFunctions()
    def __init__(self, file):
        self.df = pd.read_excel(file,header=2)

        # ustun nomlarini tozalash (probel, katta-kichik harf)
        self.df.columns = [c.strip() for c in self.df.columns]

    def rows(self):
        for _, row in self.df.iterrows():

            # price ba’zan "ЛОЖЬ" bo‘lishi mumkin
            raw_price = row.get("цена продажи за одну штуку")

            price = None
            if isinstance(raw_price, (int, float)):
                price = raw_price

            yield {
                "date":HelperFunctions.format_date(row.get("дата")),
                "order_number": row.get("Номер заказа"),
                "sku": row.get("sku"),
                "quantity": row.get("количество"),
                "price": price,
                "total_sum": row.get("общая сумма заказа"),
                "status": row.get("Транзакции по заказам и товарам статус"),
            }

    def __str__(self):
        return f"EXEL rows: {len(self.df)}"


def main():
    print("3 - soniya kuting malumotlar qayta ishlashmoqda..")
    exel = EXEL(excel_path)
    tugmalar = OneCButton()
    rasm = ImageFinder()
    kiritishlar = OneCInput()
    kutish.sleep(3)
    if exel:
        print("Malumotlar Muvaqiyatli qayta ishlandi")
        first_pos = rasm.find(tugmalar.first)
        second_pos = rasm.find(tugmalar.second)
        while True:
            if first_pos:
                boshqarish.click(first_pos)
                print("first tugmasi bosildi")
                boshqarish.press('end')
                start_pos = rasm.find(tugmalar.start,(27,906,138,90),0.1)
                print(start_pos)
                if start_pos:
                    print("start pos topildi ...")
                    boshqarish.click(start_pos)
                    boshqarish.press('delete')
                    boshqarish.press("enter")
                    boshqarish.doubleClick(start_pos)
                    sana = rasm.find(kiritishlar.ot)
                    rows = exel.rows()
                    excel_dict = next(rows)
                    print("sana kirtish maydoni topildi..")
                    if sana:
                        x,y = sana
                        boshqarish.click(x+20,y)
                        boshqarish.hotkey('ctrl','a')
                        # print(type(excel_vaqt))
                        boshqarish.press('backspace')
                        print(str(excel_dict['date']))
                        boshqarish.write(str(excel_dict['date'])+ "18:00:00")
                        print("hayr")
                        # kutish.sleep(4)
                        x_ochirish = rasm.find(tugmalar.x)
                        if x_ochirish:
                            print("x - ochirsh topildi")
                            boshqarish.click(x_ochirish)
                            comment = rasm.find(kiritishlar.comment)
                            x,y = comment
                            boshqarish.doubleClick(x+100,y)
                            boshqarish.press('backspace')
                            print(str(excel_dict['order_number']))
                            boshqarish.write(str(round(excel_dict['order_number'])))
                            print("hayr")
                            podbor = rasm.find(tugmalar.podbor)
                            if podbor:
                                print("podbor topildi-...")
                                boshqarish.click(podbor)
                                sku_add = rasm.find(kiritishlar.sku_add)
                                if sku_add:
                                    print("sku add topildi-...")
                                    boshqarish.click(sku_add)
                                    boshqarish.write(str(excel_dict['sku']))
                                    boshqarish.press('enter')
                                    boshqarish.press('enter')
                                    price = rasm.find(kiritishlar.prices)
                                    print(price)
                                    if price:
                                        print("price topildi -...")
                                        xx,yy = price
                                        kutish.sleep(2)
                                        boshqarish.doubleClick(xx,yy+23)
                                        boshqarish.write(str(excel_dict['price']))
                                        print(str(excel_dict['price']))
                                        ok = rasm.find(tugmalar.oks)
                                        if ok:
                                            print("ok topildi-...")
                                            boshqarish.click(ok)
                                            boshqarish.press('enter')
                if second_pos:
                    boshqarish.click(first_pos)
                    print("second tugmasi bosildi")
                    boshqarish.press('end')
                    start_pos = rasm.find(tugmalar.start,(27,906,138,90),0.1)
                    print(start_pos)
                    if start_pos:
                        print("start pos topildi ...")
                        boshqarish.click(start_pos)
                        boshqarish.press('delete')
                        boshqarish.press("enter")
                        boshqarish.doubleClick(start_pos)
                        sana = rasm.find(kiritishlar.ot)
                        rows = exel.rows()
                        excel_dict = next(rows)
                        print("sana kirtish maydoni topildi..")
                        if sana:
                            x,y = sana
                            boshqarish.click(x+20,y)
                            boshqarish.hotkey('ctrl','a')
                            # print(type(excel_vaqt))
                            boshqarish.press('backspace')
                            print(str(excel_dict['date']))
                            boshqarish.write(str(excel_dict['date'])+ "18:00:00")
                            print("hayr")
                            # kutish.sleep(4)
                            x_ochirish = rasm.find(tugmalar.x)
                            if x_ochirish:
                                print("x - ochirsh topildi")
                                boshqarish.click(x_ochirish)
                                comment = rasm.find(kiritishlar.comment)
                                x,y = comment
                                boshqarish.doubleClick(x+100,y)
                                boshqarish.press('backspace')
                                print(str(excel_dict['order_number']))
                                boshqarish.write(str(round(excel_dict['order_number'])))
                                print("hayr")
                                podbor = rasm.find(tugmalar.podbor)
                                if podbor:
                                    print("podbor topildi-...")
                                    boshqarish.click(podbor)
                                    sku_add = rasm.find(kiritishlar.sku_add)
                                    if sku_add:
                                        print("sku add topildi-...")
                                        boshqarish.click(sku_add)
                                        boshqarish.write(str(excel_dict['sku']))
                                        boshqarish.press('enter')
                                        boshqarish.press('enter')
                                        ok = rasm.find(tugmalar.oks)
                                        if ok:
                                            print("ok topildi-...")
                                            boshqarish.click(ok)
                                            boshqarish.press('enter')
                    




                        
        

