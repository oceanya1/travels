from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
import datetime as dt
wb=Workbook()
H=PatternFill("solid",fgColor="1F4E79"); HF=Font(bold=True,color="FFFFFF")
INP=Font(color="0000FF"); EX=Font(italic=True,color="808080")
TOT=Font(bold=True)
def table(ws,headers,widths,rows,start=1):
    for i,(h,w) in enumerate(zip(headers,widths),1):
        c=ws.cell(start,i,h); c.fill=H; c.font=HF; c.alignment=Alignment(wrap_text=True,vertical="center")
        ws.column_dimensions[get_column_letter(i)].width=w
    for r,row in enumerate(rows,start+1):
        for i,v in enumerate(row,1):
            c=ws.cell(r,i,v); c.alignment=Alignment(wrap_text=True,vertical="top")
    ws.freeze_panes=ws.cell(start+1,1)
def dv(ws,opts,rng):
    d=DataValidation(type="list",formula1='"'+",".join(opts)+'"',allow_blank=True); ws.add_data_validation(d); d.add(rng)

# Обзор
ws=wb.active; ws.title="Обзор"
ws["A1"]="Поездка: <страна / город, месяц год>"; ws["A1"].font=Font(bold=True,size=16)
ws["A2"]="Синий текст — поля для заполнения. Серый курсив — пример, удалите или замените. Остальное считается само."; ws["A2"].font=EX
rows=[("Куда","<страна, города>"),("Кто едет","Аня, Саша"),("Дата вылета",None),("Дата возвращения",None),
("Дней в поездке","=IF(AND(B6<>\"\",B7<>\"\"),B7-B6+1,\"\")"),
("Местная валюта","<KRW / USD / …>"),("Курс: рублей за 1 ед. валюты",None),("Дата курса",None),
("Бюджет на двоих, ₽",None),("Запланировано, ₽","='Бюджет'!E2"),("Потрачено, ₽","='Бюджет'!G2"),("Остаток от бюджета, ₽","=IF(B12=\"\",\"\",B12-B14)"),
("Список вещей","Google Keep: <ссылка на заметку>"),("Интерактивная страница","<ссылка>"),("Заметки",None)]
for i,(k,v) in enumerate(rows,4):
    ws.cell(i,1,k).font=Font(bold=True); c=ws.cell(i,2,v)
    if not (isinstance(v,str) and v.startswith("=")): c.font=INP
ws.column_dimensions["A"].width=32; ws.column_dimensions["B"].width=50
for r in (6,7,11): ws.cell(r,2).number_format="DD.MM.YYYY"
for r in (12,13,14,15): ws.cell(r,2).number_format='#,##0 "₽"'
ws["B10"].number_format="0.0000"

# Маршрут
ws=wb.create_sheet("Маршрут")
hdr=["День","Дата","День недели","Город","План на день","Ночёвка","Переезды","Заметки"]
table(ws,hdr,[7,12,12,16,50,26,26,30],[])
for d in range(1,31):
    r=d+1; ws.cell(r,1,d)
    ws.cell(r,2,f"=IF('Обзор'!$B$6=\"\",\"\",IF(A{r}>'Обзор'!$B$8,\"\",'Обзор'!$B$6+A{r}-1))").number_format="DD.MM"
    ws.cell(r,3,f"=IF(B{r}=\"\",\"\",PROPER(TEXT(B{r},\"dddd\")))")
    for c in (4,5,6,7,8): ws.cell(r,c).font=INP
for c,v in zip((4,5,6,7,8),("Пример: Сеул","Прилёт, заселение, прогулка по району","Отель «…», 3 ночи","Аэропорт → отель: экспресс","")):
    ws.cell(2,c,v).font=EX

# Бюджет
ws=wb.create_sheet("Бюджет")
ws["A1"]="Итого"; ws["A1"].font=TOT
hdr=["Категория","Статья","План, местн. вал.","План, ₽ (если сразу в рублях)","План итого, ₽","Факт, ₽","Факт итого, ₽","Кто платил","Комментарий"]
table(ws,hdr,[16,30,14,16,14,12,14,12,30],[],start=3)
ws.freeze_panes="A4"
cats=["Билеты","Жильё","Транспорт на месте","Еда","Развлечения и музеи","Покупки","Связь","Страховка","Документы и визы","Другое"]
for i in range(4,64):
    ws.cell(i,5,f"=IF(AND(C{i}=\"\",D{i}=\"\"),\"\",IFERROR(C{i}*'Обзор'!$B$10,0)+N(D{i}))")
    ws.cell(i,7,f"=IF(F{i}=\"\",\"\",F{i})")
    for c in (1,2,3,4,6,8,9): ws.cell(i,c).font=INP
    for c in (4,5,6,7): ws.cell(i,c).number_format='#,##0'
ws["E2"]="=SUM(E4:E63)"; ws["G2"]="=SUM(G4:G63)"
ws["D2"]="План:"; ws["F2"]="Факт:"
for c in ("E2","G2"): ws[c].number_format='#,##0 "₽"'; ws[c].font=TOT
ex=["Билеты","Перелёт туда-обратно, 2 чел.",None,None,None,None,None,"Аня",""]
for i,v in enumerate(ex,1):
    if v: ws.cell(4,i,v).font=EX
dv(ws,cats,"A4:A63"); dv(ws,["Аня","Саша","Пополам"],"H4:H63")
# by category summary
ws["K3"]="Категория"; ws["L3"]="План, ₽"; ws["M3"]="Факт, ₽"
for c in ("K3","L3","M3"): ws[c].fill=H; ws[c].font=HF
for j,cat in enumerate(cats,4):
    ws.cell(j,11,cat); ws.cell(j,12,f"=SUMIF($A$4:$A$63,K{j},$E$4:$E$63)").number_format='#,##0'
    ws.cell(j,13,f"=SUMIF($A$4:$A$63,K{j},$G$4:$G$63)").number_format='#,##0'
ws.column_dimensions["K"].width=22; ws.column_dimensions["L"].width=12; ws.column_dimensions["M"].width=12

# Брони
ws=wb.create_sheet("Брони")
hdr=["Тип","Что","С","По","Статус","Стоимость, ₽","Номер брони / ссылка","Бесплатная отмена до","Распечатано?","Заметки"]
table(ws,hdr,[14,30,11,11,14,13,30,16,12,30],[])
for c,v in zip(range(1,11),["Перелёт","Пример: Москва → Сеул","","","Ищем","","","","",""]):
    if v: ws.cell(2,c,v).font=EX
for col in ("C","D","H"):
    for r in range(2,61): ws[f"{col}{r}"].number_format="DD.MM.YYYY"
dv(ws,["Перелёт","Жильё","Поезд/автобус","Трансфер","Экскурсия","Страховка","Другое"],"A2:A60")
dv(ws,["Ищем","Выбрали","Забронировано","Оплачено","Отменено"],"E2:E60")
dv(ws,["да","нет","не нужно"],"I2:I60")

# Места и еда
ws=wb.create_sheet("Места и еда")
hdr=["Город","Название","Тип","Район / адрес","Ссылка (карта, отзывы)","Приоритет","Были?","Впечатления"]
table(ws,hdr,[14,28,14,22,34,12,9,34],[])
for c,v in zip(range(1,9),["Пример: Сеул","","Достопримечательность","","","Обязательно","",""]):
    if v: ws.cell(2,c,v).font=EX
dv(ws,["Достопримечательность","Музей","Еда","Кафе","Бар","Рынок / шопинг","Природа","Другое"],"C2:C200")
dv(ws,["Обязательно","Хорошо бы","Если будет время"],"F2:F200"); dv(ws,["да","нет"],"G2:G200")

# Документы
ws=wb.create_sheet("Документы")
hdr=["Что","Аня","Саша","Срок / действует до","Комментарий"]
table(ws,hdr,[38,14,14,18,40],[])
items=[("Загранпаспорт (срок действия)",""),("Виза или электронное разрешение на въезд","Проверить требования для граждан РФ"),
("Медицинская страховка",""),("Военный билет","Только Саша"),("Копии паспортов (бумага + в телефоне)",""),
("Распечатки: билеты, отели, страховка",""),("Наличные в валюте",""),("Карта, которая работает в стране","Российские карты за рубежом обычно не работают"),
("Связь: eSIM / сим-карта",""),("Приложения: карты, переводчик, такси",""),("Водительские права (если нужны)","")]
for r,(a,b) in enumerate(items,2):
    ws.cell(r,1,a); ws.cell(r,5,b)
    for c in (2,3,4): ws.cell(r,c).font=INP
ws.cell(5,2,"не нужно")
dv(ws,["нужно сделать","в процессе","готово","не нужно"],"B2:C40")
for r in range(2,41): ws[f"D{r}"].number_format="DD.MM.YYYY"
wb.save("Шаблон поездки.xlsx")
