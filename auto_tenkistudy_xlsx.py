import datetime
import urllib.request
from PIL import Image, ImageDraw, ImageFont
import io
import re
import sys
import csv

import os
import openpyxl
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as xlImage
#★ImageはPILと被り

#Img_r = 240    
Font_r = 25     #1999年以前のアジア天気図を使用した際に日付の記入する小ささ
Quality = 20    #出力画像のJPEG品質(0-100)
SheetSize = 1   #標準サイズに対する倍率
Img_rr = 2      #セルのサイズ(100%)に対して何倍のサイズにするのか
Img_r = SheetSize * 160 * Img_rr #リサイズ後の縦横サイズ


def cropa_to_pil(crop_a):
    if len(crop_a) != 4:
        print("invalid crop_a.")
        return None
    return (crop_a[0], crop_a[1], crop_a[0]+crop_a[2], crop_a[1]+crop_a[3])
    
def pil_to_blob(img, q, f="JPEG"):
    data = io.BytesIO()
    img.save(data, f, quality=q)
    data.seek(0)
    return data.read()
    
def get_comment_100(date):
    msg = ""
    if isinstance(date, datetime.date) and date.year >= 1996 and (date.year < 2024 or (date.year == 2024 and date.mon <= 6)):
        url = f"https://agora.ex.nii.ac.jp/cgi-bin/weather-chart/search_day.pl?year={date.year}&month={date.month}&day={date.day}&lang=ja"
        #print(url)
        req = urllib.request.Request(url)
        with urllib.request.urlopen(req) as res:
            rawdata = res.read().decode('UTF-8')
        #print(rawdata)
        results = re.findall(r"<(?:dt|dd)>(.*?)</(?:dt|dd)>", rawdata, re.DOTALL)
        for result in results:
            msg += result.strip() + "\n" 
        
    return msg.strip()
    
def get_crop_a_100(mon_i, day_i, F99):
    if mon_i >= 202407:
        crop_a = [7,4,581,569]
    elif mon_i == 201708 or (201801 >= mon_i and mon_i >= 201711) or (mon_i == 201810 and day_i % 100 <= 15):
        crop_a = [6,0,440,455]
    elif mon_i >= 200502 or (200302 >= mon_i and mon_i >= 200001):
        crop_a = []
    elif mon_i >= 200303:
        if day_i % 100 == 1:
            crop_a = [0,17,447,462]
        else:
            crop_a = [104,17,435,449]
    elif F99 and mon_i >= 199906:
        crop_a = [19,16,450,452]
    elif F99 and mon_i >= 199902:
        crop_a = [112,15,417,427]
    elif mon_i >= 199701:
        crop_a = [354,111,615,625]
    elif mon_i >= 199603:
        crop_a = [373,161,479,487]
    elif mon_i >= 199103 or (mon_i >= 198903 and day_i%2==1):
        crop_a = [393,133,567,564]
    elif mon_i >= 198903 and day_i%2 == 0:
        crop_a = [340,136,567,571]
    elif mon_i >= 195808:
        crop_a = [381,131,571,581]
    else:
        crop_a = None
    
    return crop_a

def get_img_100(Date, Dest, q, F99=True):
    #見つからない場合はカラー天気図を使用
    day_i = Date.year*10000 + Date.month*100 + Date.day
    mon_i = Date.year*100 + Date.month
    loop = True
    no_utc0 = False
    n = 0
    
    while loop:
        if mon_i == 202301 or mon_i >= 202407:
            loop = False
            #https://agora.ex.nii.ac.jp/digital-typhoon/weather-chart/wxchart/202407/SPAS_COLOR_202407010000.png
            url = f"https://agora.ex.nii.ac.jp/digital-typhoon/weather-chart/wxchart/{mon_i}/SPAS_COLOR_{day_i}0000.png"
        elif mon_i <= 198902:
            n += 1
            url = f"https://agora.ex.nii.ac.jp/digital-typhoon/weather-chart/thumb/as/1280x960/{mon_i}/{day_i}_{n}.jpg"
        elif (mon_i <= 199912 and not F99) or mon_i <=199901:
            loop = False
            url = f"https://agora.ex.nii.ac.jp/digital-typhoon/weather-chart/thumb/as/1280x960/{mon_i}/{day_i}00.jpg"
        else:
            loop = False
            url = f"https://agora.ex.nii.ac.jp/digital-typhoon/weather-chart/thumb/js/640x480/{mon_i}/{day_i}00.jpg"
        print(url)
        
        req = urllib.request.Request(url)
        with urllib.request.urlopen(req) as res:
            raw_img = res.read()
        
        raw_size = len(raw_img)
        print(f"file acquired. ({int(raw_size/1024)}kB)")
        if raw_size < 2000:
            if loop == 0:
                #普通に見つからない
                print("no img found")
                return None
            else:
                no_utc0 = True
                if n >= 10:
                    #1989以前で見つからない
                    print("no img found")
                    return None
        else:
            loop = False
    
    img = Image.open(io.BytesIO(raw_img))
    
    crop_a = get_crop_a_100(mon_i, day_i, F99)
        
    #print(cropa_to_pil(crop_a))
    if crop_a != None:
        img = img.crop(cropa_to_pil(crop_a))
    img = img.resize((Img_r,Img_r))
    
    #文字入れ
    if (mon_i <= 199912 and not F99) or mon_i <= 199901:
        draw = ImageDraw.Draw(img)
        font = ImageFont.truetype('./ipag.ttf', int(Img_r/Font_r))
        title_text = f"{int(day_i/10000)}.{int(day_i/100)%100}.{day_i%100} 00UTC"
        if no_utc0:
            title_text += "(?)"
        draw.rectangle(draw.textbbox((Img_r,Img_r), title_text, font, anchor='rd'), fill='#FFFFFF')
        draw.text((Img_r,Img_r), title_text, '#000000', font, anchor='rd')
    
    img = img.convert("RGB")
    if Dest != "blob:":
        img.save(Dest, quality=q)
        return None
    else:
        return pil_to_blob(img, q)
        
def tsf_to_xlsx(tsf, dest, q=20):
    workbook = Workbook()
    sheet = workbook.active
    sheet.column_dimensions["A"].width = 7    * SheetSize #展開名
    sheet.column_dimensions["B"].width = 13   * SheetSize #考察
    sheet.column_dimensions["C"].width = 10   * SheetSize #山行名
    sheet.column_dimensions["D"].width = 13   * SheetSize #日付 
    sheet.column_dimensions["E"].width = 20   * SheetSize #画像スペース
    sheet.column_dimensions["F"].width = 25   * SheetSize #「今日の天気図」コメント
    sheet.column_dimensions["G"].width = 100  * SheetSize #本文
    
    for row in range(len(tsf)):
        date_a = tsf[row][1].split('/')
        date = datetime.date(int(date_a[0]),int(date_a[1]),int(date_a[2]))
        
        #【C】山行名
        sheet[f"C{row+1}"].value = tsf[row][0]
        sheet[f"C{row+1}"].alignment = openpyxl.styles.Alignment(wrapText=True)
        #【D】日付
        sheet[f"D{row+1}"].value = tsf[row][1]
        sheet[f"D{row+1}"].alignment = openpyxl.styles.Alignment(wrapText=True)
        #【E】画像
        img = xlImage(io.BytesIO(get_img_100(date, "blob:", q)))
        sheet.add_image(img, f"E{row +1}")
        sheet.row_dimensions[row+1].height = Img_r * 3 / (4 * Img_rr)
        #★初期化
        img = None
        #【F】コメント
        sheet[f"F{row+1}"].value = get_comment_100(date)
        sheet[f"F{row+1}"].alignment = openpyxl.styles.Alignment(wrapText=True)
        #【G】本文
        sheet[f"G{row+1}"].value = tsf[row][2]
        sheet[f"G{row+1}"].alignment = openpyxl.styles.Alignment(wrapText=True)
    
    if dest != "blob:":
        workbook.save(dest)
    else:
        iodata = io.BytesIO()
        workbook.save(iodata)
        iodata.seek(0)
        return iodata.read()
        
if __name__ == '__main__':
    args = sys.argv
    if len(args) < 3:
        print("Usage: ./auto_tenkistudy_xlsx.py [IN_TSF] [OUT_XLSX] (IMG_QUALITY = 20)")
        exit()
    #print(sys.argv)
    #data = [["山行名","2003/6/26","あああああああ\nいいいいいい"]]
    data = []
    with open(args[1], "r", encoding='cp932') as f:
        for row in csv.reader(f):
            data.append(row)
            #print(f"{type(row)}")
    
    if len(args) >= 4:
        q = int(args[3])
    else:
        q = 20
        
    print(data)
    tsf_to_xlsx(data, args[2], q)
    #print(tsf_to_xlsx(data, "blob:"))
    #date = datetime.date(2003,6,26)
    #print(get_comment_100(date))
    #print(get_img_100(date, "img.jpg", Quality))
    #print(get_img_100(date, "blob:", Quality))
    
