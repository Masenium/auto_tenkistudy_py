import datetime
import urllib.request
from PIL import Image, ImageDraw, ImageFont
import io
import re
import sys
import csv
import json

import os
import openpyxl
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as xlImage
#★ImageはPILと被り

from pdf2image import convert_from_path, convert_from_bytes

#Img_r = 240    
Font_r = 25     #1999年以前のアジア天気図を使用した際に日付の記入する小ささ
Quality = 20    #出力画像のJPEG品質(0-100)
SheetSize = 1.5 #標準サイズに対するセル寸法の倍率
Img_rr = 3      #セルのサイズ(100%)に対して何倍のサイズにするのか
Img_r = int(SheetSize * 160 * Img_rr) #リサイズ後の縦横サイズ
Log = 1         #0:エラーのみ 1:Infoも 2:全て


def cropa_to_pil(crop_a):
    if len(crop_a) != 4:
        print("Err : invalid crop_a.")
        return None
    return (crop_a[0], crop_a[1], crop_a[0]+crop_a[2], crop_a[1]+crop_a[3])
    
def pil_to_blob(img, q, f="JPEG"):
    data = io.BytesIO()
    img.save(data, f, quality=q)
    data.seek(0)
    return data.read()
    
def get_comment_100(date):
    msg = ""
    if isinstance(date, datetime.date):
        if date.year >= 1996 and (date.year < 2024 or (date.year == 2024 and date.mon <= 6)):
            url = f"https://agora.ex.nii.ac.jp/cgi-bin/weather-chart/search_day.pl?year={date.year}&month={date.month}&day={date.day}&lang=ja"
            if Log >= 2:
                print(url) 
            req = urllib.request.Request(url)
            with urllib.request.urlopen(req) as res:
                rawdata = res.read().decode('UTF-8')
            #print(rawdata)
            results = re.findall(r"<(?:dt|dd)>(.*?)</(?:dt|dd)>", rawdata, re.DOTALL)
            for result in results:
                msg += result.strip() + "\n"
        
    return msg.strip()
    
def get_comment_fc2(date):
    msg = ""
    url = f"https://spms.web.fc2.com/comment/{date.year*100 + date.month}_comment.txt"
    if True: #Log >= 2
        print(url)
    req = urllib.request.Request(url)
    try:
        with urllib.request.urlopen(req) as res:
            rawdata = res.read().decode('UTF-8').splitlines()
        msg = f"{rawdata[(date.day-1)*2]}\n{rawdata[(date.day-1)*2+1]}"
    except urllib.error.HTTPError:
        print("Err : No comment found.")
        
    return msg
    
def get_crop_a_100(mon_i, day_i, F99):
    if mon_i >= 202407:
        crop_a = [7,4,581,569]
    elif mon_i == 201708 or (201801 >= mon_i and mon_i >= 201711) or (mon_i == 201810 and day_i % 100 <= 15):
        crop_a = [6,0,440,455]
    elif mon_i >= 200502 or (200302 >= mon_i and mon_i >= 200001):
        crop_a = None
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
        crop_a = "Err"
    
    return crop_a

def get_img_100(Date, Dest, q=20, F99=True):
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
        
        if Log >= 2:
            print(url)
        
        req = urllib.request.Request(url)
        with urllib.request.urlopen(req) as res:
            raw_img = res.read()
        
        raw_size = len(raw_img)
        if Log >= 1:
            print(f"Info: file acquired. ({int(raw_size/1024)}kB)")
        if raw_size < 2000:
            if loop == 0:
                #普通に見つからない
                print("Err : no img found(A)")
                return None
            else:
                no_utc0 = True
                if n >= 10:
                    #1989以前で見つからない
                    print("Err : no img found(B)")
                    return None
        else:
            loop = False
    
    img = Image.open(io.BytesIO(raw_img))
    
    crop_a = get_crop_a_100(mon_i, day_i, F99)
    #print(crop_a)
        
    #print(cropa_to_pil(crop_a))
    if crop_a is not None:
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
        
def get_token_lib(pid, cid=None):
    url = "https://dl.ndl.go.jp/api/restriction/issue/token/info:ndljp/pid/" + pid
    req = urllib.request.Request(url)
    with urllib.request.urlopen(req) as res:
        data = json.load(res)
    
    if cid is None:
        return data
    else:
        return str(data["timestamp"]), data["tokens"][cid]
        
def get_iCB_lib(pid):
    url = "https://dl.ndl.go.jp/api/item/search/info:ndljp/pid/" + pid
    req = urllib.request.Request(url)
    with urllib.request.urlopen(req) as res:
        data = json.load(res)
    return data["item"]["initialContentBundle"]
    
def get_pid_from_M_lib(date):
    y = date.year
    m = date.month
    
    if y*100+m >= 202210:
        j_data = json.load(open("./202511_recwm.json", 'r'))
        #print(j_data)
        try:
            return j_data[str(y * 100 + m)]
        except KeyError:
            return None
    elif y*100+m >= 199902:
        return str(12897877 + (m - 2) * 7 + 84 * (y - 1999))
    elif y*100+m == 199901:
        return "12897869"
    elif y*100+m >= 199603:
        return str(12897665 + (m - 3) * 6 + (y - 1996) * 72)
    elif y*100+m >= 192907:
        return str(12896865 + (m - 7) + (y - 1929) * 12)
    elif y*100+m == 192906:
        return "13123100"
    elif y*100+m >= 188303:
        return str(12896310 + (m - 3) + (y - 1883) * 12)
    return None
    
def get_sK_lib(date, h=0):
    #2003/01/01~           Js_YYYYMMDDHH.pdf
    #2000/01/01~2002/12/31 JS_YYYYMMDDHH.pdf
    #1999/02/01~1999/12/31 Jp_DDHH.pdf
    #1996/03/01~1999/01/31 As_DDHH.pdf
    #1990/01/01~1996/02/29 YYYYMMDD_N.pdf
    #△1883/03/01~1989/12/31 YYYYMMDD.pdf
    y = date.year
    m = date.month
    d = date.day
    h = int(h)
    
    if h == 0:
        n = 0
    elif h == 12:
        n = 1
    else:
        n = None
    
    if y*100+m >= 200301:
         return f"Js_{y*1000000+m*10000+d*100+h}.pdf"
    elif y*100+m >= 200001:
        return f"JS_{y*1000000+m*10000+d*100+h}.pdf"
    elif y*100+m >= 199902:
        return f"Jp_{str(d*100+h).zfill(4)}.pdf"
    elif y*100+m >= 199603:
        return f"As_{str(d*100+h).zfill(4)}.pdf"
    elif y*100+m >= 199001:
        return f"{y*10000+m*100+d}_{n}.pdf"
    elif y*100+m >= 188303:
        return f"{y*10000+m*100+d}.pdf"
    return None
    
def get_iCB_cid_fname_lib(pid, sortKey):
    url = "https://dl.ndl.go.jp/api/item/search/info:ndljp/pid/" + pid
    req = urllib.request.Request(url)
    with urllib.request.urlopen(req) as res:
        data = json.load(res)
    #print(data,"\n")
    
    try:
        iCB = data["item"]["initialContentBundle"]
    except KeyError:
        return "",nil,"",""
    
    #print iCB,"\n"
    #print(data["item"]["contentsBundles"][0]["contents"])
    for a in data["item"]["contentsBundles"][0]["contents"]:
        if sortKey.upper() == a["sortKey"].upper():
            return iCB, a["id"], a["fileName"], a["path"]
    
    return "",None,"",""

def make_token_url(pid, cid, raw_token, timestamp):
    token = dict()
    token["pid"] = "info:ndljp/pid/" + pid
    token["cid"] = cid
    token["token"] = raw_token
    token["timestamp"] = timestamp
    return urllib.parse.quote(json.dumps(token))
    
def get_url_lib(date, h=0):
    pid = get_pid_from_M_lib(date)
    if pid is None:
        return None
    sK = get_sK_lib(date, h)
    #print(pid,",",sK,"\n")
    iCB, cid, fn, path = get_iCB_cid_fname_lib(pid, sK)
    ts, raw_token = get_token_lib(pid, cid)
    token = make_token_url(pid, cid, raw_token, ts)
    
    return ''.join(["https://dl.ndl.go.jp/contents/", pid, "/", iCB, "/", cid, "/", fn, "?token=", token])

def get_pdf_lib(date, h=0):
    url = get_url_lib(date, h)
    #print(url)
    if url is None:
        return None
    page = 0
    
    req = urllib.request.Request(url)
    res = urllib.request.urlopen(req)
    if int(int(res.code) / 100) == 2:
        return res.read()
    else:
        return None
        
def get_pdf_jma(date, h=0):
    url = f"https://www.data.jma.go.jp/yoho/data/wxchart/archive/{date.year}_{str(date.month).zfill(2)}/PDFDATA/JSMAP/Js_{date.year}{str(date.month).zfill(2)}{str(date.day).zfill(2)}00.pdf"
    req = urllib.request.Request(url)
    try:
        res = urllib.request.urlopen(req)
    except urllib.error.HTTPError:
        return None
    if int(int(res.code) / 100) == 2:
        return res.read()
    else:
        return None

def get_1999_crop(img):
    col = img.size[0]
    row = img.size[1]
    
    brk = False
    for a in range(sum(img.size)):
        if brk:
            break
        for b in range(a):
            #print(f"({b},{a-b}){img.getpixel((b, a-b))}")
            if img.getpixel((b, a-b))[0] != 255:
                brk = True
                nw = [b, a-b]
                break
    brk = False
    for a in range(sum(img.size)):
        if brk:
            break
        for b in range(a):
            #print(f"({col-b-1},{row-a+b-1}){img.getpixel((col-b-1, row-a+b-1))}")
            
            if img.getpixel((col-b-1, row-a+b-1))[0] != 255:
                brk = True
                se = [col-b, row-a+b]
                break
    return [nw[0],nw[1],se[0]-nw[0],se[1]-nw[1]]

def get_crop_a_jma(date, img):
    crop_a = None
    if date >= datetime.date(2006,3,1):
        crop_a = [156, 28,504,520]
    elif date >= datetime.date(2000,1,1):
        crop_a = [128, 22,528,542]
    elif date >= datetime.date(1999,6,1):
        #img.resize(None, 1000)
        crop_a = get_1999_crop(img)
    elif date >= datetime.date(1999,2,1):
        crop_a = [805,111,2953,3029]
    elif date >= datetime.date(1997,1,1):
        crop_a = [1769,561,3083,3128]
        
    return crop_a

def get_img_jma(date, dest, h=0, q=20, F99=True):
    page = 0
    raw_pdf = get_pdf_jma(date, h)
    if raw_pdf is None:
        raw_pdf = get_pdf_lib(date, h)
        if raw_pdf is None:
            print("Err : No img found.")
            return None
            
    img = convert_from_bytes(raw_pdf, fmt='ppm', dpi=70)[page]
    crop_a = get_crop_a_jma(date, img)
    #print(crop_a)
    if crop_a is not None:
        img = img.crop(cropa_to_pil(crop_a))
        img = img.resize((Img_r,Img_r))
    img = img.convert("RGB")
    if dest != "blob:":
        img.save(dest, quality=q)
        return None
    else:
        return pil_to_blob(img, q)
        
############################
###オート系関数
############################

def get_comment(date):
    if isinstance(date, datetime.date):
        if date.year >= 1996 and date < datetime.date(2024,7,1):
            return get_comment_100(date)
        else:
            return get_comment_fc2(date)
            
def get_img(date, dest, q=20, F99=True):
    if isinstance(date, datetime.date):
        if date < datetime.date(1999,2,1):
            return get_img_100(date, dest, q, F99)
        else:
            return get_img_jma(date, dest, 0, q, F99)
        
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
        
        #【A】展開
        sheet[f"A{row+1}"].alignment = openpyxl.styles.Alignment(horizontal="left", vertical="center", wrapText=True)
        #【B】考察
        sheet[f"B{row+1}"].alignment = openpyxl.styles.Alignment(horizontal="left", vertical="center", wrapText=True)
        #【C】山行名
        sheet[f"C{row+1}"].value = tsf[row][0]
        sheet[f"C{row+1}"].alignment = openpyxl.styles.Alignment(horizontal="left", vertical="center", shrink_to_fit=True)
        #【D】日付
        sheet[f"D{row+1}"].value = tsf[row][1]
        sheet[f"D{row+1}"].alignment = openpyxl.styles.Alignment(horizontal="left", vertical="center", shrink_to_fit=True)
        #【E】画像
        img = xlImage(io.BytesIO(get_img(date, "blob:", q)))
        sheet.add_image(img, f"E{row +1}")
        sheet.row_dimensions[row+1].height = Img_r * 3 / (4 * Img_rr)
        #★初期化
        img = None
        #【F】コメント
        sheet[f"F{row+1}"].value = get_comment(date)
        sheet[f"F{row+1}"].alignment = openpyxl.styles.Alignment(horizontal="left", vertical="top", wrapText=True)
        #【G】本文
        sheet[f"G{row+1}"].value = tsf[row][2]
        sheet[f"G{row+1}"].alignment = openpyxl.styles.Alignment(horizontal="left", vertical="top", wrapText=True)
    
    if dest != "blob:":
        workbook.save(dest)
    else:
        iodata = io.BytesIO()
        workbook.save(iodata)
        iodata.seek(0)
        return iodata.read()
        
if __name__ == '__main__':
    print(get_img_jma(datetime.date(1998,6,26),"test1.jpg"))
    print(get_img_100(datetime.date(1998,6,26),"test2.jpg"))
    
    args = sys.argv
    if len(args) < 3:
        print("Usage: ./auto_tenkistudy_xlsx.py [IN_TSF] [OUT_XLSX] (IMG_QUALITY = 20)")
        exit()
    data = []
    with open(args[1], "r", encoding='cp932') as f:
        for row in csv.reader(f):
            data.append(row)
            #print(f"{type(row)}")
    
    if len(args) >= 4:
        q = int(args[3])
    else:
        q = 20
        
    #print(data)
    tsf_to_xlsx(data, args[2], q)
    
