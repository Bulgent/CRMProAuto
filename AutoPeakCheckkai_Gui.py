import cv2
import csv
import matplotlib.pyplot as plt
import numpy as np
import pyautogui as gui
import subprocess
import time
import win32gui
from PIL import ImageGrab, Image, ImageDraw, ImageFont
import glob
import os
import pathlib
import argparse
import pyperclip
import shutil
from playsound import playsound
import random
import datetime
import math


commencer = time.time()
#定義値-------------------------------------------------------------------------------------------------------------
#1回の最大処理数
size = 30
#基準点
#0.0秒
zes = [274,438]
#7.0秒
nis = [1335,438]
#一秒の座標幅
uns = (nis[0]-zes[0])/7
#スクリーンショットのトリミング範囲
xmin,xmax = 278,1333
ymin,ymax = 109,432
#ピーク範囲(heighttr/searx = 傾き)
#高さ閾値
heightthr = 1
#検定x幅
searx = 6
#ピーク修正範囲(±secrange秒修正する)
secrange = 0.3
#ピーク修正範囲をピクセルに変換
trange = round(secrange*uns)
#ピークの底を決める。testxが幅、thtestが高さ
testx = 3
thrtest = 0
#ピーク範囲の高さ制限(ピークと底の内分点を決める。大きいとピーク範囲の矢印が底に行きやすい。2か3くらいでいいかも?)
equi = 6
#再チェック判定(ピーク始点終点の高さの差)
revalue = 15
#再チェック判定(ピーク始点終点のピークとのx座標の幅の差)
thrwidth = 12
#ピークチェックできる状態か確認
checkTimes = 2
isPath = False
#定義式--------------------------------------------------------------------------------------------------------------

#座標----------------------------------------------------------------------------------------------------------------
positionOpenFile = (28, 65)#
positionChannel2 = (393,595)#
positionFile = (36,36) #
positionContinuousExport = (75,207)#
positionOverwrite = (67,132)#
positionOutRange = (539, 1)#
positionPeakButton = (401,68)#
positionBaseBelowButton = (279,65)#
positionBaseAboveButton = (312,63)#
#座標----------------------------------------------------------------------------------------------------------------

#CMRProを最前面に
def CRMfront():
    try:
        memoapp = win32gui.FindWindow(None,"Chromato-PRO [波形解析]")
    except:
        time.sleep(8)
        memoapp = win32gui.FindWindow(None,"Chromato-PRO [波形解析]")
    time.sleep(5)
    try:
        win32gui.SetForegroundWindow(memoapp)
    except:
        time.sleep(8)
        memoapp = win32gui.FindWindow(None,"Chromato-PRO [波形解析]")
        time.sleep(6)
        win32gui.SetForegroundWindow(memoapp)

#CMRProのセッティング
def CRMSetting():
    global sup
    try:
        #アプリ起動
        sup = subprocess.Popen(r'C:\Program Files (x86)\runtime\Chromato-PRO\ChromatoPro.exe')
        time.sleep(4)
        #CMRProを最前面に
        CRMfront()
        time.sleep(2)
    except:
        #アプリ起動
        time.sleep(5)
        sup = subprocess.Popen(r'C:\Program Files (x86)\runtime\Chromato-PRO\ChromatoPro.exe')
        time.sleep(4)
        #CMRProを最前面に
        CRMfront()
        time.sleep(2)

#ファイルを開くを押す
def ClickOpenFile():
    gui.moveTo(positionOpenFile)
    gui.click()

#ファイル名入力
def WriteFileName(name):
    time.sleep(0.5)
    gui.write(name)
    time.sleep(3)
    gui.press("enter")

#上書き保存
def Save():
    gui.moveTo(positionFile)
    gui.click()
    time.sleep(0.1)
    gui.moveTo(positionOverwrite)
    gui.click()

#グラフの全座標取得(coodで辞書式管理)
def getallcoodinate(image):
    #画像の読み込み
    img = cv2.imread(image,cv2.IMREAD_COLOR)
    #トリミング
    img = img[ymin:ymax,xmin:xmax]
    #画像の範囲取得
    height, width, channels = img.shape[:3]
    #画像の読み込み
    img = cv2.imread(image,0)
    #トリミング
    img = img[ymin:ymax,xmin:xmax]
    #画像の平滑化
    blurs = cv2.blur(img,(3,3))
    if args.debug == True:
        cv2.imwrite("blur.png", blurs)
    #グラフの端を白塗り・下に余白追加(輪郭をサーチするため)
    blurs = cv2.line(blurs,(0,0),(0,height),(255,255, 255),10)
    blurs = cv2.line(blurs,(width,0),(width,height),(255,255, 255),10)
    blurs = cv2.copyMakeBorder(blurs,50,50,0,0,cv2.BORDER_CONSTANT,value=[255,255,255])
    if args.debug == True:
        cv2.imwrite("blur_cut.png", blurs)
    #画像の2極化
    retval, im_bw = cv2.threshold(blurs, 254, 255, cv2.THRESH_BINARY)
    if args.debug == True:
        cv2.imwrite("blur_cut.png", im_bw)
    contours, hierarchy = cv2.findContours(im_bw, cv2.RETR_LIST, cv2.CHAIN_APPROX_NONE)
    # 余計な次元(余計な輪郭)削除 (NumContours, 1, NumPoints) -> (NumContours, NumPoints)
    contours = [np.squeeze(cnt, axis=1) for cnt in contours]
    conts = contours[0]
        #x座標範囲確認
    global maxxc
    global minxc
    maxxc = 0
    minxc =1000
    for cont in conts:
        if cont[0] >= maxxc:
            maxxc = cont[0]
        if cont[0] <= minxc:
            minxc = cont[0]
    rangecheck = True
    #輪郭が取得できてるか確認(できてなければ画像加工後リトライ)
    if maxxc != width-5 and minxc != 5:
        blurs = cv2.blur(img,(10,10))
        cv2.imwrite("autopeak_blur_re.png", blurs)
        time.sleep(0.1)
        blurs = cv2.imread("autopeak_blur_re.png")
        rangecheck = False
    #足りてない部分に黒点追加して補完
        checkx = [False]*width
        for numx in range(1,width):
            for numy in range(1,height):
                if blurs.item(numy,numx,0)!= 255 or blurs.item(numy,numx,1)!= 255 or blurs.item(numy,numx,2)!= 255 :
                    checkx[numx-1] = True
                    break
        for numx in range(1,width):
            if checkx[numx-1] == False:
                blurs[height-1,numx] = [0,0,0]
                blurs[height-2,numx] = [0,0,0]
    #グラフの端を白塗り
        blurs = cv2.line(blurs,(0,0),(0,height),(255,255, 255),10)
        blurs = cv2.line(blurs,(width,0),(width,height),(255,255, 255),10)
    #グラフの端に余白追加
        blurs = cv2.copyMakeBorder(blurs,50,50,0,0,cv2.BORDER_CONSTANT,value=[255,255,255])
    #画像の2極化
        cv2.imwrite("autopeak_blurs_re.png",blurs)
        blurs = cv2.imread("autopeak_blurs_re.png")
        retval, im_bw = cv2.threshold(blurs, 254, 255, cv2.THRESH_BINARY)
        if args.debug == True:
            cv2.imwrite("blur_cut_re.png", im_bw)
        im_bw = cv2.cvtColor(im_bw,cv2.COLOR_BGR2GRAY)
        contours, hierarchy = cv2.findContours(im_bw, cv2.RETR_LIST, cv2.CHAIN_APPROX_NONE)
    # 余計な次元(余計な輪郭)削除 (NumContours, 1, NumPoints) -> (NumContours, NumPoints)
        contours = [np.squeeze(cnt, axis=1) for cnt in contours]
        conts = contours[0]
    #x座標範囲確認
        maxxc = 0
        minxc =1000
        for cont in conts:
            if cont[0] >= maxxc:
                maxxc = cont[0]
            if cont[0] <= minxc:
                minxc = cont[0]
    #一意に表示(Max)
    global cood
    coodmax = {}
    for cont in conts:
        if cont[0] in coodmax:
            if cont[1] > coodmax[cont[0]]:
                coodmax[cont[0]] = cont[1]
        else:
            coodmax[cont[0]] = cont[1]
    #一意に表示(Min)
    coodmin = {}
    for cont in conts:
        if cont[0] in coodmin:
            if cont[1] < coodmin[cont[0]]:
                coodmin[cont[0]] = cont[1]
        else:
            coodmin[cont[0]] = cont[1]
    #一意に表示(Ave)
    cood = {}
    for coox in coodmax.keys():
        cood[coox]=(coodmax[coox]+coodmin[coox])//2

    if args.debug == True:
        print(cood)
        print("----------------------------")
        print(maxxc,minxc)
        if rangecheck == True:
            img = cv2.imread("blur_cut.png")
        else:
            img = cv2.imread("blur_cut_re.png")
        for coox in cood.keys():
            img[cood[coox],coox] = [0,0,255]
        cv2.imwrite("cood.png", img)

#ピーク値の修正(トリミング画像内での)
def modifypeak(k):
    k = round(k*uns+zes[0]-xmin)
    global modifiedpx
    global modifiedpy
    try:
        prepeak = [k, cood[k]]
#修正範囲  
        modifiedpy = prepeak[1]
        for coox in range(prepeak[0]-trange,prepeak[0]+trange):
            if cood[coox] <= modifiedpy:
                modifiedpy = cood[coox]
                modifiedpx = coox
    except:
        modifiedpx = k
        modifiedpy = 1000000
        for i,j in cood.items():
            if modifiedpy > j:
                modifiedpy = j

#ピーク範囲の確定(トリミング画像内での)
def peakrange():
    global peakendx
    global peakstax

#ピーク下限調査(終点)
    for peakx in range(modifiedpx,maxxc-testx + 1):
        dist = cood[peakx + testx]-cood[peakx]
        thrhx = peakx +testx
        if dist < thrtest:
            break
    thrh = (modifiedpy + (equi-1)*cood[thrhx])/equi
#ピーク範囲(終点)
    for peakx in range(modifiedpx,maxxc-searx + 1):
        dist = cood[peakx + searx]-cood[peakx]
        peakendx = peakx +searx
        if dist < heightthr and cood[peakx]>thrh:
            break
#ピーク下限調査(始点)
    for peakx in range(modifiedpx,minxc + testx - 1):
        dist = cood[peakx - testx]-cood[peakx]
        thrhx = peakx - testx
        if dist < thrtest:
            break
    thrh = (modifiedpy + (equi-1)*cood[thrhx])/equi
#ピーク範囲(始点)
    for peakx in range(modifiedpx,minxc + searx,-1):
        dist = cood[peakx-searx]-cood[peakx]
        peakstax = peakx - searx
        if dist < heightthr and cood[peakx]>thrh:
            break

#ピークの矢印座標確認
def peakcodinate(name1, name2):
    ##ピーク座標の検出##
    #画像１
    img_path_1 = r".\{}".format(name1)
    #画像２
    img_path_2 = r".\{}".format(name2)
    # 画像読み込み
    img_1 = cv2.imread(img_path_1)
    img_1 = img_1[ymin:ymax,xmin:xmax]
    img_1 = cv2.cvtColor(img_1, cv2.COLOR_BGR2RGB)
    img_2 = cv2.imread(img_path_2)
    img_2 = img_2[ymin:ymax,xmin:xmax]
    img_2 = cv2.cvtColor(img_2, cv2.COLOR_BGR2RGB)
    # ピークチェックの色付け(2つの画像の差分を抽出)
    fgbg = cv2.createBackgroundSubtractorMOG2()
    fgmask = fgbg.apply(img_1)
    fgmask = fgbg.apply(img_2)
    cv2.imwrite(r".\autopeak_checked.png",fgmask)
    #色の2極化
    im = cv2.imread('autopeak_checked.png')
    im_gray = cv2.cvtColor(im, cv2.COLOR_BGR2GRAY)
    retval, im_bw = cv2.threshold(im_gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    # 輪郭の検出
    contours, hierarchy = cv2.findContours(im_bw, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
    #座標の検出(面積最大の差分についての重心から座標検出)
    areas = np.array(list(map(cv2.contourArea,contours)))
    max_idx = np.argmax(areas)
    max_area = areas[max_idx]
    result = cv2.moments(contours[max_idx])
    global x
    global y
    x = int(result["m10"]/result["m00"])
    y = int(result["m01"]/result["m00"])

#目視確認画像の作成
def PasteImages(lst, photoNum_x, margin):
    margin_pre = 10
    #複数行ある場合
    if len(lst) > photoNum_x:
        #画像サイズの確認
        image = Image.open(lst[0])
        width, height = image.size
        #行数の計算
        photoNum_y = math.ceil(len(lst)/photoNum_x)
        #キャンバスの作成
        canvas = Image.new("RGB", (width*photoNum_x + margin*(photoNum_x + 1), height*photoNum_y + margin*(photoNum_y + 1)),(189, 195, 201))
        for i in range(len(lst)):
            image = Image.open(lst[i])
            #座標計算(0番目から)
            x_image = i%photoNum_x
            y_image = i//photoNum_x
            #ペースト
            #Recheckを強調
            if "checkphoto_1_F" in os.path.basename(lst[i]) or "checkphoto_2_F" in os.path.basename(lst[i]):
                canvas_pre = Image.new("RGB", (width + margin_pre, height + margin_pre),(255, 0, 0))
                canvas_pre.paste(image,(int(margin_pre/2), int(margin_pre/2)))
                image = canvas_pre
                paste_x = margin*(x_image+1) + width * x_image - int(margin_pre/2)
                paste_y = margin*(y_image+1) + height * y_image - int(margin_pre/2)
                canvas.paste(image, (paste_x, paste_y))
                title = os.path.basename(lst[i]).replace("checkphoto_1_F_","").replace("checkphoto_2_F_","").replace(".png",".crm")
            else:
                canvas.paste(image, (margin*(x_image+1) + width * x_image, margin*(y_image+1) + height * y_image))
                title = os.path.basename(lst[i]).replace("checkphoto_1_","").replace("checkphoto_2_","").replace(".png",".crm")
            draw = ImageDraw.Draw(canvas)
            font = ImageFont.truetype('C:\windows\Fonts\meiryo.ttc', 27)
            draw.text((margin*(x_image+1) + width * x_image, \
                    margin*(y_image+1) + height * (y_image+1)),\
                    title,\
                    font = font, fill = "black")
    
    #1行だけの場合(行サイズ調整)
    else:
        #画像サイズの確認
        image = Image.open(lst[0])
        width, height = image.size
        #行数の計算
        photoNum_y = 1
        photoNum_x = len(lst)
        #キャンバスの作成
        canvas = Image.new("RGB", (width*photoNum_x + margin*(photoNum_x + 1), height*photoNum_y + margin*(photoNum_y + 1)),(189, 195, 201))
        for i in range(len(lst)):
            image = Image.open(lst[i])
            #座標計算(0番目から)
            x_image = i
            y_image = 0
            #ペースト
            #Recheckを強調
            if "checkphoto_1_F" in os.path.basename(lst[i]) or "checkphoto_2_F" in os.path.basename(lst[i]):
                canvas_pre = Image.new("RGB", (width + margin_pre, height + margin_pre),(255, 0, 0))
                canvas_pre.paste(image,(int(margin_pre/2), int(margin_pre/2)))
                image = canvas_pre
                paste_x = margin*(x_image+1) + width * x_image - int(margin_pre/2)
                paste_y = margin*(y_image+1) + height * y_image - int(margin_pre/2)
                canvas.paste(image, (paste_x, paste_y))
                title = os.path.basename(lst[i]).replace("checkphoto_1_F_","").replace("checkphoto_2_F_","").replace(".png",".crm")
            else:
                canvas.paste(image, (margin*(x_image+1) + width * x_image, margin*(y_image+1) + height * y_image))
                title = os.path.basename(lst[i]).replace("checkphoto_1_","").replace("checkphoto_2_","").replace(".png",".crm")
            draw = ImageDraw.Draw(canvas)
            font = ImageFont.truetype('C:\windows\Fonts\meiryo.ttc', 40)
            draw.text((margin*(x_image+1) + width * x_image, \
                    margin*(y_image+1) + height * (y_image+1)),\
                    title,\
                    font = font, fill = "black")
    return canvas

#目視確認画像の作成
def ManageCheckPhotos(name, path):
    margin = 96
    lst = glob.glob(path)
    photoNum_x = 4
    photoNum_y_max = 4

    #貼り付け最大枚数以内の時
    if len(lst) <= photoNum_x * photoNum_y_max:
        im = PasteImages(lst, photoNum_x, margin)
        im.save(f"{name}.png")

    #貼り付け最大枚数より多い時
    else:
        itr = len(lst) // (photoNum_x * photoNum_y_max)
        itr_con = photoNum_x * photoNum_y_max
        for i in range(itr+1):
            fin = min((i+1)*itr_con, len(lst))
            lst_image = lst[i*itr_con:fin]
            im = PasteImages(lst_image, photoNum_x, margin)
            im.save(f"{name}_{i}.png")

#目視確認画像の作成
def CreateCheckPhotos():
    ManageCheckPhotos("Channel1", r".\checkphoto_1_*")
    if isch4s == True:
        ManageCheckPhotos("Channel2", r".\checkphoto_2_*")

#外部フォルダからCRMtoEXCELフォルダにコピー
def copyfiles_start():
    path_CRMtoEXCEL = os.getcwd()
    files = glob.glob(r"{}/*.crm".format(input_path))
    for file in files:
        shutil.copy(file,f"{path_CRMtoEXCEL}")

#CRMtoEXCELフォルダから外部フォルダにコピー
def copyfiles_end():
    path_CRMtoEXCEL = os.getcwd()
    files = glob.glob(r"{}/*.crm".format(path_CRMtoEXCEL)) + glob.glob(r"{}/Channel*.png".format(path_CRMtoEXCEL))
    os.makedirs(output_path, exist_ok=True) 
    for file in files:
        try:
            shutil.move(file,f"{output_path}")
        except:
            os.remove(f"{output_path}/{os.path.basename(file)}")
            shutil.move(file,f"{output_path}")
    if os.path.exists(r"{}/Recheck".format(path_CRMtoEXCEL)):
        shutil.move("{}/Recheck".format(path_CRMtoEXCEL),f"{output_path}/Recheck")

#クロマトグラフの最大化
def maximizeCRM():
    path = r".\images\maximizeButton.png"
    gui.click(gui.locateCenterOnScreen(path, confidence = 0.7, grayscale = True))

#入力ボックスが消えてピークチェックできる状態になっているか確認
def inputCheck():
    imagePath = r"./images/inputButton.png"
    check = gui.locateOnScreen(imagePath, confidence = 0.7, grayscale = True)
    #入力ボタンがあるとTrueで返すので反転
    #Trueなら次に進んでよし
    return not check

def pathCheck():
    global isPath
    if isPath:
        return
    imagePath = r"./images/checkPath.png"
    check = gui.locateOnScreen(imagePath, confidence = 0.9, grayscale = True)
    if not check:
        raise NameError("ファイルを開く -> ファイルの場所が「CRMProAuto」じゃないのでは??")
    else:
        isPath = True

#CH1, CH2, 7min, 9min, path
def main_APC(_CH1, _CH2, _isCH2, _is9min, _isPath, _inputPath, _outputPath):
    #音声データ
    startsound = glob.glob(r"sound/start*")
    endsound = glob.glob(r"sound/end*")

    #再チェック
    yabatani = []
    #再チェックファイル掃除
    pre_files = glob.glob("./Recheck/*")
    if bool(pre_files) == True:
        try:
            os.mkdir("./Recheck/previous")
        except:
            pass
        for file in pre_files:
            try:
                shutil.move(f"{file}","./Recheck/previous")
            except shutil.Error:
                os.remove(file)

    #デバッグ用
    parser = argparse.ArgumentParser()
    parser.add_argument("--debug",action = "store_true", help="debug時に使用(デバッグ用)")
    parser.add_argument("--first",action = "store_true", help="n2o見る時に使用(デバッグ用)")
    parser.add_argument("--axis",action = "store_true", help="横軸が9秒までの時にご使用ください")
    parser.add_argument("--path",action = "store_true", help="CRMtoEXCELフォルダ以外のファイルを実行したい時に使用して下さい。")
    parser.add_argument("--toEXCEL",action = "store_true", help="AutoPeakCheck後、そのままCRMtoEXCELを行ないます。")
    global args, isch4s, output_path, input_path
    args = parser.parse_args()

    #フォルダの掃除
    if _isPath ==True:
        pre_files = glob.glob("*.crm") + glob.glob("*.xlsx")
        if bool(pre_files) == True:
            try:
                os.mkdir("./previous")
            except:
                pass
            for file in pre_files:
                try:
                    shutil.move(f"{file}","./previous")
                except shutil.Error:
                    os.remove(file)
    for i in glob.glob("Channel*.png"):
        os.remove(i)

    if _is9min == True:
        #一秒の座標幅
        uns = (nis[0]-zes[0])/9
        #ピーク修正範囲をピクセルに変換
        trange = round(secrange*uns)

    #ピーク範囲の設定
    print("Chanel 2を使用しない場合はChannel 2に-と入力してください")
    n2os = float(_CH1)
    #チャネル2の処理を行うか否か
    if not _isCH2:
        isch4s = False
        print("Channel_1 のみ処理します。")
    else:
        ch4s = float(_CH2)
        isch4s = True
    
    #パスの確認
    if _isPath == True:
        input_path = _inputPath.replace('"','')
        output_path = _outputPath.replace('"','')
        if not os.path.exists(input_path):
            raise Exception("処理するフォルダが見当たりません")
        if not os.path.exists(output_path):
            print("出力するフォルダが見当たらないので作成します")
            os.makedirs(output_path, exist_ok=True)

        copyfiles_start()

    soundsta = random.randrange(len(startsound))
    playsound(startsound[soundsta])
    #フォルダ内のcrmファイルを抽出
    all_files = glob.glob("*.crm")



    #繰り返し数
    itr = -(-len(all_files)//size)
    con = 0
    errorCon = 0

    #ファイル何個かずつで繰り返し
    lst = []
    for (start, count) in zip(range(0, len(all_files),size),range(1,itr+1)):
        files = all_files[start:start+size]

        #CRMProの起動
        CRMSetting()

        #全画面表示にする
        gui.hotkey("win","up")
        time.sleep(5)

        #crmファイルの繰り返し処理
        for file, countist in zip(files,range(0,len(files)+1)):

            #ファイルのパス取得
            file_p = pathlib.Path(file)
            file_ab = str(file_p)

            #ファイルを開くを押す
            ClickOpenFile()
            time.sleep(2.5)
            pathCheck()

            #ファイル名入力
            pyperclip.copy(file_ab)
            gui.hotkey("ctrl","v")
            time.sleep(1.5)
            gui.press("enter")
            time.sleep(3)

            #ファイル入力画面が閉じているか確認
            for checkCount in range(checkTimes):
                if not inputCheck():
                    time.sleep(3)
                    gui.hotkey("ctrl","v")
                    time.sleep(1.5)
                    gui.press("enter")

            #一回目だけクロマトグラフの最大化
            if countist == 0:
                maximizeCRM()
                time.sleep(3)
            
            """N2Oの処理"""
            #スクリーンショット
            gui.moveTo(positionOutRange)
            ImageGrab.grab().save("autopeak_pre.png")
            #グラフの全座標取得
            getallcoodinate("autopeak_pre.png")
            
            #ピーク値の修正
            modifypeak(n2os)
            #ピーク範囲の確定
            try:
                peakrange()
            except:
                yabatani.append(file)
                print(f"{file}は多分解析済の為処理していません")
                gui.moveTo(positionOutRange)
                name = os.path.basename(file).replace(".crm","")
                ImageGrab.grab().crop((xmin,ymin,xmax,ymax)).save(f"checkphoto_1_F_{name}.png")
                Image.new("L",(xmax-xmin,ymax-ymin)).save(f"checkphoto_2_F_{name}.png")
                con += 1
                errorCon += 1
                continue
            ###ピークをつける###
            gui.click(positionPeakButton)
            gui.moveTo(modifiedpx + xmin,zes[1]-20)
            time.sleep(1)
            gui.doubleClick()
            #スクリーンショット
            gui.moveTo(positionOutRange)
            ImageGrab.grab().save("autopeak_before1.png")

            ###ベーススタート/エンド###
            gui.click(positionBaseBelowButton)
            time.sleep(0.5)
            #スクリーンショット
            gui.moveTo(positionOutRange)
            ImageGrab.grab().save("autopeak_after1.png")
            #ピーク座標矢印の検出(失敗なら片づけて次へ)
            try:
                peakcodinate("autopeak_before1.png","autopeak_after1.png")
            except:
                yabatani.append(file)
                file_list = glob.glob("autopeak*.png")
                for i in file_list:
                    os.remove(i)
                #確認スクショ作成
                gui.moveTo(positionOutRange)
                name = os.path.basename(file).replace(".crm","")
                ImageGrab.grab().crop((xmin,ymin,xmax,ymax)).save(f"checkphoto_1_F_{name}.png")
                Image.new("L",(xmax-xmin,ymax-ymin)).save(f"checkphoto_2_F_{name}.png")
                con += 1
                errorCon += 1
                continue

            #カーソルの移動
            gui.moveTo(x + xmin, y + ymin)
            time.sleep(2)
            gui.dragTo(peakendx + xmin, y+ymin, 0.75)
            gui.moveTo(x + xmin, y + ymin - 3)
            gui.dragTo(peakstax + xmin, y+ymin, 0.75)
            #スクリーンショット
            gui.click(positionPeakButton)
            gui.moveTo(positionOutRange)
            ImageGrab.grab().save("autopeak_before2.png")

            ###ピークスタート/エンド###
            gui.click(positionBaseAboveButton)
            time.sleep(0.5)
            #スクリーンショット
            gui.moveTo(positionOutRange)
            ImageGrab.grab().save("autopeak_after2.png")
            #ピーク座標の検出#
            peakcodinate("autopeak_before2.png","autopeak_after2.png")
            #カーソルの移動
            gui.moveTo(x + xmin, y + ymin)
            time.sleep(2)
            gui.dragTo(peakendx + xmin, y+ymin, 0.75)
            gui.moveTo(x + xmin, y + ymin + 3)
            gui.dragTo(peakstax + xmin, y+ymin, 0.75)

            #スクショの削除
            file_list = glob.glob("autopeak*.png")
            for i in file_list:
                os.remove(i)
            if args.debug != True:
                time.sleep(0.5)
                Save()
                time.sleep(0.5)
                #確認スクショ作成
            #再チェック判定
            if abs(cood[peakendx]-cood[peakstax])>revalue or abs(abs(modifiedpx-peakendx)-abs(modifiedpx-peakstax)) > thrwidth:
                yabatani.append(file)
                gui.moveTo(positionOutRange)
                name = os.path.basename(file).replace(".crm","")
                ImageGrab.grab().crop((xmin,ymin,xmax,ymax)).save(f"checkphoto_1_F_{name}.png")
            else:
                gui.moveTo(positionOutRange)
                name = os.path.basename(file).replace(".crm","")
                ImageGrab.grab().crop((xmin,ymin,xmax,ymax)).save(f"checkphoto_1_{name}.png")
            if args.first == True:
                os._exit()

            #CH4の処理############
            if isch4s == True:
                #チャネルの切り替え
                gui.click(positionChannel2)
                time.sleep(0.5)
                #スクリーンショット
                gui.moveTo(positionOutRange)
                ImageGrab.grab().save("autopeak_pre.png")
                #グラフの全座標取得
                getallcoodinate("autopeak_pre.png")
                #ピーク値の修正
                modifypeak(ch4s)
                #ピーク範囲の確定
                try:
                    peakrange()
                except:
                    yabatani.append(file)
                    gui.moveTo(positionOutRange)
                    ImageGrab.grab().crop((xmin,ymin,xmax,ymax)).save(f"checkphoto_2_F_{name}.png")
                    con += 1
                    errorCon += 1
                    continue

                ###ピークをつける###
                gui.click(positionPeakButton)
                gui.moveTo(modifiedpx + xmin,zes[1]-20)
                time.sleep(1)
                gui.doubleClick()
                #スクリーンショット
                gui.moveTo(positionOutRange)
                ImageGrab.grab().save("autopeak_before1.png")

                ###ベーススタート/エンド###
                gui.click(positionBaseBelowButton)
                time.sleep(0.5)
                #スクリーンショット
                gui.moveTo(positionOutRange)
                ImageGrab.grab().save("autopeak_after1.png")
                ##ピーク座標の検出##
                try:
                    peakcodinate("autopeak_before1.png","autopeak_after1.png")
                except:
                    yabatani.append(file)
                    gui.moveTo(positionOutRange)
                    ImageGrab.grab().crop((xmin,ymin,xmax,ymax)).save(f"checkphoto_2_F_{name}.png")
                    con += 1
                    errorCon += 1
                    continue

                #カーソルの移動
                gui.moveTo(x + xmin, y + ymin)
                time.sleep(2)
                gui.dragTo(peakendx + xmin, y+ymin, 0.75)
                gui.moveTo(x + xmin, y + ymin - 3)
                gui.dragTo(peakstax + xmin, y+ymin, 0.75)
                #スクリーンショット
                gui.click(positionPeakButton)
                gui.moveTo(positionOutRange)
                ImageGrab.grab().save("autopeak_before2.png")

                ###ピークスタート/エンド###
                gui.click(positionBaseAboveButton)
                time.sleep(0.5)
                #スクリーンショット
                gui.moveTo(positionOutRange)
                ImageGrab.grab().save("autopeak_after2.png")
                #ピーク座標の検出#
                peakcodinate("autopeak_before2.png","autopeak_after2.png")
                #カーソルの移動
                gui.moveTo(x + xmin, y + ymin)
                time.sleep(2)
                gui.dragTo(peakendx + xmin, y+ymin, 0.75)
                gui.moveTo(x + xmin, y + ymin + 3)
                gui.dragTo(peakstax + xmin, y+ymin, 0.75)

                #スクショの削除
                file_list = glob.glob("autopeak*.png")
                for i in file_list:
                    os.remove(i)
                
                #確認スクショ作成
                #再チェック判定
                if abs(cood[peakendx]-cood[peakstax])>revalue or abs(abs(modifiedpx-peakendx)-abs(modifiedpx-peakstax)) > thrwidth:
                    yabatani.append(file)
                    gui.moveTo(positionOutRange)
                    ImageGrab.grab().crop((xmin,ymin,xmax,ymax)).save(f"checkphoto_2_F_{name}.png")
                else:
                    gui.moveTo(positionOutRange)
                    ImageGrab.grab().crop((xmin,ymin,xmax,ymax)).save(f"checkphoto_2_{name}.png")
                #上書き保存
                if args.debug != True:
                    Save()
        time.sleep(1)
        #クロマトを閉じる
        sup.kill()


    CreateCheckPhotos()
        #確認スクショ削除
    for i in glob.glob("checkphoto*.png"):
        os.remove(i)
        
    if bool(yabatani) == True:
        #再チェック
        try:
            os.mkdir("Recheck")
        except:
            pass
        yabatani = list(set(yabatani))
        for file in yabatani:
            shutil.copy(file,"./Recheck")

    if _isPath == True:
        copyfiles_end()



    time.sleep(3)

    pro = time.time() - commencer
    soundend = random.randrange(len(endsound))
    playsound(endsound[soundend])

    #ログに書き込み
    try:
        with open(r"./data/APC-log.csv","a") as f:
            f.write(f"{datetime.datetime.now()},{len(all_files)},{pro},{errorCon}\n")
    except:
        with open(r"./data/APC-log.csv","w") as f:
            f.write(f"date, データ数 \n {datetime.datetime.now()},{len(all_files)},{pro},{errorCon}\n")

    print("Complete")

    pro_hour = pro//3600
    pro_minu = (pro - pro_hour * 3600)//60
    pro_sec = pro-pro_minu*60-pro_hour*3600

    print("経過時間は{:.0f}時間{:.0f}分{:.0f}秒です".format(pro_hour,pro_minu,pro_sec))

    if bool(yabatani) == True:
        print("Recheckフォルダのファイルはちゃんと確認した方が良いかも!!")
        print("※Recheckフォルダには解析後のコピーファイルが入っています。")
        #print("※CO2を処理した場合、Recheckフォルダには解析後の全てのコピーファイルが入っています。")