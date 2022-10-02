from utils import *

# デバッグ用
parser = argparse.ArgumentParser()
parser.add_argument("--debug",action = "store_true", help="debug時に使用(デバッグ用)")
parser.add_argument("--first",action = "store_true", help="n2o見る時に使用(デバッグ用)")
parser.add_argument("--axis",action = "store_true", help="横軸が9秒までの時にご使用ください")
parser.add_argument("--path",action = "store_true", help="CRMtoEXCELフォルダ以外のファイルを実行したい時に使用して下さい。")
parser.add_argument("--toEXCEL",action = "store_true", help="AutoPeakCheck後、そのままCRMtoEXCELを行ないます。")
args = parser.parse_args()

# 音声データ
startsound = glob.glob(r"sound/start*")
endsound = glob.glob(r"sound/end*")

# 再チェック
yabatani = []

#↓↓↓↓↓↓↓↓↓フォルダの掃除↓↓↓↓↓↓↓↓↓###########################################
# Recheckフォルダ掃除
pre_files = glob.glob("./Recheck/*")
if bool(pre_files):
    try:
        os.mkdir("./Recheck/previous")
    except:
        pass
    for file in pre_files:
        try:
            shutil.move(f"{file}","./Recheck/previous")
        except shutil.Error:
            os.remove(file)

# ピーク写真の削除
for i in glob.glob("Channel*.png"):
    os.remove(i)

# フォルダの掃除(path利用時)
if args.path:
    pre_files = glob.glob("*.crm") + glob.glob("*.xlsx")
    if bool(pre_files):
        try:
            os.mkdir("./previous")
        except:
            pass
        for file in pre_files:
            try:
                shutil.move(f"{file}","./previous")
            except shutil.Error:
                os.remove(file)
#↑↑↑↑↑↑↑↑↑↑フォルダの掃除↑↑↑↑↑↑↑↑↑↑########################################


#↓↓↓↓↓↓↓↓↓入力データの受け取り↓↓↓↓↓↓↓↓↓###########################################
# ピーク範囲の設定
print("Chanel 2を使用しない場合はChannel 2に-と入力してください")
n2os = float(input("Channel_1 のピーク時間を入力して下さい:"))
ch4s = input("Channel_2 のピーク時間を入力して下さい:")

#チャネル2の処理を行うか否か
if ch4s == "-":
    isch4s = False
    print("Channel_1 のみ処理します。")
else:
    ch4s = float(ch4s)
    isch4s = True

# パスの確認
if args.path:
    input_path = input("処理する.crmファイルが含まれているフォルダの絶対パスを入力して下さい。:").replace('"','')
    output_path = input("処理後.crmファイルを保存するフォルダの絶対パスを入力して下さい。:").replace('"','')
    if not os.path.exists(input_path):
        raise Exception("処理するフォルダが見当たりません")
    if not os.path.exists(output_path):
        raise Exception("出力するフォルダが見当たりません")
    copyfiles_start(input_path)
#↑↑↑↑↑↑↑↑↑↑入力データの受け取り↑↑↑↑↑↑↑↑↑↑########################################

# 軸の変換
if args.axis:
    #一秒の座標幅(再宣言)
    uns = (nis[0]-zes[0])/9
    #ピーク修正範囲をピクセルに変換(再宣言)
    trange = round(secrange*uns)

# サウンドを再生
soundsta = random.randrange(len(startsound))
playsound(startsound[soundsta])
#フォルダ内のcrmファイルを抽出
all_files = glob.glob("*.crm")

#繰り返し数
itr = -(-len(all_files)//size)
count = 0
errorCount = 0

#ファイル何個かずつで繰り返し(一気にやるとアプリが重くなる)
lst = []
for (start, count) in zip(range(0, len(all_files),size),range(1,itr+1)):
    # 実行するファイルを抽出
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
        maxxc, minxc, cood = getAllCoodinate("autopeak_pre.png", args.debug)
        
        #ピーク値の修正
        modifiedpx, modifiedpy = modifypeak(n2os, cood)

        #ピーク範囲の確定
        try:
            peakendx, peakstax = peakrange(maxxc, minxc, cood, modifiedpx, modifiedpy)
        except:
            yabatani.append(file)
            print(f"{file}は多分解析済の為処理していません")
            gui.moveTo(positionOutRange)
            name = os.path.basename(file).replace(".crm","")
            ImageGrab.grab().crop((xmin,ymin,xmax,ymax)).save(f"checkphoto_1_F_{name}.png")
            Image.new("L",(xmax-xmin,ymax-ymin)).save(f"checkphoto_2_F_{name}.png")
            count += 1
            errorCount += 1
            continue

        ###ピークをつける###
        gui.click(positionPeakButton)
        time.sleep(1)
        gui.doubleClick(modifiedpx + xmin, zes[1] - 20)
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
            count += 1
            errorCount += 1
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
            getAllCoodinate("autopeak_pre.png")
            #ピーク値の修正
            modifypeak(ch4s)
            #ピーク範囲の確定
            try:
                peakrange()
            except:
                yabatani.append(file)
                gui.moveTo(positionOutRange)
                ImageGrab.grab().crop((xmin,ymin,xmax,ymax)).save(f"checkphoto_2_F_{name}.png")
                count += 1
                errorCount += 1
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
                count += 1
                errorCount += 1
                continue

            #カーソルの移動
            gui.moveTo(x + xmin, y + ymin)
            time.sleep(2)
            gui.dragTo(peakendx + xmin, y+ymin, 0.75)
            gui.moveTo(x + xmin, y + ymin - 5)
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

if args.path == True:
    copyfiles_end()



time.sleep(3)

pro = time.time() - commencer
soundend = random.randrange(len(endsound))
playsound(endsound[soundend])

#ログに書き込み
try:
    with open(r"./data/APC-log.csv","a") as f:
        f.write(f"{datetime.datetime.now()},{len(all_files)},{pro},{errorCount}\n")
except:
    with open(r"./data/APC-log.csv","w") as f:
        f.write(f"date, データ数 \n {datetime.datetime.now()},{len(all_files)},{pro},{errorCount}\n")

print("Complete")

pro_hour = pro//3600
pro_minu = (pro - pro_hour * 3600)//60
pro_sec = pro-pro_minu*60-pro_hour*3600

print("経過時間は{:.0f}時間{:.0f}分{:.0f}秒です".format(pro_hour,pro_minu,pro_sec))

if bool(yabatani) == True:
    print("Recheckフォルダのファイルはちゃんと確認した方が良いかも!!")
    print("※Recheckフォルダには解析後のコピーファイルが入っています。")
    #print("※CO2を処理した場合、Recheckフォルダには解析後の全てのコピーファイルが入っています。")
    