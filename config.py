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

pathCRMPro = r'C:\Program Files (x86)\runtime\Chromato-PRO\ChromatoPro.exe'