Attribute VB_Name = "SP_SENDKEY"
Option Explicit


'マウス操作(https://thom.hateblo.jp/entry/2015/11/21/002304)
Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Declare Sub mouse_event Lib "user32" ( _
    ByVal dwFlags As Long, _
    Optional ByVal dx As Long = 0, _
    Optional ByVal dy As Long = 0, _
    Optional ByVal dwDate As Long = 0, _
    Optional ByVal dwExtraInfo As Long = 0)

'Sub マウスで画面の任意の位置をクリック()
'    SetCursorPos 100, 35  '左から100ピクセル、上から35ピクセルの位置にカーソルを移動
'    'mouse_event 2  '左ボタン押下のコード
'    'mouse_event 4  '左ボタン解放のコード
'End Sub



'強制的に最前面にさせる
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
'最小化されているか調べる
Private Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long
'元のサイズに戻す
Private Declare Function ShowWindowAsync Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

'public用宣言
Dim IE As Object
Dim pointlinks
Dim clicklinks


'クリック操作
Sub clickPoint()
    mouse_event 2  '左ボタン押下のコード
    mouse_event 4  '左ボタン解放のコード
End Sub


'HOMEボタンをクリック
Sub homePos()
    SetCursorPos 952, 911: clickPoint: Sleep (2000)
End Sub
'戻るボタンをクリック
Sub backPos()
    SetCursorPos 839, 916  '左から839ピクセル、上から916ピクセルの位置にカーソルを移動
    clickPoint
End Sub
'動画フォルダをクリック
Sub DFolderPos()
    Sleep (5000)
    SetCursorPos 945, 505: clickPoint: Sleep (6000)
End Sub
'ボタンを押して待機
Sub repeat(x As Long, y As Long)
    Dim maxSec As Long: maxSec = 0
    SetCursorPos x, y: clickPoint
    While (maxSec <> 180)
        DoEvents
        Sleep (1000)
        maxSec = maxSec + 1
    Wend
    maxSec = 0
End Sub


'---------------------------------
' STEP : 1
'---------------------------------
Sub DIVIDE()
    DFolderPos                                                  'フォルダを開く
    SetCursorPos 857, 395: clickPoint: Sleep (10000)       'ゲーム起動
    Call repeat(782, 209): Call backPos: Sleep (5000)       '1押下、待機して戻る
    Call repeat(782, 209): Call backPos: Sleep (5000)       '2押下、待機して戻る
    Call repeat(782, 209): Call backPos: Sleep (5000)       '3押下、待機して戻る
    Call repeat(782, 209): homePos                          '4広押下、160秒待機してHOME(終了)
End Sub


