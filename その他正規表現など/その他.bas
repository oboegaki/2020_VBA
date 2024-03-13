Attribute VB_Name = "Module1"
Option Explicit


'*******************************************************************
'*  　 numExtract
'*          数字のみ抽出
'*  引数：  StringValue:数字のみ抽出したい対象文字列(String)
'******************************************************************
Function numExtract(StringValue As String) As String
 
  '変数の準備
  Dim i As Integer
  Dim numText As String
 
  'Len で文字数が分かる ＝ 1 to 文字列（最終文字）まで For Nextで影響させる。
  For i = 1 To Len(StringValue)
     'Midで文字列を左から順にnumTextに格納
      numText = Mid(StringValue, i, 1)
    'もしnumText[0-9]に該当する場合 変数numExtractに格納していく
      If numText Like "[0-9]" Then: numExtract = numExtract & numText
  Next i
 
End Function


'**********************************
'* 文字一括連結(選択範囲の文字を一つに連結)
'**********************************
Function concatCells(targetRng As Range) As String
    '適当な範囲指定して､文字を連結
    concat_cellss = WorksheetFunction.Concat(targetRng)
End Function




'******************************
'* 正規表現によるマッチングしたセルの値を取得
'******************************
Function cellRegValue(pattern As String) As String
     Dim rng As Range
     For Each rng In Range("A1:G15")
            '正規表現によるマッチング
            With CreateObject("VBScript.RegExp")
                .Global = True  '全文字検索
                .pattern = "\d\d\d\d\d\d\d"
                
                
                If .test(rng) Then
                    cellRegValue = rng.Value
                    Exit For
                End If
            End With
     Next

End Function

'******************************
'* VBSファイルで、別プロセスでメッセージボックスを表示させたい場合
'******************************

Sub excute()

    '変数宣言
    Dim filePath As String
    Dim fileNo As Integer
    Dim msg As Integer
    Dim re As Integer
    
    '作成するファイルパスを指定
    filePath = "C:\〇〇\msgbox.vbs"

    'vbsファイル起動
    With CreateObject("Wscript.Shell")
        re = .Run(Command:=filePath, WaitOnReturn:=True)
    End With
    
    'Yesなら処理開始
    If re = vbYes Then
        '自動入力開始
        With CreateObject("Wscript.Shell")
             Call Application.Wait(Now + TimeValue("00:00:03")) '// 1秒停止
             
             '処理内容
        End With
    End If
End Sub












