Attribute VB_Name = "S000_GoogleTrend"
Option Explicit

Dim col As Collection
Sub run()
    'initialize
    Set col = New Collection
    lib.IE_ImageOff
    
    Call getTrend     'Googleトレンドワード取得
End Sub

'----------------------------------------------------------
' STEP1 : Googleトレンドワード取得
'----------------------------------------------------------
Private Sub getTrend()

    'Application.StatusBar = "トレンドワード起動中..."
    lib.IE_OPEN "https://trends.google.co.jp/trends/trendingsearches/daily?geo=JP", IE
    
    'HTMLエレメントセット
    Dim colTrends As IHTMLElementCollection
    Set colTrends = IE.Document.getElementsByClassName("md-list-block")

    Dim i As Long
    For i = 0 To colTrends.Length - 1
        'regexのセット・設定
        Dim regex: Set regex = CreateObject("VBScript.RegExp")
        regex.Pattern = "(\S+)"
        regex.Global = True
        regex.MultiLine = True
        
        'innerTextからトレンドワードを抽出・蓄積
        Dim result: Set result = regex.Execute(colTrends(i).innerText)
        If result.count > 2 Then
            col.Add result(1).SubMatches(0)
        End If
        Set regex = Nothing
    Next
    
    'トレンドワードが拾えなければweb検索へ
    If col.count < 1 Then
        IE.Quit
        Sleep (10000): Set IE = Nothing: lib.AllcloseIEWindow
        Exit Sub
    End If
    
    '固定文言をcollectionに追加
        With col
        .Add "新着ニュース"
        .Add "今朝の天気"
        .Add "交通情報"
        .Add "本日の日付"
    End With
    
    '格納したものをトレンドワードシートへ
    Dim idx As Long
    'ThisWorkbook.Sheets("トレンドワード").Cells.Clear
    For idx = 1 To col.count
        ThisWorkbook.Sheets("トレンドワード").Cells(idx, 1) = col.Item(idx)
    Next idx
    
    'リソース解放
    Set col = Nothing
    IE.Quit
    Sleep (10000): Set IE = Nothing: lib.AllcloseIEWindow
End Sub
 

