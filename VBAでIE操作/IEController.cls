VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IEController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private IE As SHDocVw.InternetExplorer
Private Sub Class_Initialize()
    'Dim IE As SHDocVw.InternetExplorer: Set IE = New SHDocVw.InternetExplorer
    Debug.Print "initalized"
    'Set IE = New SHDocVw.InternetExplorer
    'Dim IE As SHDocVw.InternetExplorer: Set IE = New SHDocVw.InternetExplorer
End Sub

'------------------------------------
'IEオブジェクト作成
'------------------------------------
Function SetUrl(url As String) As IWebBrowser2
    Dim IE As SHDocVw.InternetExplorer: Set IE = New SHDocVw.InternetExplorer
    'Dim IE As Object: Set IE = CreateObject("InternetExplorer.Application")
    IE.Navigate url
    
    '戻り値
    Set SetUrl = IE
End Function

'------------------------------------
'呼び込むまで待機
'------------------------------------
Sub loading(IE As IWebBrowser2)

    'サイトの読み込み待ち
    IE.Visible = False
    Dim waitMax As Long: waitMax = 0
    Do While IE.Busy = True Or IE.ReadyState <> 4
        DoEvents
        Application.Wait Now + TimeValue("0:00:01")
        waitMax = waitMax + 1
        If waitMax > 20 Then Exit Do
    Loop
    IE.Visible = True
    
    'JavaScript読み込み待ち
    Application.Wait Now + TimeValue("0:00:03")
    
End Sub


