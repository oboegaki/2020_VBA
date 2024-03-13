Attribute VB_Name = "play"
Option Explicit

'---------------------------------------------------------------------------
'自動操作を開始する : 毎日3:15
'---------------------------------------------------------------------------
Sub auto_play()
   Shell "cmd /c C:\Windows\System32\schtasks.exe /run /tn wifiタスク名"
    Sleep 20000
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    S001.run
    Debug.Print "S001_ok"
    
    S002.run
    Debug.Print "S002_ok"

    ThisWorkbook.Save
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    Application.Wait Now + TimeValue("0:00:05")
    Shell "C:\Windows\System32\cmd.exe /c C:\〇〇\sleep.bat"
End Sub

'---------------------------------------------------------------------------
'手動で操作する
'---------------------------------------------------------------------------
Sub onplay()

    S001.run
    Debug.Print "S001_ok"
    
    S002.run
    Debug.Print "S002_ok"
    
    ThisWorkbook.Save
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    With CreateObject("WScript.Shell").popup("手動操作がしました", 20, "終了", vbInformation): End With
End Sub

'---------------------------------------------------------------------------
'全てのエクセルWindowを強制終了
'---------------------------------------------------------------------------
Private Sub Excel_Task_Kill()
    Application.DisplayAlerts = False
    ThisWorkbook.Save
    Application.DisplayAlerts = True
    Sleep 3000

    Shell "taskkill /F /IM excel.exe"

End Sub

'---------------------------------------------------------------------------
'特定のタスク名を終了場合
'---------------------------------------------------------------------------
Private Sub targetTaskKill()

    Dim wdApp   As New Word.Application
    Dim wdApp As Object: Set wdApp = CreateObject("Word.Application")
    Dim wdtasks As Object: Set wdtasks = wdApp.Tasks
    Dim i As Long

    For i = 1 To wdtasks.count
        DoEvents
        If wdtasks(i).Visible = True And InStr(wdtasks(i).Name, "タスク名") > 0 Then
                wdtasks(i).Kill
        End If
    Next i

    Set wdtasks = Nothing                        '解放
    wdApp.Quit                                   '解放しないとタスクが残ってしまう！    Set wdApp = Nothing                          '解放


End Sub

