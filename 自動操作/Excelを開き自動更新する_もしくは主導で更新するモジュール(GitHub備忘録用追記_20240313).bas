Attribute VB_Name = "play"
Option Explicit

'---------------------------------------------------------------------------
'����������J�n���� : ����3:15
'---------------------------------------------------------------------------
Sub auto_play()
   Shell "cmd /c C:\Windows\System32\schtasks.exe /run /tn wifi�^�X�N��"
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
    Shell "C:\Windows\System32\cmd.exe /c C:\�Z�Z\sleep.bat"
End Sub

'---------------------------------------------------------------------------
'�蓮�ő��삷��
'---------------------------------------------------------------------------
Sub onplay()

    S001.run
    Debug.Print "S001_ok"
    
    S002.run
    Debug.Print "S002_ok"
    
    ThisWorkbook.Save
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    With CreateObject("WScript.Shell").popup("�蓮���삪���܂���", 20, "�I��", vbInformation): End With
End Sub

'---------------------------------------------------------------------------
'�S�ẴG�N�Z��Window�������I��
'---------------------------------------------------------------------------
Private Sub Excel_Task_Kill()
    Application.DisplayAlerts = False
    ThisWorkbook.Save
    Application.DisplayAlerts = True
    Sleep 3000

    Shell "taskkill /F /IM excel.exe"

End Sub

'---------------------------------------------------------------------------
'����̃^�X�N�����I���ꍇ
'---------------------------------------------------------------------------
Private Sub targetTaskKill()

    Dim wdApp   As New Word.Application
    Dim wdApp As Object: Set wdApp = CreateObject("Word.Application")
    Dim wdtasks As Object: Set wdtasks = wdApp.Tasks
    Dim i As Long

    For i = 1 To wdtasks.count
        DoEvents
        If wdtasks(i).Visible = True And InStr(wdtasks(i).Name, "�^�X�N��") > 0 Then
                wdtasks(i).Kill
        End If
    Next i

    Set wdtasks = Nothing                        '���
    wdApp.Quit                                   '������Ȃ��ƃ^�X�N���c���Ă��܂��I    Set wdApp = Nothing                          '���


End Sub

