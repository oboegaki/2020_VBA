Attribute VB_Name = "Enavi"
Option Explicit


'�����I�ɍőO�ʂɂ�����
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
'�ŏ�������Ă��邩���ׂ�
Private Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long
'���̃T�C�Y�ɖ߂�
Private Declare Function ShowWindowAsync Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

'public�p�錾
Dim IE As Object

Sub run()
    'initialize
    Set IE = CreateObject("InternetExplorer.Application")
    
    '���O�C���y�[�W�J��
    Call IEStart("https://�Z�Z")
    Call login
    
End Sub

'----------------------------------------------------------
'STEP1 : IE�N��
'----------------------------------------------------------
Private Sub IEStart(ByVal url As String)
    IE.Visible = True
    IE.Navigate url

    '�ŏ�������Ă���ꍇ�͌��ɖ߂�(9=RESTORE:�ŏ����O�̏��)
    If IsIconic(IE.hWnd) Then
        ShowWindowAsync IE.hWnd, &H9
    End If
    SetForegroundWindow (IE.hWnd)     '�őO�ʂɕ\��
    
    Call loading
End Sub


'----------------------------------------------------------
'STEP2 : ���O�C��
'----------------------------------------------------------
Private Sub login()

    'ID�EPASS����
    IE.Document.getElementById("u").Value = "���[�U��"
    IE.Document.getElementById("p").Value = "�p�X���[�h"

    '4�b�ҋ@
    'WScript.Sleep 5000
    Application.Wait Now + TimeValue("0:00:04")
    
    '���O�C��
    IE.Document.getElementById("loginButton").Click
    Call loading
    
    'Enavi�g�b�v��
    Call linksClick

End Sub


'----------------------------------------------------------
'STEP3 : ���O�C����Ƀ����N���N���b�N�������y�[�W�ɑJ��
'----------------------------------------------------------
Private Sub linksClick()
    
    Call IEStart("https://�Z�Z")
    
    'IE�I�u�W�F�N�g���̃����N���`�F�b�N���A�N���b�N�����������N��T��
     Dim i As Long
     For i = 0 To IE.Document.links.Length - 1
       '�擾���������N��A�h���X�𕶎���P�ʂŔ�r
       If IE.Document.links(i).href = "javascript:void(0);" Then
          '�N���b�N�Ώۂ�������N���b�N���ă��[�v�𔲂���
           IE.Document.links(i).Click
           Exit For
       End If
    Next i
    
    Call loading
        
    Application.Wait Now + TimeValue("0:00:10")
    
    IE.Quit
    Set IE = Nothing
End Sub


Private Sub loading()
    '�T�C�g�̓ǂݍ��ݑ҂�
    Do While IE.Busy = True Or IE.ReadyState <> 4
        DoEvents
    Loop
    'JavaScript�ǂݍ��ݑ҂�
    Application.Wait Now + TimeValue("0:00:03")
End Sub


