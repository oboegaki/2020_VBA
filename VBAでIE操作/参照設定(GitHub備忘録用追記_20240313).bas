Attribute VB_Name = "�Q�Ɛݒ�"
Option Explicit

Private Sub init()

    '�t�@�C���V�X�e���I�u�W�F�N�g
    Dim FSO As Scripting.FileSystemObject: Set FSO = New Scripting.FileSystemObject
    'Dim FSO As Object: Set FSO = CreateObject("Scripting.FileSystemObject")
    
    'WSH�V�F���I�u�W�F�N�g
    Dim wsh As IWshRuntimeLibrary.WshShell: Set wsh = New IWshRuntimeLibrary.WshShell
    'Dim WSH As Object: Set wsh = CreateObject("WScript.Shell")
    
    '�V�F���I�u�W�F�N�g
    Dim sh As Shell32.Shell: Set sh = New Shell32.Shell
    'Dim SH As Object: Set SH = CreateObject("Shell.Application")
    
    '���K�\���I�u�W�F�N�g
    Dim re As VBScript_RegExp_55.RegExp: Set re = New VBScript_RegExp_55.RegExp
    'Dim RE As Object: Set RE = CreateObject("VBScript.RegExp")
    
    'WMI�I�u�W�F�N�g
    Dim WMI As WbemScripting.SWbemLocator: Set WMI = New WbemScripting.SWbemLocator
    'Dim WMI As Object: Set WMI = CreateObject("WbemScripting.SWbemLocator")
    
    'ADO�R�l�N�V�����I�u�W�F�N�g
    Dim CN As ADODB.Connection: Set CN = New ADODB.Connection
    'Dim CN As Object: Set CN = CreateObject("ADODB.Connection")
    
    'ADO���R�[�h�Z�b�g�I�u�W�F�N�g
    Dim RS As ADODB.Recordset: Set RS = New ADODB.Recordset
    'Dim RS As Object: Set RS = CreateObject("ADODB.Recordset")
    
    'ADO�X�g���[���I�u�W�F�N�g
    Dim adoStream As ADODB.Stream: Set adoStream = New ADODB.Stream
    'Dim adoStream As Object: Set adoStream = CreateObject("ADODB.Stream")
    
    '�f�B�N�V���i���[�I�u�W�F�N�g
    Dim dic As Scripting.Dictionary: Set dic = New Scripting.Dictionary
    'Dim DIC As Object: Set DIC = CreateObject("Scripting.Dictionary")
    
    '�f�[�^�I�u�W�F�N�g
    Dim clipboard As MSForms.DataObject: Set clipboard = New MSForms.DataObject
    'Dim clipboard As Object: Set clipboard = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    
    'InternetExplorer�I�u�W�F�N�g
    Dim IE As SHDocVw.InternetExplorer: Set IE = New SHDocVw.InternetExplorer
    'Dim IE As Object: Set IE = CreateObject("InternetExplorer.Application")
    
    '���
    Set sh = Nothing
    Set re = Nothing
    Set WMI = Nothing
    Set CN = Nothing
    Set RS = Nothing
    Set adoStream = Nothing
    Set dic = Nothing
    Set clipboard = Nothing
    Set IE = Nothing

End Sub

