Attribute VB_Name = "参照設定"
Option Explicit

Private Sub init()

    'ファイルシステムオブジェクト
    Dim FSO As Scripting.FileSystemObject: Set FSO = New Scripting.FileSystemObject
    'Dim FSO As Object: Set FSO = CreateObject("Scripting.FileSystemObject")
    
    'WSHシェルオブジェクト
    Dim wsh As IWshRuntimeLibrary.WshShell: Set wsh = New IWshRuntimeLibrary.WshShell
    'Dim WSH As Object: Set wsh = CreateObject("WScript.Shell")
    
    'シェルオブジェクト
    Dim sh As Shell32.Shell: Set sh = New Shell32.Shell
    'Dim SH As Object: Set SH = CreateObject("Shell.Application")
    
    '正規表現オブジェクト
    Dim re As VBScript_RegExp_55.RegExp: Set re = New VBScript_RegExp_55.RegExp
    'Dim RE As Object: Set RE = CreateObject("VBScript.RegExp")
    
    'WMIオブジェクト
    Dim WMI As WbemScripting.SWbemLocator: Set WMI = New WbemScripting.SWbemLocator
    'Dim WMI As Object: Set WMI = CreateObject("WbemScripting.SWbemLocator")
    
    'ADOコネクションオブジェクト
    Dim CN As ADODB.Connection: Set CN = New ADODB.Connection
    'Dim CN As Object: Set CN = CreateObject("ADODB.Connection")
    
    'ADOレコードセットオブジェクト
    Dim RS As ADODB.Recordset: Set RS = New ADODB.Recordset
    'Dim RS As Object: Set RS = CreateObject("ADODB.Recordset")
    
    'ADOストリームオブジェクト
    Dim adoStream As ADODB.Stream: Set adoStream = New ADODB.Stream
    'Dim adoStream As Object: Set adoStream = CreateObject("ADODB.Stream")
    
    'ディクショナリーオブジェクト
    Dim dic As Scripting.Dictionary: Set dic = New Scripting.Dictionary
    'Dim DIC As Object: Set DIC = CreateObject("Scripting.Dictionary")
    
    'データオブジェクト
    Dim clipboard As MSForms.DataObject: Set clipboard = New MSForms.DataObject
    'Dim clipboard As Object: Set clipboard = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    
    'InternetExplorerオブジェクト
    Dim IE As SHDocVw.InternetExplorer: Set IE = New SHDocVw.InternetExplorer
    'Dim IE As Object: Set IE = CreateObject("InternetExplorer.Application")
    
    '解放
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

