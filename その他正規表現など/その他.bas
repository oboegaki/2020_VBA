Attribute VB_Name = "Module1"
Option Explicit


'*******************************************************************
'*  �@ numExtract
'*          �����̂ݒ��o
'*  �����F  StringValue:�����̂ݒ��o�������Ώە�����(String)
'******************************************************************
Function numExtract(StringValue As String) As String
 
  '�ϐ��̏���
  Dim i As Integer
  Dim numText As String
 
  'Len �ŕ������������� �� 1 to ������i�ŏI�����j�܂� For Next�ŉe��������B
  For i = 1 To Len(StringValue)
     'Mid�ŕ�����������珇��numText�Ɋi�[
      numText = Mid(StringValue, i, 1)
    '����numText[0-9]�ɊY������ꍇ �ϐ�numExtract�Ɋi�[���Ă���
      If numText Like "[0-9]" Then: numExtract = numExtract & numText
  Next i
 
End Function


'**********************************
'* �����ꊇ�A��(�I��͈͂̕�������ɘA��)
'**********************************
Function concatCells(targetRng As Range) As String
    '�K���Ȕ͈͎w�肵�Ĥ������A��
    concat_cellss = WorksheetFunction.Concat(targetRng)
End Function




'******************************
'* ���K�\���ɂ��}�b�`���O�����Z���̒l���擾
'******************************
Function cellRegValue(pattern As String) As String
     Dim rng As Range
     For Each rng In Range("A1:G15")
            '���K�\���ɂ��}�b�`���O
            With CreateObject("VBScript.RegExp")
                .Global = True  '�S��������
                .pattern = "\d\d\d\d\d\d\d"
                
                
                If .test(rng) Then
                    cellRegValue = rng.Value
                    Exit For
                End If
            End With
     Next

End Function

'******************************
'* VBS�t�@�C���ŁA�ʃv���Z�X�Ń��b�Z�[�W�{�b�N�X��\�����������ꍇ
'******************************

Sub excute()

    '�ϐ��錾
    Dim filePath As String
    Dim fileNo As Integer
    Dim msg As Integer
    Dim re As Integer
    
    '�쐬����t�@�C���p�X���w��
    filePath = "C:\�Z�Z\msgbox.vbs"

    'vbs�t�@�C���N��
    With CreateObject("Wscript.Shell")
        re = .Run(Command:=filePath, WaitOnReturn:=True)
    End With
    
    'Yes�Ȃ珈���J�n
    If re = vbYes Then
        '�������͊J�n
        With CreateObject("Wscript.Shell")
             Call Application.Wait(Now + TimeValue("00:00:03")) '// 1�b��~
             
             '�������e
        End With
    End If
End Sub












