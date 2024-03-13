Attribute VB_Name = "S000_GoogleTrend"
Option Explicit

Dim col As Collection
Sub run()
    'initialize
    Set col = New Collection
    lib.IE_ImageOff
    
    Call getTrend     'Google�g�����h���[�h�擾
End Sub

'----------------------------------------------------------
' STEP1 : Google�g�����h���[�h�擾
'----------------------------------------------------------
Private Sub getTrend()

    'Application.StatusBar = "�g�����h���[�h�N����..."
    lib.IE_OPEN "https://trends.google.co.jp/trends/trendingsearches/daily?geo=JP", IE
    
    'HTML�G�������g�Z�b�g
    Dim colTrends As IHTMLElementCollection
    Set colTrends = IE.Document.getElementsByClassName("md-list-block")

    Dim i As Long
    For i = 0 To colTrends.Length - 1
        'regex�̃Z�b�g�E�ݒ�
        Dim regex: Set regex = CreateObject("VBScript.RegExp")
        regex.Pattern = "(\S+)"
        regex.Global = True
        regex.MultiLine = True
        
        'innerText����g�����h���[�h�𒊏o�E�~��
        Dim result: Set result = regex.Execute(colTrends(i).innerText)
        If result.count > 2 Then
            col.Add result(1).SubMatches(0)
        End If
        Set regex = Nothing
    Next
    
    '�g�����h���[�h���E���Ȃ����web������
    If col.count < 1 Then
        IE.Quit
        Sleep (10000): Set IE = Nothing: lib.AllcloseIEWindow
        Exit Sub
    End If
    
    '�Œ蕶����collection�ɒǉ�
        With col
        .Add "�V���j���[�X"
        .Add "�����̓V�C"
        .Add "��ʏ��"
        .Add "�{���̓��t"
    End With
    
    '�i�[�������̂��g�����h���[�h�V�[�g��
    Dim idx As Long
    'ThisWorkbook.Sheets("�g�����h���[�h").Cells.Clear
    For idx = 1 To col.count
        ThisWorkbook.Sheets("�g�����h���[�h").Cells(idx, 1) = col.Item(idx)
    Next idx
    
    '���\�[�X���
    Set col = Nothing
    IE.Quit
    Sleep (10000): Set IE = Nothing: lib.AllcloseIEWindow
End Sub
 

