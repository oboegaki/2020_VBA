Attribute VB_Name = "SP_SENDKEY"
Option Explicit


'�}�E�X����(https://thom.hateblo.jp/entry/2015/11/21/002304)
Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Declare Sub mouse_event Lib "user32" ( _
    ByVal dwFlags As Long, _
    Optional ByVal dx As Long = 0, _
    Optional ByVal dy As Long = 0, _
    Optional ByVal dwDate As Long = 0, _
    Optional ByVal dwExtraInfo As Long = 0)

'Sub �}�E�X�ŉ�ʂ̔C�ӂ̈ʒu���N���b�N()
'    SetCursorPos 100, 35  '������100�s�N�Z���A�ォ��35�s�N�Z���̈ʒu�ɃJ�[�\�����ړ�
'    'mouse_event 2  '���{�^�������̃R�[�h
'    'mouse_event 4  '���{�^������̃R�[�h
'End Sub



'�����I�ɍőO�ʂɂ�����
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
'�ŏ�������Ă��邩���ׂ�
Private Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long
'���̃T�C�Y�ɖ߂�
Private Declare Function ShowWindowAsync Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

'public�p�錾
Dim IE As Object
Dim pointlinks
Dim clicklinks


'�N���b�N����
Sub clickPoint()
    mouse_event 2  '���{�^�������̃R�[�h
    mouse_event 4  '���{�^������̃R�[�h
End Sub


'HOME�{�^�����N���b�N
Sub homePos()
    SetCursorPos 952, 911: clickPoint: Sleep (2000)
End Sub
'�߂�{�^�����N���b�N
Sub backPos()
    SetCursorPos 839, 916  '������839�s�N�Z���A�ォ��916�s�N�Z���̈ʒu�ɃJ�[�\�����ړ�
    clickPoint
End Sub
'����t�H���_���N���b�N
Sub DFolderPos()
    Sleep (5000)
    SetCursorPos 945, 505: clickPoint: Sleep (6000)
End Sub
'�{�^���������đҋ@
Sub repeat(x As Long, y As Long)
    Dim maxSec As Long: maxSec = 0
    SetCursorPos x, y: clickPoint
    While (maxSec <> 180)
        DoEvents
        Sleep (1000)
        maxSec = maxSec + 1
    Wend
    maxSec = 0
End Sub


'---------------------------------
' STEP : 1
'---------------------------------
Sub DIVIDE()
    DFolderPos                                                  '�t�H���_���J��
    SetCursorPos 857, 395: clickPoint: Sleep (10000)       '�Q�[���N��
    Call repeat(782, 209): Call backPos: Sleep (5000)       '1�����A�ҋ@���Ė߂�
    Call repeat(782, 209): Call backPos: Sleep (5000)       '2�����A�ҋ@���Ė߂�
    Call repeat(782, 209): Call backPos: Sleep (5000)       '3�����A�ҋ@���Ė߂�
    Call repeat(782, 209): homePos                          '4�L�����A160�b�ҋ@����HOME(�I��)
End Sub


