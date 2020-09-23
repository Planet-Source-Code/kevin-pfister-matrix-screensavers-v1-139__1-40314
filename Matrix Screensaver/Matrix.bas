Attribute VB_Name = "ModGeneral"
Option Explicit

Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Sub Main()
    If App.PrevInstance = True Then
        End
    End If
    Select Case Left$(UCase$(Command$), 2)
        Case "/A"         'change password
            'no password support (yet)
        Case "/C"         'config
            FrmConfig.Show
        Case "/S"         'displaY screensaver
            FrmMain.Show
    End Select
End Sub

