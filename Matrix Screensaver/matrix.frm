VERSION 5.00
Begin VB.Form FrmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "The Matrix By Kevin Pfister"
   ClientHeight    =   9660
   ClientLeft      =   120
   ClientTop       =   690
   ClientWidth     =   10110
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00008000&
   Icon            =   "matrix.frx":0000
   ScaleHeight     =   40.25
   ScaleMode       =   4  'Character
   ScaleWidth      =   84.25
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer TmrFrameRate 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1080
      Tag             =   "Falling Code"
      Top             =   600
   End
   Begin VB.PictureBox PicMatrix 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   1560
      ScaleHeight     =   57
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   0
      Tag             =   "Falling Code"
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer TmrApply 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   600
      Tag             =   "Startup"
      Top             =   600
   End
   Begin VB.Timer TmrMain3 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   120
      Tag             =   "Knock"
      Top             =   600
   End
   Begin VB.Timer TmrMain2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1080
      Tag             =   "Tracing"
      Top             =   120
   End
   Begin VB.Timer TmrMain1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   600
      Tag             =   "Tracing"
      Top             =   120
   End
   Begin VB.Timer TmrMain 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Tag             =   "Falling Code"
      Top             =   120
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal IntX As Long, ByVal IntY As Long) As Long

'#######################################################################
'3 Matrix Screensavers made by Kevin Pfister
'#######################################################################

'READ FIRST

'The program is preset to be a screensaver, this means that it will not run in VB
'To make it work, in properties change the startup object form Sub Main to frmMain.
'If you stop the program from running in Vb ie. Ctrl+Break the cursor will be
'invisible


'General Variables
Dim IntLastXPos As Integer    'For use in Checking the mouse movements
Dim IntLastYPos As Integer 'For use in Checking the mouse movements
Dim IntFrames As Integer
Dim IntFrameRate As Integer
Dim IntActiveScreensaver As Integer

'Falling Code Variables
Dim IntBackGroundPic() As Integer
Dim IntLengthOfDrop() As Integer   'Length of Dropping column
Dim IntLeading() As Integer   'To hold the IntLeading letters
Dim IntLetter() As Integer   'The symbol
Dim IntColour() As Integer    'The IntColour of the symbol
Dim IntLngWaitLngBeforeClear() As Integer        'To hold the length of time LngBefore the symbol fades
Dim IntMaxLength As Integer   'The maximum length of the column
Dim IntMaxLngWait As Integer   'The maximum Waiting time Before clearing
Dim IntDropCols As Integer   'The StrNumber of dropping coloumns
Dim IntFadeSpeed As Integer   'The fading speed of the symbols
Dim IntFromTop As Integer   'If the column starts falling from the top or from a random position
Dim IntWillFade As Integer   'Will the letter fade or not
Dim IntMultipleColours As Integer   'Is it single or multiple Colours
Dim IntFntsize As Integer
Dim LngOneCol As Long
Dim BlnUseBackGround As Boolean
Dim StrImageFile As String
Dim IntCodeColour As Integer

'Tracing Variables
Dim IntYNums(1 To 30) As Integer
Dim IntXNums(1 To 60) As Integer
Dim IntTextDone As Integer   'How much has been drawn to the screen already
Dim IntSTextF As Integer
Dim StrPhoneNo(1 To 11) As String   'The seperate parts of the phone StrNumber
Dim IntAnim As Integer    'Change draw IntColour (1 -> 0 -> 1 -> 0...)
Dim LngXSpace As Long  'Where the StrNumbers are to be drawn
Dim LngYSpace  As Long 'Where the StrNumbers are to be drawn
Dim LngRanNum As Long  'If random StrNumber was choosen
Dim LngTraceCol As Long
Dim LngYCoord As Long
Dim LngXCoord(1 To 11) As Long
Dim LngWait As Long
Dim BlnCols(60) As Boolean    'The different columns, when clearing
Dim BlnPhoneOn(1 To 11) As Boolean 'To IntCheck if the phone StrNumber is to be shown
Dim StrNumber As String    'The phone StrNumber to be traced
Dim StrNumbers(60, 30) As Integer 'all the StrNumbers
Dim StrStartText As String   'Text to be drawn to the screen

'Knock, Knock Variables
Dim IntTxtSpeed(4) As Integer
Dim IntMatrixDone As Integer
Dim IntCurrentTxt As Integer
Dim StrTxtMatrix(4) As String

Private Sub Form_Click()
    Call ExitProgram
End Sub

Private Sub Form_DblClick()
    Call ExitProgram
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Call ExitProgram
End Sub

'#######################################################################
'General Section
'#######################################################################

Private Sub Form_Load()
    Dim IntX As Integer
    Dim IntY As Integer
    Dim IntCurrent As Integer
    Dim IntDoFill As Integer
    Dim IntPNo As Integer
    'ShowCursor (0)  'Make the cursor invisible
    FrmMain.WindowState = 2
    ForeColor = RGB(0, 220, 0)  'Change the forecolor to the default shade of green
    
    '#######################################################################
    'General Settings
    '#######################################################################
    
    IntActiveScreensaver = Val(GetSetting("Kevin Pfister's Matrix", "Options", "Which", 0))
    
    '#######################################################################
    'Falling Code Settings
    '#######################################################################
    
    IntMaxLength = GetSetting("Kevin Pfister's Matrix", "Drops", "MaxDrop", 100)   'Retieve the Maximum length of the columns
    IntMaxLngWait = GetSetting("Kevin Pfister's Matrix", "Drops", "BeforeClean", 100)   'Retieve the maximum LngWaiting time LngBefore clearing the symbol
    IntDropCols = GetSetting("Kevin Pfister's Matrix", "Drops", "DropsRunning", 20)       'Retieve the StrNumber of dropping columns
    IntFadeSpeed = GetSetting("Kevin Pfister's Matrix", "Drops", "FadeSpeed", 2)          'Retieve the fading speed of the columns
    IntFromTop = GetSetting("Kevin Pfister's Matrix", "Options", "FromTop", 1)   'Retieve if the columns start from the top or from a random position
    IntFntsize = Val(GetSetting("Kevin Pfister's Matrix", "Options", "Size", "12"))   'Retieve font size
    IntWillFade = GetSetting("Kevin Pfister's Matrix", "Colour", "Fade", 0)       'Retieve if the symbols fade or not
    IntMultipleColours = GetSetting("Kevin Pfister's Matrix", "Colour", "MColours", 1)   'Retieve if it are different shades of green
    TmrMain.Interval = 1000 / GetSetting("Kevin Pfister's Matrix", "Options", "Frame Rate", 100)
    IntCodeColour = GetSetting("Kevin Pfister's Matrix", "Options", "Colour", 1)
    If Val(GetSetting("Kevin Pfister's Matrix", "Options", "UseImage", 1)) = 0 Then
        BlnUseBackGround = False
    Else
        BlnUseBackGround = True
    End If
    StrImageFile = GetSetting("Kevin Pfister's Matrix", "Options", "BckImage", "C:\Agent.jpg")
    
    '#######################################################################
    'Tracing Settings
    '#######################################################################
    
    StrNumber = GetSetting("Kevin Pfister's Matrix", "Options", "StrNumber", "0000000000")
    LngRanNum = GetSetting("Kevin Pfister's Matrix", "Options", "Random", 1)
    LngTraceCol = RGB(0, 220, 0)
    LngXSpace = Width / 45
    LngYSpace = Height / 35
    LngYCoord = LngYSpace * 3
    For IntX = 1 To 11
        LngXCoord(IntX) = LngXSpace * (2 + IntX)
    Next
    For IntX = 1 To 60
        IntXNums(IntX) = LngXSpace * (2 + IntX)
    Next
    For IntY = 1 To 30
        IntYNums(IntY) = LngYSpace * (4 + IntY)
    Next
    '#######################################################################
    'Knock Knock Neo Settings
    '#######################################################################
    
    StrTxtMatrix(1) = "Wake up,  Neo. . ."
    IntTxtSpeed(1) = 150
    StrTxtMatrix(2) = "The Matrix has you. . ."
    IntTxtSpeed(2) = 150
    StrTxtMatrix(3) = "Follow the white rabbit."
    IntTxtSpeed(3) = 150
    StrTxtMatrix(4) = "Knock,  Knock,  Neo.."
    IntTxtSpeed(4) = 1
    IntCurrent = 1
    
    'This sets the resolution of the screensaver
    'The lower the resolution of the screen, the faster the screensaver will be
    'because it reduces the loop sizes
    'OnlY use higher resolutions with faster computer
    Randomize Timer 'randomize the screensaver
    
    If IntActiveScreensaver = 0 Then  'Falling Code
        TmrApply.Enabled = True
    ElseIf IntActiveScreensaver = 1 Then 'Tracing
        For IntDoFill = 1 To 60
            BlnCols(IntDoFill) = 1
        Next
        For IntPNo = 1 To 11
            If LngRanNum = 1 Then
                StrPhoneNo(IntPNo) = Int(Rnd * 9)
            Else
                StrPhoneNo(IntPNo) = Mid(StrNumber, IntPNo, 1)
            End If
        Next
        StrStartText = "Call Trans opt: Rec " + Str$(Date) + " " + Str$(Time) + " Rec:Log> "
        ForeColor = RGB(0, 220, 0)
        ScaleMode = 1
        
        Font = "MS Serif"
        TmrMain1.Enabled = True
    ElseIf IntActiveScreensaver = 2 Then 'Knock,Knock
        
        Font = "Arial"
        ForeColor = &H9BAC9B
        TmrMain3.Enabled = True
        IntCurrentTxt = 1
    End If
End Sub

'#######################################################################
'Falling Code Section
'#######################################################################

'Two different subroutines are used to maximise performance

Sub OneIntColour() 'The routine for drawing one IntColour
    Dim IntX As Integer
    Dim IntY As Integer
    Dim IntDrops As Integer
    For IntX = 1 To FrmMain.ScaleWidth
        For IntY = 1 To FrmMain.ScaleHeight + 5
            If IntLetter(IntX, IntY) <> 0 Then
                If IntLeading(IntX, IntY) = 1 Then 'Is it IntLeading
                    If IntY <= FrmMain.ScaleHeight + 4 Then 'Is it smaller than the screen height                    If IntLengthOfDrop(IntX,IntY) > 0 Then 'Is there still IntDrops in this column
                        If IntLengthOfDrop(IntX, IntY) > 0 Then 'Is there still IntDrops in this column
                            If IntLetter(IntX, IntY + 1) > 0 Then
                                Call Clear(IntX, IntY + 1)
                            End If
                            IntLengthOfDrop(IntX, IntY + 1) = IntLengthOfDrop(IntX, IntY) - 1
                            IntLeading(IntX, IntY + 1) = 2
                            IntLetter(IntX, IntY + 1) = Int(Rnd * 43) + 65
                            IntLeading(IntX, IntY) = 0
                            IntLngWaitLngBeforeClear(IntX, IntY) = IntMaxLngWait
                            FrmMain.CurrentX = IntX
                            FrmMain.CurrentY = IntY - 4
                            FrmMain.ForeColor = LngOneCol
                            FrmMain.Print Chr(IntLetter(IntX, IntY))
                            Call ShowHigh(IntX, IntY + 1)
                        Else    'End of Drop(Kill Letter/Symbol)
                            If IntLeading(IntX, IntY) = 1 Then
                                Call Clear(IntX, IntY)
                            End If
                        End If
                    Else
                        IntLeading(IntX, IntY) = 0
                        Call Clear(IntX, IntY)
                    End If
                End If
                If IntLeading(IntX, IntY) = 1 Or IntLeading(IntX, IntY) = 2 Then
                    IntLeading(IntX, IntY) = 1
                    IntDrops = IntDrops + 1
                End If
                If IntLngWaitLngBeforeClear(IntX, IntY) > 0 Then 'Is the Letter/Symbol dieing?
                    IntLngWaitLngBeforeClear(IntX, IntY) = IntLngWaitLngBeforeClear(IntX, IntY) - 1
                    If IntLngWaitLngBeforeClear(IntX, IntY) = 0 Then
                        Call Clear(IntX, IntY)
                    End If
                End If
            End If
        Next
    Next
    If IntDrops < IntDropCols Then
        Dim IntMakeNew As Integer
        For IntMakeNew = IntDrops To IntDropCols
            IntX = Int(Rnd * FrmMain.ScaleWidth) + 1
            If IntFromTop = 1 Then
                IntY = Int(Rnd * 5) + 1
            Else
                IntY = Int(Rnd * FrmMain.ScaleHeight) + 1
            End If
            If IntLetter(IntX, IntY) > 0 Then
                Call Clear(IntX, IntY)
            End If
            IntLengthOfDrop(IntX, IntY) = Int(Rnd * IntMaxLength)
            IntLeading(IntX, IntY) = 1
            IntLetter(IntX, IntY) = 64 + Int(Rnd * 26)
            Call ShowHigh(IntX, IntY)
        Next
    End If
End Sub

Sub MoreThanOneIntColour()
    Dim IntX As Integer
    Dim IntY As Integer
    Dim IntDrops As Integer
    Dim IntMakeNew As Integer
    For IntX = 1 To FrmMain.ScaleWidth
        For IntY = 1 To FrmMain.ScaleHeight + 5
            If IntLetter(IntX, IntY) <> 0 Then
                If IntLeading(IntX, IntY) = 1 Then 'Is it IntLeading
                    If IntY <= FrmMain.ScaleHeight + 4 Then 'Is it smaller than the screen height
                        If IntLengthOfDrop(IntX, IntY) > 0 Then 'Is there still IntDrops in this column
                            If IntLetter(IntX, IntY + 1) <> 0 Then
                                Call Clear(IntX, IntY + 1)
                            End If
                            IntLengthOfDrop(IntX, IntY + 1) = IntLengthOfDrop(IntX, IntY) - 1
                            IntLeading(IntX, IntY) = 0
                            IntLeading(IntX, IntY + 1) = 2
                            IntLetter(IntX, IntY + 1) = Int(Rnd * 43) + 65
                            If BlnUseBackGround = True Then
                                IntColour(IntX, IntY + 1) = IntBackGroundPic(IntX, IntY + 1) + Rnd * 40 'Set the IntColour
                            Else
                                IntColour(IntX, IntY + 1) = Rnd * 100 + 100
                            End If
                            IntLngWaitLngBeforeClear(IntX, IntY) = IntMaxLngWait
                            Call ShowColor(IntX, IntY)
                            Call ShowHigh(IntX, IntY + 1)
                        Else    'End of Drop(Kill Letter/Symbol)
                            If IntLeading(IntX, IntY) = 1 Then
                                Call Clear(IntX, IntY)
                            End If
                        End If
                    Else
                        IntLeading(IntX, IntY) = 0
                        Call Clear(IntX, IntY)
                    End If
                ElseIf IntLeading(IntX, IntY) = 2 Then
                    IntLeading(IntX, IntY) = 1
                    IntDrops = IntDrops + 1
                End If
                If IntLngWaitLngBeforeClear(IntX, IntY) > 0 Then 'Is the Letter/Symbol dieing?
                    IntLngWaitLngBeforeClear(IntX, IntY) = IntLngWaitLngBeforeClear(IntX, IntY) - 1
                    If IntLngWaitLngBeforeClear(IntX, IntY) = 0 Or IntColour(IntX, IntY) = 0 Then
                        Call Clear(IntX, IntY)
                    ElseIf IntWillFade = 1 Then   'Is fading ativated
                        IntColour(IntX, IntY) = IntColour(IntX, IntY) - IntFadeSpeed
                        If IntColour(IntX, IntY) < 0 Then
                            IntColour(IntX, IntY) = 0
                            Call Clear(IntX, IntY)
                        ElseIf IntLeading(IntX, IntY) = 0 Then
                            Call ShowColor(IntX, IntY)
                        End If
                    End If
                End If
            End If
        Next
    Next
    If IntDrops < IntDropCols Then
        For IntMakeNew = IntDrops To IntDropCols
            IntX = Int(Rnd * FrmMain.ScaleWidth) + 1
            If IntFromTop = 1 Then
                IntY = Int(Rnd * 5) + 1
            Else
                IntY = Int(Rnd * FrmMain.ScaleHeight) + 1
            End If
            If IntLetter(IntX, IntY) > 0 Then
                Call Clear(IntX, IntY)
            End If
            IntLengthOfDrop(IntX, IntY) = Int(Rnd * IntMaxLength)
            IntLeading(IntX, IntY) = 1
            IntLetter(IntX, IntY) = 64 + Int(Rnd * 26)
            If BlnUseBackGround = True Then
                IntColour(IntX, IntY) = IntBackGroundPic(IntX, IntY + 1) + Rnd * 40 'Set the IntColour
            Else
                IntColour(IntX, IntY) = Rnd * 100 + 100
            End If
            Call ShowHigh(IntX, IntY)
        Next
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If IntLastXPos = 0 And IntLastYPos = 0 Then
        IntLastXPos = X
        IntLastYPos = Y
    End If
    If Abs(X - IntLastXPos) > 20 Or Abs(Y - IntLastYPos) > 20 Then
        Call ExitProgram
    Else
        IntLastXPos = X
        IntLastYPos = Y
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call ExitProgram
End Sub

Private Sub Form_Terminate()
    Call ExitProgram
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call ExitProgram
End Sub

Private Sub TmrApply_Timer()
    Dim DoRand As Integer
    Dim XR As Integer
    Dim YR As Integer
    Dim temp As Long
    Dim IntX As Integer
    Dim IntY As Integer
    Dim Loading As Integer
    Dim AddNum As Integer
    TmrApply.Enabled = False
    
    'Change the font
    Font = "Matrix"   'Use the Matrix Font
    Font.Bold = False
    Font.Size = IntFntsize
    If IntCodeColour = 0 Then
        LngOneCol = RGB(150, 0, 0)
    ElseIf IntCodeColour = 1 Then
        LngOneCol = RGB(0, 150, 0)
    ElseIf IntCodeColour = 2 Then
        LngOneCol = RGB(0, 0, 150)
    End If
    'Change the variable sizes to fit the screen size
    ReDim IntLengthOfDrop(FrmMain.ScaleWidth, FrmMain.ScaleHeight + 5) As Integer
    ReDim IntLeading(FrmMain.ScaleWidth, FrmMain.ScaleHeight + 5) As Integer
    ReDim IntLetter(FrmMain.ScaleWidth, FrmMain.ScaleHeight + 5) As Integer
    ReDim IntColour(FrmMain.ScaleWidth, FrmMain.ScaleHeight + 5) As Integer
    ReDim IntLngWaitLngBeforeClear(FrmMain.ScaleWidth, FrmMain.ScaleHeight + 5) As Integer
    ReDim IntBackGroundPic(FrmMain.ScaleWidth, FrmMain.ScaleHeight + 5) As Integer
    
    If Dir(StrImageFile) = "" Then 'IntCheck to see if file exists
        BlnUseBackGround = False
    End If
    If BlnUseBackGround = True And IntMultipleColours = 1 Then
        PicMatrix.Picture = LoadPicture(StrImageFile)
        If BlnUseBackGround = True Then
            AddNum = 1
            Dim R1 As Integer
            Dim G1 As Integer
            Dim B1 As Integer
            For IntX = 1 To FrmMain.ScaleWidth
                Loading = Loading + AddNum
                If Loading = 15 Or Loading = 0 Then
                    AddNum = -AddNum
                End If
                Font = "Arial"   'Use the Matrix Font
                Font.Bold = False
                Font.Size = 12
                FrmMain.Cls
                FrmMain.CurrentX = FrmMain.ScaleWidth / 2 - 15
                FrmMain.CurrentY = FrmMain.ScaleHeight / 2
                FrmMain.Print "Processing Image " & String(Loading, ".")
                Font = "Matrix"   'Use the Matrix Font
                Font.Bold = False
                Font.Size = IntFntsize
                DoEvents
                For IntY = 1 To FrmMain.ScaleHeight
                    temp = GetPixel(PicMatrix.hDC, Int(PicMatrix.ScaleWidth / FrmMain.ScaleWidth * IntX), Int(PicMatrix.ScaleHeight / FrmMain.ScaleHeight * IntY))
                    GetRgb temp, R1, G1, B1
                    temp = Int((R1 + G1 + B1) / 3)
                    IntBackGroundPic(IntX, IntY + 4) = temp
                    IntBackGroundPic(IntX, IntY + 4) = (IntBackGroundPic(IntX, IntY + 4) + 1) / 100 * 80 + 20
                Next
            Next
        End If
        FrmMain.Cls
        PicMatrix.Picture = LoadPicture("") 'Free up memory
    End If
    For DoRand = 1 To IntDropCols 'Create the starting IntDrops
        XR = Int(Rnd * FrmMain.ScaleWidth) + 1  'The IntX position
        YR = Int(Rnd * (FrmMain.ScaleHeight + 5)) + 1   'The IntY position
        IntLengthOfDrop(XR, YR) = Int(Rnd * IntMaxLength)      'The Length of the drop
        IntLeading(XR, YR) = 1 'Make it a IntLeading symbol
        IntLetter(XR, YR) = Int(Rnd * 43) + 65         'Set the letter/symbol
        If IntMultipleColours = 1 Then                 'If multiple IntColours are enabled
            If BlnUseBackGround = True Then
                IntColour(XR, YR) = IntBackGroundPic(XR, YR) + Rnd * 40  'Set the IntColour
            Else
                IntColour(XR, YR) = Rnd * 100 + 100
            End If
        End If
    Next
    TmrMain.Enabled = True
    TmrFrameRate.Enabled = True
End Sub

Sub GetRgb(ByVal Color As Long, ByRef red As Integer, ByRef green As Integer, ByRef blue As Integer)
    Dim temp As Long
    temp = (Color And 255)
    red = temp And 255
    temp = Int(Color / 256)
    green = temp And 255
    temp = Int(Color / 65536)
    blue = temp And 255
End Sub

Private Sub TmrFrameRate_Timer()
    If IntFrameRate = 0 Then
        IntFrameRate = IntFrames
    Else
        IntFrameRate = (IntFrameRate + IntFrames) / 2
    End If
    IntFrames = 0
End Sub


Private Sub TmrMain_Timer()
    FrmMain.WindowState = 2
    If IntMultipleColours = 0 Then
        Call OneIntColour
    Else
        Call MoreThanOneIntColour
    End If
    IntFrames = IntFrames + 1
End Sub

Sub Clear(IntX, IntY) 'Clears a letter bIntY redrawing it as black
    FrmMain.ForeColor = vbBlack
    FrmMain.CurrentX = IntX
    FrmMain.CurrentY = IntY - 4
    FrmMain.Print Chr(IntLetter(IntX, IntY))
    IntLeading(IntX, IntY) = 0
    IntLngWaitLngBeforeClear(IntX, IntY) = 0
    IntLetter(IntX, IntY) = 0
    IntColour(IntX, IntY) = 0
End Sub

Sub ShowHigh(IntX, IntY) 'Shows a highlighted letter
    FrmMain.ForeColor = vbWhite
    FrmMain.CurrentX = IntX
    FrmMain.CurrentY = IntY - 4
    FrmMain.Print Chr(IntLetter(IntX, IntY))
End Sub

Sub ShowColor(IntX, IntY) 'Shows a IntColoured letter
    If IntCodeColour = 0 Then
        FrmMain.ForeColor = RGB(IntColour(IntX, IntY), 0, 0)
    ElseIf IntCodeColour = 1 Then
        FrmMain.ForeColor = RGB(0, IntColour(IntX, IntY), 0)
    ElseIf IntCodeColour = 2 Then
        FrmMain.ForeColor = RGB(0, 0, IntColour(IntX, IntY))
    End If
    FrmMain.CurrentX = IntX
    FrmMain.CurrentY = IntY - 4
    FrmMain.Print Chr(IntLetter(IntX, IntY))
End Sub

Sub ShowBlack(IntX, IntY) 'Shows a IntColoured letter
    FrmMain.ForeColor = vbBlack
    FrmMain.CurrentX = IntX
    FrmMain.CurrentY = IntY - 4
    FrmMain.Print Chr(IntLetter(IntX, IntY))
End Sub

'#######################################################################
'Call Tracing Section
'#######################################################################

Sub NewNewStrNumbers()
    Dim IntNewCol As Integer
    Dim IntVerts As Integer
    'Fills the Grid with random StrNumbers
    For IntNewCol = 1 To 60
        BlnCols(IntNewCol) = True
        For IntVerts = 1 To 30
            StrNumbers(IntNewCol, IntVerts) = Int(Rnd * 10)
        Next
    Next
End Sub

Private Sub TmrMain1_Timer()
    IntTextDone = IntTextDone + 1
    FrmMain.Cls
    FrmMain.CurrentX = 500
    FrmMain.CurrentY = 500
    IntAnim = 1 - IntAnim
    FrmMain.Print Mid$(StrStartText, 1, IntTextDone);
    FrmMain.ForeColor = RGB(0, 100 + (IntAnim * 150), 0)
    FrmMain.ForeColor = LngTraceCol
    If IntTextDone = Len(StrStartText) Then
        StrStartText = "Trace Program Running"
        IntTextDone = 0
        If IntSTextF = 1 Then
            TmrMain1.Enabled = False
            TmrMain2.Enabled = True
        End If
        IntSTextF = 1
        LngWaitFor (1)
    End If
    Call NewNewStrNumbers
End Sub

Private Sub TmrMain2_Timer()
    Dim IntDoPhone As Integer
    Dim IntNoPhone As Integer
    Dim IntDoClear As Integer
    Dim IntComplete As Integer
    Dim IntCheck As Integer
    Dim IntDoHor As Integer
    Dim IntDoVert As Integer
    Dim BlnExitMe As Boolean
    FrmMain.Cls
    For IntDoPhone = 1 To 11
        If BlnPhoneOn(IntDoPhone) = True Then
            CurrentX = LngXCoord(IntDoPhone)
            CurrentY = LngYCoord
            Print StrPhoneNo(IntDoPhone)
        End If
    Next
    LngWait = LngWait + 1
    If LngWait = 20 Then
        LngWait = 0
        BlnExitMe = False
        Do
            IntNoPhone = Int(Rnd * 11) + 1
            If BlnPhoneOn(IntNoPhone) = False Then
                BlnExitMe = True
                BlnPhoneOn(IntNoPhone) = True
                For IntDoClear = IntNoPhone To 60 Step 10
                    BlnCols(IntDoClear) = False
                Next
            End If
            IntComplete = 0
            For IntCheck = 1 To 11
                If BlnPhoneOn(IntCheck) = True Then
                    IntComplete = IntComplete + 1
                End If
            Next
            If IntComplete = 11 Then
                TmrMain2.Enabled = False
                Call Finish
            End If
        Loop Until BlnExitMe = True
    End If
    For IntDoHor = 1 To 60
        If BlnCols(IntDoHor) = True Then
            For IntDoVert = 30 To 1 Step -1
                FrmMain.CurrentX = IntXNums(IntDoHor)
                FrmMain.CurrentY = IntYNums(IntDoVert)
                FrmMain.ForeColor = RGB(0, 150 + Rnd * 100, 0)
                FrmMain.Print StrNumbers(IntDoHor, IntDoVert)
                StrNumbers(IntDoHor, IntDoVert) = StrNumbers(IntDoHor, IntDoVert - 1)
            Next
        End If
        StrNumbers(IntDoHor, 1) = Int(Rnd * 10)
    Next
    FrmMain.ForeColor = LngTraceCol
End Sub

Sub Finish()
    Dim IntDoPhone As Integer
    For IntDoPhone = 1 To 11
        If BlnPhoneOn(IntDoPhone) = True Then
            CurrentX = LngXCoord(IntDoPhone)
            CurrentY = LngYCoord
            Print StrPhoneNo(IntDoPhone)
        End If
    Next
    CurrentX = 500
    CurrentY = 500
    Print "Trace Program: Completed "
    Call ClearUp
End Sub

Sub ClearUp()
    Dim IntPNo As Integer
    Call NewNewStrNumbers
    For IntPNo = 1 To 11
        BlnPhoneOn(IntPNo) = False
        If LngRanNum = 1 Then
            StrPhoneNo(IntPNo) = Int(Rnd * 9)
        Else
            StrPhoneNo(IntPNo) = Mid$(StrNumber, IntPNo, 1)
        End If
    Next
    LngWaitFor (30)
    StrStartText = "Call Trans opt: Rec " + Str$(Date) + " " + Str$(Time) + " Rec:Log> "
    IntTextDone = 0
    IntSTextF = 0
    TmrMain1.Enabled = True
End Sub

'#######################################################################
'Knock,Knock Neo... Section
'#######################################################################

Private Sub Tmrmain3_Timer()
    IntMatrixDone = IntMatrixDone + 1
    Cls
    CurrentY = 3
    CurrentX = 6
    Print Mid$(StrTxtMatrix(IntCurrentTxt), 1, IntMatrixDone);
    If IntMatrixDone = Len(StrTxtMatrix(IntCurrentTxt)) Then
        IntMatrixDone = 0
        IntCurrentTxt = IntCurrentTxt + 1
        If IntCurrentTxt = 5 Then
            TmrMain3.Enabled = False
            Call Doneall
        End If
        LngWaitFor (5)
        TmrMain3.Interval = IntTxtSpeed(IntCurrentTxt)
    End If
End Sub

Sub Doneall()
    TmrMain3.Enabled = False
    LngWaitFor (30)
    TmrMain3.Enabled = True
    IntCurrentTxt = 1
    IntMatrixDone = 0
    TmrMain3.Interval = IntTxtSpeed(IntCurrentTxt)
End Sub

Sub LngWaitFor(Interval)
    Dim LngBefore As Long
    LngBefore = Timer
    Do
        DoEvents
    Loop Until Timer - LngBefore > Interval
End Sub

Sub ExitProgram()
    ShowCursor (1)  'Make the cursor visible
    If IntActiveScreensaver = 0 Then
        SaveSetting "Kevin Pfister's Matrix", "Speed", "FrameRate", IntFrameRate
    End If
    End
End Sub
