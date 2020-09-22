VERSION 5.00
Begin VB.Form FRMscreen 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrClear 
      Left            =   600
      Top             =   2040
   End
   Begin VB.Timer tmrStart 
      Left            =   1200
      Top             =   960
   End
End
Attribute VB_Name = "FRMscreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'_____________________2000_______________________
'][-][-][-][-][-][-][-][-][-][-][-][-][-][-][-]['
'][-][-][-][-]One Computer Software[-][-][-][-]['
'][-][-][-][-][-Screen Saver Demo-][-][-][-][-]['
'][-][-][-][-][-]DeI3oe@aol.com [-][-][-][-][-]['
'][-][-][-][-][-][-][-][-][-][-][-][-][-][-][-]['

Option Explicit
Dim hWnd1 As Long
Dim MOUSEx, MOUSEy As Integer

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Const SWP_HIDEWINDOW = &H80
    Const SWP_SHOWWINDOW = &H40

Sub STARTit()
    Dim R, G, B As Integer 'holds colors
    Dim XPos, YPos As Integer 'holds the positions where dot will go
    
    R = 225 * Rnd 'sets red to random
    G = 225 * Rnd 'sets green to random
    B = 225 * Rnd 'sets blue to random
    
    XPos = Rnd * ScaleWidth 'sets the horizontal position
    YPos = Rnd * ScaleHeight 'sets the vertical position
    
    PSet (XPos, YPos), RGB(R, G, B) 'plots the points with random color
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call SetWindowPos(hWnd1, 0, 0, 0, 0, 0, SWP_SHOWWINDOW) 'shows taskbar
    End 'quits if anything on keyboard is hit
    
End Sub

Private Sub Form_Load()
    'makes the frm the same size as ur monitor
    Me.Width = Screen.Width
    Me.Height = Screen.Height
    
    'centers frm
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    
    'sets the timer speed
    tmrStart.Interval = 1
    
    'sets the timer speed that will clear the field
    tmrClear.Interval = 60000 '= 1 minute
    tmrClear.Tag = 0
    
    'sets up mouse stuff
    MOUSEx = 0
    MOUSEy = 0
    
    'hides taskbar
    hWnd1 = FindWindow("Shell_traywnd", "")
    Call SetWindowPos(hWnd1, 0, 0, 0, 0, 0, SWP_HIDEWINDOW)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'gets the mouse position
    If MOUSEx = 0 Then 'determines if already got mouse position
        MOUSEx = x 'gets the X position of mouse
        MOUSEy = y 'gets the Y position of mouse
    End If
    
    If x + y > MOUSEx + MOUSEy + 10 Then
        Call SetWindowPos(hWnd1, 0, 0, 0, 0, 0, SWP_SHOWWINDOW) 'shows taskbar
        End 'ends project if mouse is moved toward right
    End If
    If x - y < MOUSEx - MOUSEy - 10 Then
        Call SetWindowPos(hWnd1, 0, 0, 0, 0, 0, SWP_SHOWWINDOW) 'shows taskbar
        End 'ends project if mouse is moved toward left
    End If
End Sub

Private Sub tmrClear_Timer()
    If tmrClear.Tag = 2 Then 'means it will wait for 3 minutes to do this
        Me.Refresh 'clears the field
        tmrClear.Tag = 0 'resets this
        Exit Sub 'exits sub so it doesnt count a minute and will start from 0
    End If
    
    tmrClear.Tag = tmrClear.Tag + 1 'counts the minutes
End Sub

Private Sub tmrStart_Timer()
    STARTit 'starts the nice little star field!!!!!!!!!!!!!!
End Sub
