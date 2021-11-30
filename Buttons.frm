VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5310
   ClientLeft      =   2100
   ClientTop       =   1560
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   6615
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   1440
      Width           =   3135
   End
   Begin VB.Image OverImage 
      Height          =   555
      Left            =   2760
      Picture         =   "Buttons.frx":0000
      Top             =   3840
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Image UpImage 
      Height          =   555
      Left            =   2760
      Picture         =   "Buttons.frx":162E
      Top             =   3120
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Image DownImage 
      Height          =   555
      Left            =   2760
      Picture         =   "Buttons.frx":2C5C
      Top             =   2400
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Image ButtonPicture1 
      Height          =   555
      Index           =   3
      Left            =   240
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Image ButtonPicture1 
      Height          =   555
      Index           =   2
      Left            =   240
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Image ButtonPicture1 
      Height          =   555
      Index           =   1
      Left            =   240
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Image ButtonPicture1 
      Height          =   555
      Index           =   0
      Left            =   240
      Top             =   3960
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" _
  (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, _
   ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
   
Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" _
                                            (ByVal dwError As Long, _
                                             ByVal lpstrBuffer As String, _
                                             ByVal uLength As Long) As Long

Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" _
    (ByVal lpszLongPath As String, ByVal lpszShortPath As String, _
    ByVal cchBuffer As Long) As Long

'This is for the button rollovers
Dim MouseOver
Dim MousePress
Dim NewIndex

'This is for playing the wave files
Dim MouseOverSound As String
Dim MousePressSound As String
Dim MouseUpSound As String

Const MouseOverMCI As String = "WAVEOVER"
Const MousePressMCI As String = "WAVEPRESS"
Const MouseUpMCI As String = "WAVEUP11"


Private Sub ButtonPicture1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If MousePress Then Exit Sub
   StopSounds
   ButtonPicture1(Index).Picture = DownImage.Picture
   lblStatus.Caption = "Mouse Down"
   PlayWav MousePressMCI
   MousePress = True
End Sub

Private Sub ButtonPicture1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If MouseOver Then Exit Sub
   StopSounds
   ButtonPicture1(Index).Picture = OverImage.Picture
   lblStatus.Caption = "Mouse Over - Button"
   PlayWav MouseOverMCI
   NewIndex = Index
   MouseOver = True
End Sub

Private Sub ButtonPicture1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not MousePress Then Exit Sub
       StopSounds
       PlayWav MouseUpMCI
       ButtonPicture1(Index).Picture = UpImage.Picture
       lblStatus.Caption = "Mouse Up"
       MousePress = False
End Sub

Private Sub Form_Load()
Dim str1 As String

str1 = Space$(255)
MouseOverSound = "boink.wav"
MousePressSound = "bleeb.wav"
MouseUpSound = "type.wav"

''Load the sounds
LoadSound MouseOverSound, MouseOverMCI
LoadSound MousePressSound, MousePressMCI
LoadSound MouseUpSound, MouseUpMCI

Debug.Print mciSendString("PLAY WAVEUP11 FROM 0", str1, 0, 0)

Dim i As Integer
    lblStatus.Caption = "Ready?"
    For i = ButtonPicture1.LBound To ButtonPicture1.UBound
        ButtonPicture1(i).Picture = UpImage.Picture
    Next i
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not MouseOver Then Exit Sub
   StopSounds
   lblStatus.Caption = "Mouse Over - Form"
   MouseOver = False
   MousePress = False
   ButtonPicture1(NewIndex).Picture = UpImage.Picture
End Sub

Private Sub Form_Unload(Cancel As Integer)
 'This shouldn't be needed but it
 'can't hurt to stop the sound
  StopSounds
    
 'Unload the form and remove any references
  Unload Me
  Set Form1 = Nothing
End Sub

Public Function PlayWav(Alias As String)
 Dim rt As Long, ErrorString As String
 'Play the sound
 rt = mciSendString("PLAY " & Alias & " FROM 0", 0&, 0, 0)
 
 If rt <> 0 Then
    ErrorString = Space$(255)
    mciGetErrorString rt, ErrorString, Len(ErrorString)
    MsgBox "Error: " & ErrorString
 End If
 
End Function

Private Sub LoadSound(Filename As String, Alias As String)
Dim CommandString As String, ErrorString As String
Dim ShortPathName As String
Dim AppPath As String
Dim rt As Long

  ''Get the path name
  AppPath = App.Path
  If Right$(AppPath, 1) <> "\" Then
      AppPath = AppPath & "\"
  End If
    
  ''Allocate space for short path name
  ShortPathName = Space$(255)
  ''Get the short path name since MCI only accepts those
  GetShortPathName AppPath, ShortPathName, Len(ShortPathName)
  
  ''Remove empty spaces and the trailing NULL character
  ShortPathName = Left$(ShortPathName, Len(Trim$(ShortPathName)) - 1)
  'Build the command string
  CommandString = "OPEN " & ShortPathName & Filename & " TYPE WAVEAUDIO ALIAS " & Alias
  
  'Open the sound
   rt = mciSendString(CommandString, 0&, 0, 0)
   
   If rt <> 0 Then ''Non 0 = error
        ErrorString = Space$(255)
        mciGetErrorString rt, ErrorString, Len(ErrorString)
        MsgBox "Error: " & ErrorString
   End If

End Sub

Private Sub StopSounds()

    mciSendString "STOP " & MouseOverMCI, 0&, 0, 0
    mciSendString "STOP " & MouseUpMCI, 0&, 0, 0
    mciSendString "STOP " & MousePressMCI, 0&, 0, 0
    
End Sub

Private Sub UpImage_Click()

End Sub
