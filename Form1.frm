VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   BackColor       =   &H8000000B&
   Caption         =   "Playlist"
   ClientHeight    =   4710
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10440
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleWidth      =   10440
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Previous"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   3
      Top             =   3720
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6600
      TabIndex        =   2
      Top             =   3720
      Width           =   2775
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   960
      TabIndex        =   1
      Text            =   "Select song"
      Top             =   3000
      Width           =   8415
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   2535
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   8415
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   14843
      _cy             =   4471
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()
Form3.rs.MoveFirst
s = Form3.rs.Fields("id") & ". " & Form3.rs.Fields("name") & " - " & Form3.rs.Fields("artist") & " ( " & Form3.rs.Fields("language") & " - " & Form3.rs.Fields("year") & " )"
While s <> Combo1.Text
    Form3.rs.MoveNext
    s = Form3.rs.Fields("id") & ". " & Form3.rs.Fields("name") & " - " & Form3.rs.Fields("artist") & " ( " & Form3.rs.Fields("language") & " - " & Form3.rs.Fields("year") & " )"
Wend
WindowsMediaPlayer1.URL = App.Path & "\" & Form3.rs.Fields("path")
End Sub

Private Sub func_play()
Form3.rs.MoveFirst
s = Form3.rs.Fields("id") & ". " & Form3.rs.Fields("name") & " - " & Form3.rs.Fields("artist") & " ( " & Form3.rs.Fields("language") & " - " & Form3.rs.Fields("year") & " )"
While s <> Combo1.Text
    Form3.rs.MoveNext
    s = Form3.rs.Fields("id") & ". " & Form3.rs.Fields("name") & " - " & Form3.rs.Fields("artist") & " ( " & Form3.rs.Fields("language") & " - " & Form3.rs.Fields("year") & " )"
Wend
WindowsMediaPlayer1.URL = App.Path & "\" & Form3.rs.Fields("path")
WindowsMediaPlayer1.Controls.play
End Sub

Private Sub Command1_Click()
If Combo1.ListCount - 1 > Combo1.ListIndex Then
    Combo1.ListIndex = Combo1.ListIndex + 1
    func_play
Else
    MsgBox ("No songs in the queue"), vbExclamation, "Playlist ended"
End If
End Sub

Private Sub Command2_Click()
If Combo1.ListIndex > 0 Then
    Combo1.ListIndex = Combo1.ListIndex - 1
    func_play
End If
End Sub

Private Sub Form_Load()
n = 0
End Sub

Private Sub WindowsMediaPlayer1_PlayStateChange(ByVal NewState As Long)
If NewState = 8 And Combo1.ListCount - 1 > Combo1.ListIndex Then
    Combo1.ListIndex = Combo1.ListIndex + 1
    func_play
End If
If NewState = 10 Then
    n = n + 1
End If
If NewState = 3 Then
    n = 0
End If
If n = 3 Then
    n = 0
    WindowsMediaPlayer1.Controls.play
End If
End Sub

