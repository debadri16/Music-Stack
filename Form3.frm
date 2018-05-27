VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Music Directory"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12960
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   12960
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000006&
      Caption         =   "STEP    FORWARD   >>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   3480
      TabIndex        =   0
      Top             =   6120
      Width           =   5775
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "  MUSIC      STACK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3240
      TabIndex        =   1
      Top             =   1200
      Width           =   6135
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public db As Database, rs As Recordset, n As Long

Private Sub Command1_Click()
    Form2.Show
    Unload Me
End Sub

Private Sub Form_Load()
Set db = OpenDatabase(App.Path & "\music.mdb")
Set rs = db.OpenRecordset("select * from songs")
End Sub

