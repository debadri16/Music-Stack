VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H80000002&
   Caption         =   "Filter songs"
   ClientHeight    =   4185
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7635
   LinkTopic       =   "Form2"
   ScaleHeight     =   4185
   ScaleWidth      =   7635
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3240
      TabIndex        =   7
      Top             =   2280
      Width           =   2775
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3240
      TabIndex        =   5
      Top             =   1800
      Width           =   2775
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3240
      TabIndex        =   4
      Top             =   1320
      Width           =   2775
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3240
      TabIndex        =   3
      Top             =   840
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2400
      TabIndex        =   2
      Top             =   3000
      Width           =   2235
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select your choice(one or more) or keep blank"
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
      Left            =   600
      TabIndex        =   9
      Top             =   240
      Width           =   6135
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "YEAR : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "LANGUAGE : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "ARTIST : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "NAME : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   840
      Width           =   1695
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form3.rs.MoveFirst
Form1.Combo1.Clear
Form1.Combo1.Text = "Select Song:"
If Combo1.Text = "" Then
    If Combo2.Text = "" Then
        If Combo3.Text = "" Then
            If Combo4.Text = "" Then
                While Form3.rs.EOF <> True
                    Form1.Combo1.AddItem (Form3.rs.Fields("id") & ". " & Form3.rs.Fields("name") & " - " & Form3.rs.Fields("artist") & " ( " & Form3.rs.Fields("language") & " - " & Form3.rs.Fields("year") & " )")
                    Form3.rs.MoveNext
                Wend
            Else
                While Form3.rs.EOF <> True
                    If Combo4.Text = Form3.rs.Fields("year") Then
                        Form1.Combo1.AddItem (Form3.rs.Fields("id") & ". " & Form3.rs.Fields("name") & " - " & Form3.rs.Fields("artist") & " ( " & Form3.rs.Fields("language") & " - " & Form3.rs.Fields("year") & " )")
                    End If
                    Form3.rs.MoveNext
                Wend
            End If
        Else
            If Combo4.Text = "" Then
                While Form3.rs.EOF <> True
                    If Combo3.Text = Form3.rs.Fields("language") Then
                        Form1.Combo1.AddItem (Form3.rs.Fields("id") & ". " & Form3.rs.Fields("name") & " - " & Form3.rs.Fields("artist") & " ( " & Form3.rs.Fields("language") & " - " & Form3.rs.Fields("year") & " )")
                    End If
                    Form3.rs.MoveNext
                Wend
            Else
                While Form3.rs.EOF <> True
                    If Combo4.Text = Form3.rs.Fields("year") And Combo3.Text = Form3.rs.Fields("language") Then
                        Form1.Combo1.AddItem (Form3.rs.Fields("id") & ". " & Form3.rs.Fields("name") & " - " & Form3.rs.Fields("artist") & " ( " & Form3.rs.Fields("language") & " - " & Form3.rs.Fields("year") & " )")
                    End If
                    Form3.rs.MoveNext
                Wend
            End If
        End If
    Else
        If Combo3.Text = "" Then
            If Combo4.Text = "" Then
                While Form3.rs.EOF <> True
                    If Combo2.Text = Form3.rs.Fields("artist") Then
                        Form1.Combo1.AddItem (Form3.rs.Fields("id") & ". " & Form3.rs.Fields("name") & " - " & Form3.rs.Fields("artist") & " ( " & Form3.rs.Fields("language") & " - " & Form3.rs.Fields("year") & " )")
                    End If
                    Form3.rs.MoveNext
                Wend
            Else
                While Form3.rs.EOF <> True
                    If Combo4.Text = Form3.rs.Fields("year") And Combo2.Text = Form3.rs.Fields("artist") Then
                        Form1.Combo1.AddItem (Form3.rs.Fields("id") & ". " & Form3.rs.Fields("name") & " - " & Form3.rs.Fields("artist") & " ( " & Form3.rs.Fields("language") & " - " & Form3.rs.Fields("year") & " )")
                    End If
                    Form3.rs.MoveNext
                Wend
            End If
        Else
            If Combo4.Text = "" Then
                While Form3.rs.EOF <> True
                    If Combo3.Text = Form3.rs.Fields("language") And Combo2.Text = Form3.rs.Fields("artist") Then
                        Form1.Combo1.AddItem (Form3.rs.Fields("id") & ". " & Form3.rs.Fields("name") & " - " & Form3.rs.Fields("artist") & " ( " & Form3.rs.Fields("language") & " - " & Form3.rs.Fields("year") & " )")
                    End If
                    Form3.rs.MoveNext
                Wend
            Else
                While Form3.rs.EOF <> True
                    If Combo4.Text = Form3.rs.Fields("year") And Combo3.Text = Form3.rs.Fields("language") And Combo2.Text = Form3.rs.Fields("artist") Then
                        Form1.Combo1.AddItem (Form3.rs.Fields("id") & ". " & Form3.rs.Fields("name") & " - " & Form3.rs.Fields("artist") & " ( " & Form3.rs.Fields("language") & " - " & Form3.rs.Fields("year") & " )")
                    End If
                    Form3.rs.MoveNext
                Wend
            End If
        End If
    End If
Else
    If Combo2.Text = "" Then
        If Combo3.Text = "" Then
            If Combo4.Text = "" Then
                While Form3.rs.EOF <> True
                    If Combo1.Text = Form3.rs.Fields("name") Then
                        Form1.Combo1.AddItem (Form3.rs.Fields("id") & ". " & Form3.rs.Fields("name") & " - " & Form3.rs.Fields("artist") & " ( " & Form3.rs.Fields("language") & " - " & Form3.rs.Fields("year") & " )")
                    End If
                    Form3.rs.MoveNext
                Wend
            Else
                While Form3.rs.EOF <> True
                    If Combo4.Text = Form3.rs.Fields("year") And Combo1.Text = Form3.rs.Fields("name") Then
                        Form1.Combo1.AddItem (Form3.rs.Fields("id") & ". " & Form3.rs.Fields("name") & " - " & Form3.rs.Fields("artist") & " ( " & Form3.rs.Fields("language") & " - " & Form3.rs.Fields("year") & " )")
                    End If
                    Form3.rs.MoveNext
                Wend
            End If
        Else
            If Combo4.Text = "" Then
                While Form3.rs.EOF <> True
                    If Combo3.Text = Form3.rs.Fields("language") And Combo1.Text = Form3.rs.Fields("name") Then
                        Form1.Combo1.AddItem (Form3.rs.Fields("id") & ". " & Form3.rs.Fields("name") & " - " & Form3.rs.Fields("artist") & " ( " & Form3.rs.Fields("language") & " - " & Form3.rs.Fields("year") & " )")
                    End If
                    Form3.rs.MoveNext
                Wend
            Else
                While Form3.rs.EOF <> True
                    If Combo4.Text = Form3.rs.Fields("year") And Combo1.Text = Form3.rs.Fields("name") And Combo3.Text = Form3.rs.Fields("language") Then
                        Form1.Combo1.AddItem (Form3.rs.Fields("id") & ". " & Form3.rs.Fields("name") & " - " & Form3.rs.Fields("artist") & " ( " & Form3.rs.Fields("language") & " - " & Form3.rs.Fields("year") & " )")
                    End If
                    Form3.rs.MoveNext
                Wend
            End If
        End If
    Else
        If Combo3.Text = "" Then
            If Combo4.Text = "" Then
                While Form3.rs.EOF <> True
                    If Combo2.Text = Form3.rs.Fields("artist") And Combo1.Text = Form3.rs.Fields("name") Then
                        Form1.Combo1.AddItem (Form3.rs.Fields("id") & ". " & Form3.rs.Fields("name") & " - " & Form3.rs.Fields("artist") & " ( " & Form3.rs.Fields("language") & " - " & Form3.rs.Fields("year") & " )")
                    End If
                    Form3.rs.MoveNext
                Wend
            Else
                While Form3.rs.EOF <> True
                    If Combo4.Text = Form3.rs.Fields("year") And Combo1.Text = Form3.rs.Fields("name") And Combo2.Text = Form3.rs.Fields("artist") Then
                        Form1.Combo1.AddItem (Form3.rs.Fields("id") & ". " & Form3.rs.Fields("name") & " - " & Form3.rs.Fields("artist") & " ( " & Form3.rs.Fields("language") & " - " & Form3.rs.Fields("year") & " )")
                    End If
                    Form3.rs.MoveNext
                Wend
            End If
        Else
            If Combo4.Text = "" Then
                While Form3.rs.EOF <> True
                    If Combo3.Text = Form3.rs.Fields("language") And Combo1.Text = Form3.rs.Fields("name") And Combo2.Text = Form3.rs.Fields("artist") Then
                        Form1.Combo1.AddItem (Form3.rs.Fields("id") & ". " & Form3.rs.Fields("name") & " - " & Form3.rs.Fields("artist") & " ( " & Form3.rs.Fields("language") & " - " & Form3.rs.Fields("year") & " )")
                    End If
                    Form3.rs.MoveNext
                Wend
            Else
                While Form3.rs.EOF <> True
                    If Combo4.Text = Form3.rs.Fields("year") And Combo1.Text = Form3.rs.Fields("name") And Combo3.Text = Form3.rs.Fields("language") And Combo2.Text = Form3.rs.Fields("artist") Then
                        Form1.Combo1.AddItem (Form3.rs.Fields("id") & ". " & Form3.rs.Fields("name") & " - " & Form3.rs.Fields("artist") & " ( " & Form3.rs.Fields("language") & " - " & Form3.rs.Fields("year") & " )")
                    End If
                    Form3.rs.MoveNext
                Wend
            End If
        End If
    End If
End If
Form1.Show
End Sub

Private Sub Form_Load()
Form3.rs.MoveFirst
Combo3.AddItem ("Hindi")
Combo3.AddItem ("Bengali")
Combo3.AddItem ("English")
While Form3.rs.EOF <> True
    Combo1.AddItem (Form3.rs.Fields("name"))
    
    add_artist (Form3.rs.Fields("artist"))
    add_year (Form3.rs.Fields("year"))
    
    Form3.rs.MoveNext
Wend
End Sub

Private Sub add_artist(str As String)
n = Combo2.ListCount
i = 0
flag = 0
While i < n
    If Combo2.List(i) = str Then
        flag = 1
    End If
    i = i + 1
Wend
If flag = 0 Then
    Combo2.AddItem (str)
End If
End Sub

Private Sub add_year(str2 As String)
n2 = Combo4.ListCount
i2 = 0
flag2 = 0
indx = -1
While i2 < n2
    If Val(Combo4.List(i2)) < Val(str2) Then
        indx = indx + 1
    End If
    If Combo4.List(i2) = str2 Then
        flag2 = 1
    End If
    i2 = i2 + 1
Wend
If flag2 = 0 Then
    Combo4.AddItem str2, indx + 1
End If
End Sub

