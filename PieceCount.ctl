VERSION 5.00
Begin VB.UserControl PieceCount 
   ClientHeight    =   4560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3195
   ControlContainer=   -1  'True
   ScaleHeight     =   4560
   ScaleWidth      =   3195
   Begin VB.TextBox Pieces_text 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   120
      TabIndex        =   11
      Top             =   150
      Width           =   2920
   End
   Begin VB.CommandButton Command1 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Index           =   9
      Left            =   2040
      TabIndex        =   9
      Top             =   555
      Width           =   1000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Index           =   8
      Left            =   1080
      TabIndex        =   8
      Top             =   555
      Width           =   1000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Index           =   7
      Left            =   120
      TabIndex        =   7
      Top             =   555
      Width           =   1000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Index           =   6
      Left            =   2040
      TabIndex        =   6
      Top             =   1515
      Width           =   1000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Index           =   5
      Left            =   1080
      TabIndex        =   5
      Top             =   1515
      Width           =   1000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   1515
      Width           =   1000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Index           =   3
      Left            =   2040
      TabIndex        =   3
      Top             =   2475
      Width           =   1000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Index           =   2
      Left            =   1080
      TabIndex        =   2
      Top             =   2475
      Width           =   1000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   2475
      Width           =   1000
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Index           =   10
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3435
      Width           =   1960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   3435
      Width           =   1000
   End
End
Attribute VB_Name = "PieceCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
    If Index < 10 Then
        Pieces_text.Text = Pieces_text + Trim(str(Index))
    Else
        Pieces_text.Text = ""
    End If

    
End Sub

Public Property Get PIECECOUNT() As String
    PIECECOUNT = Pieces_text.Text
    Pieces_text.Text = ""
End Property

Public Property Let CNC(oldCnc As String)

    Pieces_text.Text = oldCnc
End Property



