VERSION 5.00
Begin VB.Form Pieces 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CNC Selection"
   ClientHeight    =   6375
   ClientLeft      =   7350
   ClientTop       =   5325
   ClientWidth     =   4350
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   4350
   Begin VB.Timer TimerCncPicker 
      Interval        =   10000
      Left            =   120
      Top             =   3600
   End
   Begin PaperlessCell.PieceCount PieceCount1 
      Height          =   4335
      Left            =   600
      TabIndex        =   3
      Top             =   840
      Width           =   3135
      _extentx        =   5530
      _extenty        =   7646
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2280
      TabIndex        =   2
      Top             =   5400
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "ENTER CNC"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "Pieces"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Piece_Count = Val(Trim(PieceCount1.PIECECOUNT))
If Piece_Count > 32676 Then
    MsgBox ("INVALID, CNC RE-ENTER")
    Exit Sub
End If
 newCnc = Piece_Count
 Pieces.Hide
 DoEvents
End Sub

Private Sub Command2_Click()
    Pieces.Hide
    DoEvents
End Sub

Private Sub Form_Activate()
    If newCnc <> 0 Then
        PieceCount1.CNC = str(newCnc)
    Else
        PieceCount1.CNC = ""
    End If
    TimerCncPicker.Enabled = True
End Sub


Private Sub Form_Load()
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 4
End Sub

Private Sub TimerCncPicker_Timer()
    TimerCncPicker.Enabled = False
    Pieces.Hide
    DoEvents
End Sub
