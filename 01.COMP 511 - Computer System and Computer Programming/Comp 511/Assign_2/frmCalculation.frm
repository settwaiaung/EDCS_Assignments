VERSION 5.00
Begin VB.Form frmCalculation 
   Caption         =   "Calculation"
   ClientHeight    =   6075
   ClientLeft      =   6900
   ClientTop       =   2835
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6075
   ScaleWidth      =   6000
   Begin VB.TextBox txtSecondNumber 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   12
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox txtResult 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   11
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txtFirstNumber 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   10
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   6
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   5
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Frame frmOperation 
      Caption         =   "Operation"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   1935
      Begin VB.CommandButton cmdDivision 
         Caption         =   "Division"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   2280
         Width           =   1575
      End
      Begin VB.CommandButton cmdSubtract 
         Caption         =   "Subtract"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CommandButton cmdMultiplication 
         Caption         =   "Multiplication"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   1680
         Width           =   1575
      End
      Begin VB.CommandButton cmdSum 
         Caption         =   "Sum"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Label lblSecondnumber 
      Alignment       =   1  'Right Justify
      Caption         =   "Second Number ="
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   9
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label lblResult 
      Alignment       =   1  'Right Justify
      Caption         =   "Result ="
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   8
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label lblFirstNumber 
      Alignment       =   1  'Right Justify
      Caption         =   "First Number ="
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   1080
      Width           =   1935
   End
End
Attribute VB_Name = "frmCalculation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClear_Click()
txtFirstNumber.Text = ""
txtSecondNumber.Text = ""
txtResult.Text = ""
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub Form_Load()
Dim FirstNo As Double
Dim SecondNo As Double
Dim Result As Double
End Sub

Private Sub cmdSum_Click()
FirstNo = Val(txtFirstNumber.Text)
SecondNo = Val(txtSecondNumber.Text)
Result = FirstNo + SecondNo
txtResult.Text = Format(Result)
End Sub

Private Sub cmdSubtract_Click()
FirstNo = Val(txtFirstNumber.Text)
SecondNo = Val(txtSecondNumber.Text)
Result = FirstNo - SecondNo
txtResult.Text = Format(Result)

End Sub

Private Sub cmdMultiplication_Click()
FirstNo = Val(txtFirstNumber.Text)
SecondNo = Val(txtSecondNumber.Text)
Result = FirstNo * SecondNo
txtResult.Text = Format(Result)

End Sub
Private Sub cmdDivision_Click()
FirstNo = Val(txtFirstNumber.Text)
SecondNo = Val(txtSecondNumber.Text)
Result = FirstNo \ SecondNo
txtResult.Text = Format(Result)
End Sub

