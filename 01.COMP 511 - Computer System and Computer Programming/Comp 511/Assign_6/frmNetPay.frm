VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Net Pay"
   ClientHeight    =   2610
   ClientLeft      =   9510
   ClientTop       =   6450
   ClientWidth     =   4905
   LinkTopic       =   "Form1"
   ScaleHeight     =   2610
   ScaleWidth      =   4905
   Begin VB.CommandButton cmdCal 
      Caption         =   "Calculate"
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txtTotal 
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Total Amount"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCal_Click()
T = CDbl(txtTotal.Text)

Select Case T

Case Is > 50000
N = T - (T * 20 / 100)
K = MsgBox("Your Net Pay is " & N & ".", vbInformation + vbOKOnly, "Net Pay")

Case Is > 30000
N = T - (T * 15 / 100)
K = MsgBox("Your Net Pay is " & N & ".", vbInformation + vbOKOnly, "Net Pay")

Case Is > 10000
N = T - (T * 10 / 100)
K = MsgBox("Your Net Pay is " & N & ".", vbInformation + vbOKOnly, "Net Pay")

Case Is <= 10000
K = MsgBox("Insufficient Input!", vbInformation + vbOKOnly, "Net Pay")

End Select
End Sub

Private Sub Form_Load()
Dim T As Double
Dim K As Double
Dim N As Double
End Sub
