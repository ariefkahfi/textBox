VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3735
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   ScaleHeight     =   3735
   ScaleWidth      =   6990
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   495
      Left            =   5160
      TabIndex        =   3
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdMessage 
      Caption         =   "My Message"
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Date"
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   2175
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDelete_Click()
    Text1.Text = ""
End Sub

Private Sub cmdEdit_Click()
    Dim Tanggal
    Tanggal = Format(Date, "dddd, d mmm yyyy")
    Text1.Text = "Now is: " + Tanggal
End Sub

Private Sub cmdMessage_Click()
    Text1.Text = "Belajar Visual Basic ? Sapa Takut"
End Sub

Private Sub Form_Load()
    Text1.Text = ""
End Sub
