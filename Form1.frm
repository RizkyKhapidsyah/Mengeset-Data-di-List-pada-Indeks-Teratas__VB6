VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Mengeset Data di List pada Indeks Teratas"
   ClientHeight    =   1605
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3765
   LinkTopic       =   "Form1"
   ScaleHeight     =   1605
   ScaleWidth      =   3765
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   1200
      Width           =   2415
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
'Ketika Anda memilih data di Combo1, maka data indeks 'yang sesuai dengan indeks Combo1 terpilih akan berada 'paling atas dari List1.
    Dim indx As Integer
    indx = Combo1.Text
    List1.TopIndex = indx
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim itemName As String
    For i = 0 To 100
        itemName = "Ini data indeks ke-" & i
        List1.AddItem itemName
        Combo1.AddItem i
    Next
    Combo1.Text = Combo1.List(0)
End Sub

