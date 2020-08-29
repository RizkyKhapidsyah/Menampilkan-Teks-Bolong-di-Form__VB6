VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Menampilkan Teks Bolong di Form"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   5805
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   1215
      Left            =   720
      ScaleHeight     =   1155
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   600
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Const TXT = "Rizky Khapidsyah"
    Dim i As Long
    Dim hRgn As Long
    Picture1.AutoRedraw = True

    'Pilih huruf. Sesuaikan dengan keinginan Anda...
    Picture1.Font.Name = "Times New Roman"
    Picture1.Font.Bold = True
    Picture1.Font.Size = 50

    'Buat ukuran Picture1 cukup besar
    Picture1.Width = Picture1.TextWidth(TXT)
    Picture1.Height = Picture1.TextHeight(TXT)

    'Untuk letak Picture1
    BeginPath Picture1.hdc
    Picture1.CurrentX = 0
    Picture1.CurrentY = 0
    Picture1.Print TXT
    EndPath Picture1.hdc

    'Gambar teks...
    StrokePath Picture1.hdc
End Sub

