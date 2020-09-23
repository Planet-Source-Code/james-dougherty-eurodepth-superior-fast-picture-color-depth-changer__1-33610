VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "EuroDepth Demo"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   578
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   424
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbDepth 
      Height          =   315
      Left            =   1973
      TabIndex        =   2
      Top             =   4170
      Width           =   2415
   End
   Begin VB.PictureBox picDestination 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   4110
      Left            =   0
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   270
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   420
      TabIndex        =   1
      Top             =   4560
      Width           =   6360
   End
   Begin VB.PictureBox picSource 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   4110
      Left            =   0
      Picture         =   "frmMain.frx":511F
      ScaleHeight     =   270
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   420
      TabIndex        =   0
      Top             =   0
      Width           =   6360
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CDepth As New ColorDepth

Private Sub cmbDepth_Click()
If cmbDepth.ListIndex = 0 Then
 CDepth.Set_1_4_8_BitDepth_2_16_256_Colors picSource, Colors2_1Bit, picDestination
ElseIf cmbDepth.ListIndex = 1 Then
 CDepth.Set_1_4_8_BitDepth_2_16_256_Colors picSource, Colors16_4Bit, picDestination
ElseIf cmbDepth.ListIndex = 2 Then
 CDepth.Set_1_4_8_BitDepth_2_16_256_Colors picSource, Colors256_8Bit, picDestination
ElseIf cmbDepth.ListIndex = 3 Then
 CDepth.Set24BitDepth_32k_64k_Colors picSource, picDestination
End If
End Sub

Private Sub Form_Load()
cmbDepth.AddItem "2 Colors (1 Bit)", 0
cmbDepth.AddItem "16 Colors (4 Bit)", 1
cmbDepth.AddItem "256 Colors (8 Bit)", 2
cmbDepth.AddItem "32k, 64k Colors (24 Bit)", 3
cmbDepth.ListIndex = 3
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set CDepth = Nothing
End Sub
