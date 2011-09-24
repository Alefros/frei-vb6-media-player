VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frm_aparencia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aparência"
   ClientHeight    =   1980
   ClientLeft      =   750
   ClientTop       =   2580
   ClientWidth     =   2640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   2640
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton Opt1 
      Caption         =   "Option1"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.OptionButton Opt2 
      Caption         =   "Option2"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.OptionButton Opt3 
      Caption         =   "Option3"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Option4"
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.OptionButton Opt5 
      Caption         =   "Option5"
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.OptionButton Opt6 
      Caption         =   "Option6"
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.Skin Skin6 
      Left            =   1320
      OleObjectBlob   =   "frm_aparencia.frx":0000
      Top             =   1440
   End
   Begin ACTIVESKINLibCtl.Skin Skin5 
      Left            =   1200
      OleObjectBlob   =   "frm_aparencia.frx":3B9BB
      Top             =   1440
   End
   Begin ACTIVESKINLibCtl.Skin Skin4 
      Left            =   1080
      OleObjectBlob   =   "frm_aparencia.frx":65374
      Top             =   1440
   End
   Begin ACTIVESKINLibCtl.Skin Skin3 
      Left            =   960
      OleObjectBlob   =   "frm_aparencia.frx":7FBB9
      Top             =   1440
   End
   Begin ACTIVESKINLibCtl.Skin Skin2 
      Left            =   840
      OleObjectBlob   =   "frm_aparencia.frx":95A92
      Top             =   1440
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   720
      OleObjectBlob   =   "frm_aparencia.frx":117B93
      Top             =   1440
   End
   Begin ACTIVESKINLibCtl.Skin Skin7 
      Left            =   1320
      OleObjectBlob   =   "frm_aparencia.frx":1349FA
      Top             =   1440
   End
   Begin ACTIVESKINLibCtl.Skin Skin8 
      Left            =   1200
      OleObjectBlob   =   "frm_aparencia.frx":1703B5
      Top             =   1440
   End
   Begin ACTIVESKINLibCtl.Skin Skin9 
      Left            =   1080
      OleObjectBlob   =   "frm_aparencia.frx":199D6E
      Top             =   1440
   End
   Begin ACTIVESKINLibCtl.Skin Skin10 
      Left            =   960
      OleObjectBlob   =   "frm_aparencia.frx":1B45B3
      Top             =   1440
   End
   Begin ACTIVESKINLibCtl.Skin Skin11 
      Left            =   840
      OleObjectBlob   =   "frm_aparencia.frx":1CA48C
      Top             =   1440
   End
   Begin ACTIVESKINLibCtl.Skin Skin12 
      Left            =   720
      OleObjectBlob   =   "frm_aparencia.frx":24C58D
      Top             =   1440
   End
End
Attribute VB_Name = "frm_aparencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
'            Skin1.ApplySkin Me.hWnd
End Sub

Private Sub Opt1_Click()
            frm_wmp.Skin1.ApplySkin frm_wmp.hWnd
            Skin1.ApplySkin Me.hWnd
End Sub

Private Sub Opt2_Click()
            frm_wmp.Skin2.ApplySkin frm_wmp.hWnd
            Skin2.ApplySkin Me.hWnd
End Sub

Private Sub Opt3_Click()
            frm_wmp.Skin3.ApplySkin frm_wmp.hWnd
            Skin3.ApplySkin Me.hWnd
End Sub

Private Sub Opt5_Click()
            frm_wmp.Skin5.ApplySkin frm_wmp.hWnd
            Skin5.ApplySkin Me.hWnd
End Sub

Private Sub Opt6_Click()
            frm_wmp.Skin6.ApplySkin frm_wmp.hWnd
            Skin6.ApplySkin Me.hWnd
End Sub

Private Sub Option4_Click()
            frm_wmp.Skin4.ApplySkin Me.hWnd
            Skin4.ApplySkin Me.hWnd
End Sub
