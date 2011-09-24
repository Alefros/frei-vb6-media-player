VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "Skin.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frm_wmp 
   Caption         =   "Alef Player"
   ClientHeight    =   10620
   ClientLeft      =   60
   ClientTop       =   735
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_wmp.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10620
   ScaleWidth      =   15240
   Begin VB.CommandButton cmd_proximo 
      Caption         =   ">>"
      Height          =   375
      Left            =   2880
      TabIndex        =   22
      Top             =   9840
      Width           =   615
   End
   Begin VB.CommandButton cmd_anterior 
      Caption         =   "<<"
      Height          =   375
      Left            =   2160
      TabIndex        =   21
      Top             =   9840
      Width           =   615
   End
   Begin ACTIVESKINLibCtl.Skin Skin6 
      Left            =   6960
      OleObjectBlob   =   "frm_wmp.frx":08CA
      Top             =   10080
   End
   Begin ACTIVESKINLibCtl.Skin Skin5 
      Left            =   6840
      OleObjectBlob   =   "frm_wmp.frx":3C285
      Top             =   10080
   End
   Begin ACTIVESKINLibCtl.Skin Skin4 
      Left            =   6720
      OleObjectBlob   =   "frm_wmp.frx":65C3E
      Top             =   10080
   End
   Begin ACTIVESKINLibCtl.Skin Skin3 
      Left            =   6600
      OleObjectBlob   =   "frm_wmp.frx":80483
      Top             =   10080
   End
   Begin ACTIVESKINLibCtl.Skin Skin2 
      Left            =   6480
      OleObjectBlob   =   "frm_wmp.frx":9635C
      Top             =   10080
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   495
      Left            =   4800
      TabIndex        =   19
      Top             =   9840
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      _Version        =   327682
      BorderStyle     =   1
      Max             =   100
      TickStyle       =   3
   End
   Begin ACTIVESKINLibCtl.SkinLabel skn_volume 
      Height          =   375
      Left            =   3840
      OleObjectBlob   =   "frm_wmp.frx":11845D
      TabIndex        =   18
      Top             =   9960
      Width           =   735
   End
   Begin VB.CommandButton cmd_play 
      Caption         =   "Play"
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   9840
      Width           =   855
   End
   Begin VB.CommandButton cmd_pause 
      Caption         =   "Pause"
      Height          =   375
      Left            =   1080
      TabIndex        =   16
      Top             =   9840
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   7200
      Top             =   9840
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   0
      TabIndex        =   9
      Top             =   8760
      Width           =   15255
      Begin VB.CommandButton cmd_mais 
         Caption         =   "+"
         Height          =   270
         Left            =   12600
         TabIndex        =   20
         ToolTipText     =   "Visualize e altere informações da música executada no momento"
         Top             =   600
         Width           =   255
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   12735
      End
      Begin VB.TextBox txt_musica 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   12375
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   12960
         OleObjectBlob   =   "frm_wmp.frx":1184BF
         TabIndex        =   12
         Top             =   600
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel lbl_total 
         Height          =   255
         Left            =   14040
         OleObjectBlob   =   "frm_wmp.frx":118523
         TabIndex        =   13
         Top             =   600
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   12960
         OleObjectBlob   =   "frm_wmp.frx":118589
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel lblTempo_corrido 
         Height          =   255
         Left            =   14040
         OleObjectBlob   =   "frm_wmp.frx":1185F1
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmds 
      Caption         =   "Tela cheia"
      Height          =   270
      Left            =   9840
      TabIndex        =   8
      Top             =   10080
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8895
      Left            =   10920
      TabIndex        =   3
      Top             =   -120
      Visible         =   0   'False
      Width           =   4335
      Begin VB.CommandButton cmd_addlista 
         Caption         =   "Adicionar"
         Height          =   375
         Left            =   1680
         TabIndex        =   7
         ToolTipText     =   "Adicionar a lista"
         Top             =   8400
         Width           =   1095
      End
      Begin VB.CommandButton cmd_esconde 
         Caption         =   "Esconder"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   6
         Top             =   8520
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Salvar Lista"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   8400
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid mfg_lista 
         Height          =   7935
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   13996
         _Version        =   393216
         GridLines       =   2
         FormatString    =   "           |                   Música                                     "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      FillColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   8775
      Left            =   120
      ScaleHeight     =   8745
      ScaleWidth      =   15225
      TabIndex        =   1
      Top             =   0
      Width           =   15255
      Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
         Height          =   2400
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   3300
         URL             =   ""
         rate            =   1
         balance         =   0
         currentPosition =   0
         defaultFrame    =   ""
         playCount       =   1
         autoStart       =   -1  'True
         currentMarker   =   0
         invokeURLs      =   -1  'True
         baseURL         =   ""
         volume          =   50
         mute            =   0   'False
         uiMode          =   "none"
         stretchToFit    =   0   'False
         windowlessVideo =   0   'False
         enabled         =   -1  'True
         enableContextMenu=   -1  'True
         fullScreen      =   0   'False
         SAMIStyle       =   ""
         SAMILang        =   ""
         SAMIFilename    =   ""
         captioningID    =   ""
         enableErrorDialogs=   0   'False
         _cx             =   5821
         _cy             =   4233
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   6360
      OleObjectBlob   =   "frm_wmp.frx":118657
      Top             =   10080
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7680
      Top             =   9960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmd_procurar 
      Caption         =   "Procurar"
      Height          =   375
      Left            =   8280
      TabIndex        =   0
      Top             =   9840
      Width           =   1215
   End
   Begin VB.Menu mnu_arquivo 
      Caption         =   "Arquivo"
      Begin VB.Menu mnu_novo 
         Caption         =   "Novo"
         Begin VB.Menu mnu_lr 
            Caption         =   "Lista de reprodução"
         End
      End
      Begin VB.Menu mnu_abrir 
         Caption         =   "Abrir"
      End
   End
   Begin VB.Menu mnu_exibir 
      Caption         =   "Exibir"
      Begin VB.Menu mnu_lrep 
         Caption         =   "Listas de reprodução"
      End
      Begin VB.Menu mnu_lista 
         Caption         =   "Lista de execução"
      End
   End
   Begin VB.Menu mnu_opt 
      Caption         =   "Opções"
      Begin VB.Menu mnu_conf 
         Caption         =   "Configurações"
      End
      Begin VB.Menu mnu_ap 
         Caption         =   "Aparência"
      End
   End
   Begin VB.Menu mnu_sobre 
      Caption         =   "Sobre"
   End
End
Attribute VB_Name = "frm_wmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim L_Colunas, l_linha, n As Long
Dim L_codcli As String
Dim L_codmusic
Dim num As String
Dim nome As String
Dim a As String
Dim cmusica As String 'Caminho da música
Dim g, autor As String
Dim cg, ca As Integer
Dim nlinha As String ' número da linha
Option Explicit

Private Sub cmdn_Click()
On Error Resume Next
            WindowsMediaPlayer1.fullScreen = False
End Sub
Private Sub cmd_addlista_Click()
                Dim autor As String
                Dim codautor As Integer
              With CommonDialog1
                .ShowOpen
                cmusica = .FileName
            End With
                Call abrir
                    If tabmusic.State = 1 Then tabmusic.Close
                        tabmusic.Open "select * from Musicas where caminho = '" & cmusica & "'"
                            If tabmusic.RecordCount = 0 Then
                                    Call regmusic 'registrar a música que ainda não foi executada
                                Exit Sub
                            ElseIf tabmusic.RecordCount = 1 Then
                                    codautor = tabmusic!autor
                                If tabautores.State = 1 Then tabautores.Close
                                    tabautores.Open "select * from autores where cod_autor = " & codautor
                                        If tabautores.RecordCount <> 0 Then
                                            autor = tabautores!autor
                                        End If
                                With mfg_lista
                                    .FormatString = " Código| Nome                                          | Autor                            |Caminho                                          "
                                    .TextMatrix(mfg_lista.Rows - 1, 0) = tabmusic!cod_musica
                                    .TextMatrix(mfg_lista.Rows - 1, 1) = tabmusic!nome
                                    .TextMatrix(mfg_lista.Rows - 1, 2) = autor
                                    .TextMatrix(mfg_lista.Rows - 1, 3) = tabmusic!caminho
                                End With
                        End If
                            mfg_lista.Rows = mfg_lista.Rows + 1
End Sub

Private Sub cmd_anterior_Click()
'            Dim n As Long
'
'            n = 2
'            n = l_linha
'            l_linha = l_linha - 1
'            L_codmusic = mfg_lista.TextMatrix(l_linha, 0)
'            txt_musica = mfg_lista.TextMatrix(l_linha, 1)
'                Call abrir
'                        If L_codmusic = "" Then
'                            Exit Sub
'                        End If
'                    If tabmusic.State = 1 Then tabmusic.Close
'                        tabmusic.Open "select * from Musicas where cod_musica = " & L_codmusic
'                        If tabmusic.RecordCount = 1 Then
'                            WindowsMediaPlayer1.URL = tabmusic!caminho
'                            WindowsMediaPlayer1.Controls.play
'                            Call cmd_play_Click
'                        End If
End Sub

Private Sub cmd_esconde_Click()
            Frame1.Visible = False
            WindowsMediaPlayer1.Width = WindowsMediaPlayer1.Width + Frame1.Width
End Sub



Private Sub cmd_pause_Click()
            WindowsMediaPlayer1.Controls.pause
End Sub

Private Sub cmd_play_Click()
            WindowsMediaPlayer1.Controls.play
End Sub

Private Sub cmd_proximo_Click()
            Dim n As Long
        
            n = "0"
            n = l_linha
            l_linha = l_linha + 1
            L_codmusic = mfg_lista.TextMatrix(l_linha, 0)
            txt_musica = mfg_lista.TextMatrix(l_linha, 1)
                Call abrir
                        If L_codmusic = "" Then
                            Exit Sub
                        End If
                    If tabmusic.State = 1 Then tabmusic.Close
                        tabmusic.Open "select * from Musicas where cod_musica = " & L_codmusic
                        If tabmusic.RecordCount = 1 Then
                            WindowsMediaPlayer1.URL = tabmusic!caminho
                            WindowsMediaPlayer1.Controls.play
                            Call cmd_play_Click
                        End If
                        
End Sub

Private Sub cmds_Click()
On Error Resume Next
If WindowsMediaPlayer1.fullScreen = True Then
    WindowsMediaPlayer1.fullScreen = False
    Exit Sub
ElseIf WindowsMediaPlayer1.fullScreen = False Then
        WindowsMediaPlayer1.fullScreen = True
        Exit Sub
End If

End Sub
Private Sub abrirmusica()
On Error Resume Next
        With CommonDialog1
                .Filter = "Todos os formatos"
                .DialogTitle = "Abrir mídia..."
                .ShowOpen
                
                WindowsMediaPlayer1.URL = .FileName
                txt_musica = .FileTitle
        End With
                cmusica = WindowsMediaPlayer1.URL
                        Call abrir
                    If tabmusic.State = 1 Then tabmusic.Close
                        tabmusic.Open "select * from Musicas where caminho = '" & cmusica & "'"
                            If tabmusic.RecordCount = 0 Then
                                Call regmusic
                            ElseIf tabmusic.RecordCount = 1 Then
                                    tabmusic!nome = txt_musica.Text
                                    tabmusic!caminho = cmusica
                                    tabmusic!duracao = WindowsMediaPlayer1.currentMedia.durationString
                                    tabmusic!genero = tabmusic!genero
                                    tabmusic!autor = tabmusic!autor
                                    tabmusic!execucoes = tabmusic!execucoes + 1
                                    tabmusic.Update
                            End If
End Sub
Private Sub regmusic()
            On Error Resume Next
            Call vga
                tabmusic.AddNew
                    tabmusic!nome = CommonDialog1.FileTitle
                    tabmusic!caminho = cmusica
                    tabmusic!duracao = WindowsMediaPlayer1.currentMedia.durationString
                    tabmusic!genero = cg
                    tabmusic!autor = ca
                    tabmusic!execucoes = "1"
                tabmusic.Update
                    Call cmd_addlista_Click
End Sub
Private Sub vga() 'vga = verifica genero e autor
            Call abrir
''''''''''''''' verificar / adicionar gênero
a:
                g = "Desconhecido"
                If tabgenero.State = 1 Then tabgenero.Close
                    tabgenero.Open "select * from Generos where genero = '" & g & "'"
                        If tabgenero.RecordCount = 0 Then
                            tabgenero.AddNew
                                tabgenero!genero = g
                            tabgenero.Update
                            GoTo a:
                        ElseIf tabgenero.RecordCount <> 0 Then
                                cg = tabgenero!cod_genero
                        End If
                        
''''''''''''''' verificar / adicionar autor
b:
                autor = "Desconhecido"
                If tabautores.State = 1 Then tabautores.Close
                    tabautores.Open "select * from Autores where autor = '" & autor & "'"
                        If tabautores.RecordCount = 0 Then
                            tabautores.AddNew
                                tabautores!autor = autor
                            tabautores.Update
                            GoTo b:
                        ElseIf tabautores.RecordCount <> 0 Then
                                ca = tabautores!cod_autor
                        End If
End Sub
Private Sub File1_Click()
            WindowsMediaPlayer1.URL = File1.FileName
End Sub
Private Sub Command3_Click()
            MsgBox "Descupe-nos função temporariamente indisponível !", vbCritical, "Alef Player"
                Exit Sub
        InputBox "Nome da lista...", "Alef Player"
End Sub
Private Sub Form_Load()
                 Call abrir
                Slider1.Value = WindowsMediaPlayer1.settings.volume
                WindowsMediaPlayer1.Height = Picture1.Height
                WindowsMediaPlayer1.Width = Picture1.Width
                Skin1.ApplySkin Me.hWnd
                    num = 1
            
End Sub
Private Sub fechar()
On Error Resume Next
Call abrir_banco
            If tabmusic.State = 1 Then tabmusic.Close
            If tabgenero.State = 1 Then tabgenero.Close
            If tablexe.State = 1 Then tablexe.Close
            If tablrepro.State = 1 Then tablrepro.Close
            If tabautores.State = 1 Then tabautores.Close
End Sub
Private Sub abrir()
            Call abrir_banco
            Call fechar
                tabmusic.Open "Musicas", conectar, adOpenKeyset, adLockOptimistic
                tabgenero.Open "Generos", conectar, adOpenKeyset, adLockOptimistic
                tablexe.Open "Listas_de_execucoes", conectar, adOpenKeyset, adLockOptimistic
                tablrepro.Open "Listas_de_reproducoes", conectar, adOpenKeyset, adLockOptimistic
                tabautores.Open "Autores", conectar, adOpenKeyset, adLockOptimistic
End Sub
Private Sub Text1_LostFocus()
'Text1 = Text1 + ".mp3"
End Sub
Private Sub lista()
Dim lista As String
            If List1.ListCount <> 0 Then
                    lista = List1.ListIndex = -1
                    a = lista
            End If
End Sub
Private Sub HScroll1_Scroll()
WindowsMediaPlayer1.Controls.currentPosition = HScroll1.Value
End Sub
Private Sub mfg_lista_Click()
        On Error Resume Next
            l_linha = mfg_lista.Row
            L_codmusic = mfg_lista.TextMatrix(l_linha, 0)
            L_codcli = mfg_lista.TextMatrix(l_linha, 2)
                Call abrir
                    If tabmusic.State = 1 Then tabmusic.Close
                        tabmusic.Open "select * from Musicas where cod_musica like '" & L_codmusic & "'"
                        If tabmusic.RecordCount = 1 Then
                           WindowsMediaPlayer1.URL = tabmusic!caminho
                        End If
            WindowsMediaPlayer1.Controls.play
'                mfg_lista.CellForeColor = vbBlue
            txt_musica = mfg_lista.TextMatrix(l_linha, 1)
End Sub
Private Sub mnu_abrir_Click()
            Call abrirmusica
End Sub
Private Sub mnu_ap_Click()
            frm_aparencia.Show
End Sub
Private Sub mnu_lista_Click()
            
            If mnu_lista.Caption = "Lista de execução" Then
                mnu_lista.Caption = "° Lista de execução"
                Frame1.Visible = True
                WindowsMediaPlayer1.Width = WindowsMediaPlayer1.Width - Frame1.Width
            ElseIf mnu_lista.Caption = "° Lista de execução" Then
                    mnu_lista.Caption = "Lista de execução"
                    Frame1.Visible = False
                    WindowsMediaPlayer1.Width = WindowsMediaPlayer1.Width + Frame1.Width
            End If
End Sub
Private Sub msg()
MsgBox "agora"
End Sub
Private Sub Slider1_Click()
            WindowsMediaPlayer1.settings.volume = Slider1.Value
End Sub
Private Sub Timer1_Timer()
            lblTempo_corrido.Caption = WindowsMediaPlayer1.Controls.currentPositionString
                If WindowsMediaPlayer1.playState = 3 Then
                    lbl_total.Caption = WindowsMediaPlayer1.currentMedia.durationString
                If WindowsMediaPlayer1.currentMedia.duration > 32767 Then
                Else
                    HScroll1.Max = WindowsMediaPlayer1.currentMedia.duration
                If HScroll1.Max = 0 Then
                Else
                    HScroll1.Value = WindowsMediaPlayer1.Controls.currentPosition
                End If
                End If
                End If
''''''''''''''''''''''''''''''''''''' Utilizar para trocar de música
'            If lblTempo_corrido = lbl_total Then
'                MsgBox "agora"
'''''''''
'                n = "0"
'                n = l_linha
'                l_linha = l_linha + 1
'                L_codmusic = mfg_lista.TextMatrix(l_linha, 0)
                
'''''''''            L_codcli = mfg_lista.TextMatrix(l_linha, 2)
'                WindowsMediaPlayer1.URL = L_codcli
'            End If

'''''''''
'''''''''            End If
End Sub
Private Sub WindowsMediaPlayer1_EndOfStream(ByVal Result As Long)
''''''
'MsgBox "agora"
End Sub
Private Sub WindowsMediaPlayer1_MediaChange(ByVal Item As Object)
'''''''Dar o play
WindowsMediaPlayer1.Controls.play
End Sub
Private Sub WindowsMediaPlayer1_PlaylistChange(ByVal Playlist As Object, ByVal change As WMPLibCtl.WMPPlaylistChangeEventType)
            WindowsMediaPlayer1.Controls.play
End Sub
Private Sub WindowsMediaPlayer1_StatusChange()
Dim n As Long
        If WindowsMediaPlayer1.playState = wmppsMediaEnded Then
            Call cmd_proximo_Click
'            n = "0"
'            n = l_linha
'            l_linha = l_linha + 1
'            L_codmusic = mfg_lista.TextMatrix(l_linha, 0)
'            txt_musica = mfg_lista.TextMatrix(l_linha, 1)
'                Call abrir
'                        If L_codmusic = "" Then
'                            Exit Sub
'                        End If
'                    If tabmusic.State = 1 Then tabmusic.Close
'                        tabmusic.Open "select * from Musicas where cod_musica = " & L_codmusic
'                        If tabmusic.RecordCount = 1 Then
'                            WindowsMediaPlayer1.URL = tabmusic!caminho
'                            WindowsMediaPlayer1.Controls.play
'                            Call cmd_play_Click
'                        End If
'        End If
'        On Error Resume Next
'        If WindowsMediaPlayer1.playState = 6 Then
'                    WindowsMediaPlayer1.Controls.play
                    
                    
        End If
End Sub
Private Sub tocar()
WindowsMediaPlayer1.Controls.play
End Sub

