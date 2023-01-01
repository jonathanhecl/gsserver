VERSION 5.00
Begin VB.Form frmG_Main 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GSS >> Panel de Control"
   ClientHeight    =   6375
   ClientLeft      =   4725
   ClientTop       =   2820
   ClientWidth     =   9720
   Icon            =   "frmG_Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6375
   ScaleWidth      =   9720
   Begin VB.Timer Actualizador 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   9120
      Top             =   5160
   End
   Begin VB.CommandButton cmdActualizar 
      BackColor       =   &H0000FF00&
      Caption         =   "&Actualizar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4920
      Width           =   4335
   End
   Begin VB.CheckBox Actualizar 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000C000&
      Caption         =   "Actualizar automaticamente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   18
      Top             =   5400
      Width           =   3855
   End
   Begin VB.ListBox LstUsuarios 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   2955
      Left            =   5040
      TabIndex        =   17
      Top             =   550
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "&Escaneador de PJs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5880
      Width           =   4815
   End
   Begin VB.TextBox Mensaje 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   5055
      MaxLength       =   1024
      TabIndex        =   15
      Top             =   3600
      Width           =   3840
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FF00&
      Caption         =   ">>"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Left            =   8940
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Enviar Privado"
      Top             =   3600
      Width           =   495
   End
   Begin VB.ListBox MSX 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   2340
      ItemData        =   "frmG_Main.frx":1042
      Left            =   240
      List            =   "frmG_Main.frx":1044
      TabIndex        =   13
      Top             =   3795
      Width           =   4335
   End
   Begin VB.CommandButton cmdResetear 
      BackColor       =   &H0000FF00&
      Caption         =   "&Resetear todos los Socket's"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3720
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.TextBox BroadMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   360
      MaxLength       =   255
      TabIndex        =   0
      Top             =   480
      Width           =   4215
   End
   Begin VB.CommandButton cmdEnviarGM 
      BackColor       =   &H0000FF00&
      Caption         =   "&Enviar por Consola"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1200
      Width           =   2055
   End
   Begin VB.CommandButton cmdHacerBackUp 
      BackColor       =   &H0000FF00&
      Caption         =   "&Hacer un BackUp del mundo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2640
      Width           =   4335
   End
   Begin VB.CommandButton cmdCargarBackup 
      BackColor       =   &H0000FF00&
      Caption         =   "&Cargar el BackUp del mundo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2520
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.CommandButton cmdGrabar 
      BackColor       =   &H0000FF00&
      Caption         =   "&Guardar los Personajes y Clanes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3120
      Width           =   4335
   End
   Begin VB.CommandButton cmdEnviarT 
      BackColor       =   &H0000FF00&
      Caption         =   "E&nviar por Consola"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   2175
   End
   Begin VB.CommandButton cmdVentanaT 
      BackColor       =   &H0000FF00&
      Caption         =   "En&viar por Ventana"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1560
      Width           =   2175
   End
   Begin VB.CommandButton cmdEnviarVentanaGM 
      BackColor       =   &H0000FF00&
      Caption         =   "Envia&r por Ventana"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label EstadoDat 
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   6960
      TabIndex        =   25
      Top             =   4230
      Width           =   2415
   End
   Begin VB.Label PosDat 
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   6960
      TabIndex        =   24
      Top             =   4500
      Width           =   2415
   End
   Begin VB.Label MPDat 
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   5040
      TabIndex        =   23
      Top             =   4500
      Width           =   1815
   End
   Begin VB.Label HPDat 
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   5040
      TabIndex        =   22
      Top             =   4230
      Width           =   1815
   End
   Begin VB.Label Jugando 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   " Usuarios Jugando: 0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5400
      TabIndex        =   21
      Top             =   240
      Width           =   3660
   End
   Begin VB.Label NickDat 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5040
      TabIndex        =   20
      Top             =   3960
      Width           =   4335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      Caption         =   " Mensajes:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   900
   End
   Begin VB.Shape Shape4 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   2535
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   3720
      Width           =   4575
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   345
      Top             =   465
      Width           =   4240
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      Caption         =   " Para los GM's:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2520
      TabIndex        =   8
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      Caption         =   " Para TODOS:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      Caption         =   " Acciones:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   2280
      Width           =   840
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   1935
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   4575
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   1455
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   2160
      Width           =   4575
   End
   Begin VB.Shape Shape7 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   2980
      Left            =   5030
      Top             =   540
      Width           =   4400
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   5045
      Top             =   3580
      Width           =   3885
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   5655
      Left            =   4800
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "frmG_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Actualizador_Timer()
On Error Resume Next
Dim NickDATT As String
NickDATT = NickDat.Caption
If cmdActualizar.Enabled = True Then Call cmdActualizar_Click
NickDat.Caption = NickDATT
Call InfoUser
End Sub

Private Sub Actualizar_Click()
If Actualizar.Value = 1 Then
    Actualizador.Enabled = True
Else
    Actualizador.Enabled = False
End If
End Sub

Private Sub cmdActualizar_Click()
cmdActualizar.Enabled = False
DoEvents
NickDat.Caption = ""
Dim numeroJuaz As Integer
numeroJuaz = 0
' La lista
Dim h As Long
LstUsuarios.Clear
For h = 1 To LastUser
    If UserList(h).ConnID <> -1 And UserList(h).flags.UserLogged = True Then
        ' Nombre " - Nivel: " El nivel - Num:~" Numero
        LstUsuarios.AddItem UserList(h).Name & " - Nivel: " & UserList(h).Stats.ELV & " - Num:~" & str(h)
        numeroJuaz = numeroJuaz + 1
    ElseIf UserList(h).ConnID <> -1 Then
        LstUsuarios.AddItem "Usuario conectandose..."
    End If
Next h

NumUsers = numeroJuaz

Jugando.Caption = "Usuarios Jugando: " & NumUsers

cmdActualizar.Enabled = True
End Sub

Private Sub cmdCargarBackup_Click()

cmdCargarBackup.Enabled = False
cmdGrabar.Enabled = False
cmdHacerBackUp.Enabled = False


'Se asegura de que los sockets estan cerrados e ignora cualquier err
On Error Resume Next

If frmGeneral.Visible Then frmGeneral.Estado.SimpleText = "Reiniciando."
' Barra de progreso!!
FrmStat.Show

If FileExist(App.Path & "\logs\errores.log", vbNormal) Then Kill App.Path & "\logs\errores.log"
If FileExist(App.Path & "\logs\connect.log", vbNormal) Then Kill App.Path & "\logs\Connect.log"
If FileExist(App.Path & "\logs\HackAttemps.log", vbNormal) Then Kill App.Path & "\logs\HackAttemps.log"
If FileExist(App.Path & "\logs\Asesinatos.log", vbNormal) Then Kill App.Path & "\logs\Asesinatos.log"
If FileExist(App.Path & "\logs\Resurrecciones.log", vbNormal) Then Kill App.Path & "\logs\Resurrecciones.log"
If FileExist(App.Path & "\logs\Teleports.Log", vbNormal) Then Kill App.Path & "\logs\Teleports.Log"


#If UsarAPI Then
Call apiclosesocket(SockListen)
#Else
frmGeneral.Socket1.Cleanup
frmGeneral.Socket2(0).Cleanup
#End If

Dim LoopC As Integer

For LoopC = 1 To MaxUsers
    Call CloseSocket(LoopC)
Next
  

LastUser = 0
NumUsers = 0

ReDim Npclist(1 To MAXNPCS) As Npc 'NPCS
ReDim CharList(1 To MAXCHARS) As Integer

Call LoadSini
Call CargarBackUp
Call LoadOBJData_Nuevo

#If UsarAPI Then
SockListen = ListenForConnect(Puerto, frmGeneral.hwnd, "")

#Else
frmGeneral.Socket1.AddressFamily = AF_INET
frmGeneral.Socket1.protocol = IPPROTO_IP
frmGeneral.Socket1.SocketType = SOCK_STREAM
frmGeneral.Socket1.Binary = False
frmGeneral.Socket1.Blocking = False
frmGeneral.Socket1.BufferSize = 1024

frmGeneral.Socket2(0).AddressFamily = AF_INET
frmGeneral.Socket2(0).protocol = IPPROTO_IP
frmGeneral.Socket2(0).SocketType = SOCK_STREAM
frmGeneral.Socket2(0).Blocking = False
frmGeneral.Socket2(0).BufferSize = 2048

'Escucha
frmGeneral.Socket1.LocalPort = Puerto
frmGeneral.Socket1.listen
#End If

If frmGeneral.Visible Then frmGeneral.Estado.SimpleText = "Escuchando conexiones entrantes ..."

cmdCargarBackup.Enabled = True
cmdGrabar.Enabled = True
cmdHacerBackUp.Enabled = True


End Sub

Private Sub cmdEnviarGM_Click()
Call SendData(ToAdmins, 0, 0, "||" & "GM's: " & BroadMsg.Text & FONTTYPE_FIGHT & ENDC)

End Sub

Private Sub cmdEnviarT_Click()
Call SendData(ToAll, 0, 0, "||<Host> " & BroadMsg.Text & FONTTYPE_TALK & ENDC)

End Sub

Private Sub cmdEnviarVentanaGM_Click()
Call SendData(ToAdmins, 0, 0, "!!" & BroadMsg.Text & ENDC)

End Sub

Private Sub cmdGrabar_Click()
cmdCargarBackup.Enabled = False
cmdGrabar.Enabled = False
cmdHacerBackUp.Enabled = False
Me.MousePointer = 11
Call GuardarUsuarios
Call SaveGuildsDB
Me.MousePointer = 0
Call FrmMensajes.msg("Nota", "Personajes y Clanes guardados!")
cmdCargarBackup.Enabled = True
cmdGrabar.Enabled = True
cmdHacerBackUp.Enabled = True
End Sub

Private Sub cmdHacerBackUp_Click()
cmdCargarBackup.Enabled = False
cmdGrabar.Enabled = False
cmdHacerBackUp.Enabled = False
On Error GoTo eh
    Me.MousePointer = 11
    FrmStat.Show
    Call DoBackUp
    Me.MousePointer = 0
    Call FrmMensajes.msg("Nota", "WORLDSAVE OK!!")
    cmdCargarBackup.Enabled = True
    cmdGrabar.Enabled = True
    cmdHacerBackUp.Enabled = True
Exit Sub
eh:
Call LogError("Error en WORLDSAVE")
cmdCargarBackup.Enabled = True
cmdGrabar.Enabled = True
cmdHacerBackUp.Enabled = True
End Sub

Private Sub cmdResetear_Click()
#If UsarAPI Then
Dim i As Long

If MsgBox("Esta seguro que desea Reiniciar los Socket's? Se cerrarán todas las conexiones activas.", vbYesNo, "Reiniciar Sockets") = vbYes Then
    'Cierra el socket de escucha
    If SockListen >= 0 Then Call apiclosesocket(SockListen)
    
    'Cierra todas las conexiones
    For i = 1 To MaxUsers
        If UserList(i).ConnID <> -1 Then
            Call CloseSocket(i)
        End If
    Next i
    
    'Inicia el socket de escucha
    SockListen = ListenForConnect(Puerto, hWndMsg, "")
    
    'Comprueba si el proc de la ventana es el correcto
    Dim TmpWProc As Long
    TmpWProc = GetWindowLong(hWndMsg, GWL_WNDPROC)
    If TmpWProc <> ActualWProc Then
        MsgBox "Incorrecto proc de ventana (" & TmpWProc & " <> " & ActualWProc & ")"
        Call LogApiSock("INCORRECTO PROC DE VENTANA")
        OldWProc = TmpWProc
        If OldWProc <> 0 Then
            SetWindowLong frmMain.hwnd, GWL_WNDPROC, AddressOf WndProc
            ActualWProc = GetWindowLong(frmMain.hwnd, GWL_WNDPROC)
        End If
    End If
End If
#End If

End Sub

Private Sub cmdVentanaT_Click()
Call SendData(ToAll, 0, 0, "!!" & BroadMsg.Text & ENDC)

End Sub

Private Sub Command1_Click()
On Error Resume Next
EscaneadorDePJs.Show
End Sub

Private Sub Command2_Click()
On Error Resume Next
Dim Nick As String
Dim tIndex As Integer
If Len(NickDat.Caption) > 1 Then
    Nick = ReadField(1, NickDat.Caption, Asc("-"))
    Nick = Left(Nick, Len(Nick) - 1)
    tIndex = NameIndex(Nick)
    If tIndex <= 0 Then Exit Sub
    Call SendData(ToIndex, tIndex, 0, "||HOST le dice a Usted: " & Mensaje.Text & FONTTYPE_WHISPER)
    Call LogCOSAS("Host", Time & " HOST le dice a Usted: " & Mensaje.Text)
    Mensaje.Text = ""
End If
End Sub

Private Sub Form_Activate()
Actualizar.Value = 1
Actualizador.Enabled = True
End Sub

Private Sub Form_Initialize()
'XP
End Sub

Private Sub Form_Load()
Dim NoxKo As Boolean
NoxKo = True
If frmGeneral.Visible = False Then NoxKo = True
If NoxKo = True Then
    frmGeneral.Visible = False
    Me.Visible = False
End If
Me.Left = 0
Me.Top = 0
Call cmdActualizar_Click

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Me.Hide
If frmGeneral.mnuCerrarCorrectamente.Checked = True Then Exit Sub
If frmGeneral.mnuCerrar.Checked = True Then Exit Sub
Cancel = True
End Sub

Private Sub LstUsuarios_Click()
On Error Resume Next
NickDat.Caption = LstUsuarios.Text
Call InfoUser
End Sub

' [GS] InfoUser ;)
Sub InfoUser()
On Error Resume Next
If frmG_Main.Visible = False Then Exit Sub
If Len(NickDat.Caption) < 2 Then Exit Sub
Dim Nick As String
Dim tIndex As Integer

Nick = ReadField(1, NickDat.Caption, Asc("-"))
Nick = Left(Nick, Len(Nick) - 1)
tIndex = NameIndex(Nick)
If tIndex <= 0 Then
    NickDat.Caption = ""
    HPDat.Caption = "HP: -/-"
    MPDat.Caption = "MP: -/-"
    PosDat.Caption = "Pos: -"
    EstadoDat.Caption = "Estado: -"
    ' Esta offline?
    Exit Sub
Else
    ' Esta online :D
    HPDat.Caption = "HP: " & UserList(tIndex).Stats.MinHP & "/" & UserList(tIndex).Stats.MaxHP
    MPDat.Caption = "MP: " & UserList(tIndex).Stats.MinMAN & "/" & UserList(tIndex).Stats.MaxMAN
    PosDat.Caption = "Pos: " & UserList(tIndex).Pos.Map & " - " & UserList(tIndex).Pos.X & "," & UserList(tIndex).Pos.Y
    If UserList(tIndex).Counters.Saliendo = True Then
        EstadoDat.Caption = "Estado: Saliendo"
    ElseIf UserList(tIndex).flags.PuedeAtacar = False Then
        EstadoDat.Caption = "Estado: Atacando"
    ElseIf UserList(tIndex).flags.PuedeLanzarSpell = False Then
        EstadoDat.Caption = "Estado: Lanzando Hechizos"
    ElseIf UserList(tIndex).flags.Ceguera = 1 Then
        EstadoDat.Caption = "Estado: Cegado"
    ElseIf UserList(tIndex).flags.Trabajando = True Then
        EstadoDat.Caption = "Estado: Trabajando"
    ElseIf UserList(tIndex).flags.Comerciando = True Then
        EstadoDat.Caption = "Estado: Comerciando"
    ElseIf UserList(tIndex).flags.Paralizado = 1 Then
        EstadoDat.Caption = "Estado: Paralizado"
    ElseIf UserList(tIndex).flags.Meditando = True Then
        EstadoDat.Caption = "Estado: Meditando"
    ElseIf UserList(tIndex).flags.PuedeMoverse = False Then
        EstadoDat.Caption = "Estado: Moviendose"
    End If
End If

End Sub
' [/GS]


Private Sub MSX_DblClick()
On Error Resume Next
Call FrmMensajes.msg("Mensaje", MSX.Text)
End Sub
