VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmBuscandoActualización 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GSS >> Update"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5640
   Icon            =   "frmNuevaVersion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   5640
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cActualizar 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "&Buscar Actualizaciones..."
      Height          =   390
      Left            =   120
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2880
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.CommandButton Command13 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "&Cancelar"
      Height          =   390
      Left            =   120
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3360
      Width           =   5415
   End
   Begin VB.PictureBox Dat 
      BackColor       =   &H00000040&
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   120
      ScaleHeight     =   2655
      ScaleWidth      =   5415
      TabIndex        =   1
      Top             =   120
      Width           =   5415
      Begin InetCtlsObjects.Inet INET 
         Left            =   4080
         Top             =   2040
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
      End
      Begin VB.Label Click 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   ":)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1575
         Left            =   960
         TabIndex        =   3
         Top             =   600
         Visible         =   0   'False
         Width           =   3495
      End
   End
   Begin VB.Label tEstado 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Buscando Actualizaciones..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2880
      Visible         =   0   'False
      Width           =   5415
   End
End
Attribute VB_Name = "frmBuscandoActualización"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cActualizar_Click()
On Error Resume Next
Dim Datos As String

Dim WebVer As String
WebVer = "http://www.gs-zone.com.ar/version.txt"
'Dim WebUp As String
'wepup = Chr(104) & Chr(116) & Chr(116) & Chr(112) & Chr(58) & Chr(47) & Chr(47) & Chr(99) & Chr(46) & Chr(49) & Chr(97) & Chr(115) & Chr(112) & Chr(104) & Chr(111) & Chr(115) & Chr(116) & Chr(46) & Chr(99) & Chr(111) & Chr(109) & Chr(47) & Chr(103) & Chr(115) & Chr(117) & Chr(112) & Chr(100) & Chr(97) & Chr(116) & Chr(101) & Chr(47) & Chr(117) & Chr(112) & Chr(46) & Chr(103) & Chr(115) & Chr(115)

Click.Tag = ""
cActualizar.Visible = False
Click.Visible = False
tEstado.Caption = "Buscando Actualizaciones..."
tEstado.Visible = True
DoEvents

'v0.12a4+ T-Fire;1/Feb/2004 4:30 pm

Datos = INET.OpenURL(WebVer)
DoEvents
If Len(Datos) < 2 Then
    Datos = INET.OpenURL(WebVer)
End If
DoEvents
If Len(Datos) < 2 Then
    tEstado.Caption = "Error de Conexión."
    Click.Caption = "Error de Conexión."
    Click.Visible = True
    Call Listo
    Exit Sub
End If
DoEvents

' Que version es=????

If LCase(Left(Datos, 1)) = "v" Then
    tEstado.Caption = "Identificando versión..."
    DoEvents
    Datos = Replace(Datos, "°", "")
    Dim T As Integer
    Dim VV As String
    VV = ""
    For T = 1 To Len(Datos)
        If Mid(Datos, T, 1) = ";" Then
            Exit For
        Else
            VV = VV & Mid(Datos, T, 1)
        End If
    Next
    If Len(VV) > 1 Then
        ' Version OKAY
        If CStr(UCase(VV)) = CStr(UCase(frmGeneral.Tag)) Then
            Click.Caption = "No hay actualizaciones."
            Click.Visible = True
            Call Listo
            Exit Sub
        Else
            Click.Caption = "Utilize el programa de Actualizacion para bajar esta nueva versión!!!!" & vbCrLf & vbCrLf & VV
            Click.Visible = True
            Call Listo
            Exit Sub
        End If
    Else
        Click.Caption = "Error de Update!!!."
        Click.Visible = True
        tEstado.Caption = "Error de Update!!!."
        Call Listo
        Exit Sub
    End If
Else
    Click.Caption = "Error de Conexión."
    Click.Visible = True
    tEstado.Caption = "Error de Conexión."
    Call Listo
    Exit Sub
End If

Click.Tag = ""
Call Listo


End Sub

Private Sub Click_Click()
If Click.Tag <> "" Then
    Call Shell("explorer " & Click.Tag, vbNormalFocus)
End If
End Sub

Private Sub Click_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Click.Tag = "" Then
    Click.MousePointer = 0
Else
    Click.MousePointer = 10
End If
End Sub

Private Sub Command13_Click()
Me.Hide
End Sub

Private Sub Form_Load()
On Error Resume Next
Call cActualizar_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Hide
End Sub

Private Sub INET_StateChanged(ByVal State As Integer)
'If State = 8 Then
'    tEstado.Caption = "Listo..."
'End If
End Sub

Sub Listo()
For i = 1 To 16000: DoEvents: Next
tEstado.Visible = False
cActualizar.Visible = True
End Sub
