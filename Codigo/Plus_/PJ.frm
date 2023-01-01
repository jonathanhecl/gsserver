VERSION 5.00
Begin VB.Form PJ 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PJ"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4590
   Icon            =   "PJ.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5985
   ScaleWidth      =   4590
   Begin VB.CommandButton SEG 
      BackColor       =   &H0000FF00&
      Caption         =   "&SEGURO"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   5520
      Width           =   4335
   End
   Begin VB.Timer Vigilando 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3960
      Top             =   2400
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H0000FF00&
      Caption         =   "&BanIP"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5160
      Width           =   1935
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H0000FF00&
      Caption         =   "&Ban"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4800
      Width           =   1935
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H0000FF00&
      Caption         =   "&Echar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4440
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0000FF00&
      Caption         =   "&Encarcelar 5 minutos"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4080
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0000FF00&
      Caption         =   "Teletransportar a Lindos"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5160
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000FF00&
      Caption         =   "Teletransportar a Nix"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4800
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FF00&
      Caption         =   "Teletransportar a Banderbill"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4440
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "Teletransportar a Ullathorpe"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4080
      Width           =   2295
   End
   Begin VB.ListBox Dialogos 
      BackColor       =   &H00004000&
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
      Height          =   1035
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   4335
   End
   Begin VB.PictureBox MAP 
      BackColor       =   &H00000000&
      Height          =   1575
      Left            =   2400
      ScaleHeight     =   100
      ScaleMode       =   0  'Usuario
      ScaleWidth      =   100
      TabIndex        =   7
      Top             =   480
      Width           =   2055
      Begin VB.Shape XX 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000FF00&
         DrawMode        =   7  'Invert
         FillColor       =   &H0000FF00&
         Height          =   45
         Left            =   600
         Shape           =   1  'Square
         Top             =   480
         Width           =   60
      End
   End
   Begin VB.Label DAT 
      BackStyle       =   0  'Transparent
      Caption         =   "Label12"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   10
      Left            =   3240
      TabIndex        =   31
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label DAT 
      BackStyle       =   0  'Transparent
      Caption         =   "Label12"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   9
      Left            =   3000
      TabIndex        =   30
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label DAT 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label12"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   8
      Left            =   3120
      TabIndex        =   29
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label DAT 
      BackStyle       =   0  'Transparent
      Caption         =   "Label12"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   7
      Left            =   600
      TabIndex        =   28
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label DAT 
      BackStyle       =   0  'Transparent
      Caption         =   "Label12"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   6
      Left            =   600
      TabIndex        =   27
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label DAT 
      BackStyle       =   0  'Transparent
      Caption         =   "Label12"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   5
      Left            =   840
      TabIndex        =   26
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label DAT 
      BackStyle       =   0  'Transparent
      Caption         =   "Label12"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   4
      Left            =   840
      TabIndex        =   25
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label DAT 
      BackStyle       =   0  'Transparent
      Caption         =   "Label12"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   3
      Left            =   840
      TabIndex        =   24
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label DAT 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label12"
      ForeColor       =   &H0000FF00&
      Height          =   495
      Index           =   2
      Left            =   1080
      TabIndex        =   23
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label DAT 
      BackStyle       =   0  'Transparent
      Caption         =   "Label12"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   22
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label DAT 
      BackStyle       =   0  'Transparent
      Caption         =   "Label12"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   21
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Haciendo:"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2400
      TabIndex        =   20
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Estado:"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2400
      TabIndex        =   19
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "MP:"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "HP:"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Ubicacion:"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Clase:"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Raza:"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Genero:"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Experiencia:"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nivel:"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "PJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conectado As Boolean

Private Sub Command1_Click()
On Error Resume Next
Call WarpUserChar(Vigilando.Tag, Ullathorpe.MAP, Ullathorpe.x, Ullathorpe.y, True)
End Sub

Private Sub Command2_Click()
On Error Resume Next
Call WarpUserChar(Vigilando.Tag, Banderbill.MAP, Banderbill.x, Banderbill.y, True)
End Sub

Private Sub Command3_Click()
On Error Resume Next
Call WarpUserChar(Vigilando.Tag, Nix.MAP, Nix.x, Nix.y, True)
End Sub

Private Sub Command4_Click()
On Error Resume Next
Call WarpUserChar(Vigilando.Tag, Lindos.MAP, Lindos.x, Lindos.y, True)
End Sub

Private Sub Command5_Click()
On Error Resume Next
Call Encarcelar(Vigilando.Tag, 5, "")
Call SendData(ToAdmins, 0, 0, "||" & UserList(Vigilando.Tag).Name & " ha sido encarcelado por el Host." & FONTTYPE_VENENO)
End Sub

Private Sub Command6_Click()
On Error Resume Next
Call SendData(ToAll, 0, 0, "||El Host expulso a " & UserList(Vigilando.Tag).Name & "." & FONTTYPE_INFO)
Call CloseSocket(Vigilando.Tag)
End Sub

Private Sub Command7_Click()
On Error Resume Next
Call LogBan(Vigilando.Tag, Vigilando.Tag, "BAN DESDE EL HOST")
Call SendData(ToAdmins, 0, 0, "||El Host a expulsado y baneado a " & UserList(Vigilando.Tag).Name & "." & FONTTYPE_FIGHT)
UserList(Vigilando.Tag).flags.ban = 1
Call CloseSocket(Vigilando.Tag)
End Sub

Private Sub Command8_Click()
On Error Resume Next
' [NEW]
Dim BanIP As String
BanIP = UserList(Vigilando.Tag).IP
BanIps.Add BanIP
Call LogBan(Vigilando.Tag, Vigilando.Tag, "Ban por IP desde Nick Desde el HOST")
Call SendData(ToAll, 0, 0, "||El Host expulso y baneo a " & UserList(Vigilando.Tag).Name & "." & FONTTYPE_FIGHT)
UserList(Vigilando.Tag).flags.ban = 1
Call CloseSocket(Vigilando.Tag)
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Show
End Sub

Function Vigilar(ByVal Username As String)
Vigilando.Enabled = False
Vigilando.Tag = Username
Vigilando.Interval = 1
Vigilando.Enabled = True
Me.Show
End Function

Private Sub Option1_Click()

End Sub

Private Sub SEG_Click()
If Conectado = False Or Command1.Enabled = True Then
    Command1.Enabled = False
    Command2.Enabled = False
    Command3.Enabled = False
    Command4.Enabled = False
    Command5.Enabled = False
    Command6.Enabled = False
    Command7.Enabled = False
    Command8.Enabled = False
Else
    Command1.Enabled = True
    Command2.Enabled = True
    Command3.Enabled = True
    Command4.Enabled = True
    Command5.Enabled = True
    Command6.Enabled = True
    Command7.Enabled = True
    Command8.Enabled = True
End If
End Sub

Private Sub Vigilando_Timer()
If Vigilando.Tag <> "" Then
    Dim UserIndex As Integer
    UserIndex = Vigilando.Tag
    If UserIndex >= 0 Then
        If Conectado = False Then
            Dialogos.AddItem Time & " - CONECTADO"
        End If
        Conectado = True
        DAT(0).Caption = UserList(UserIndex).Name
        DAT(1).Caption = UserList(UserIndex).Stats.ELV
        DAT(2).Caption = UserList(UserIndex).Stats.exp & "/ " & UserList(UserIndex).Stats.ELU
        DAT(3).Caption = Num2Gen(UserList(UserIndex).genero)
        DAT(4).Caption = Num2Raza(UserList(UserIndex).raza)
        DAT(5).Caption = Num2Clase(UserList(UserIndex).clase)
        DAT(6).Caption = UserList(UserIndex).Stats.MinHP & "/" & UserList(UserIndex).Stats.MaxHP
        DAT(7).Caption = UserList(UserIndex).Stats.MinMAN & "/" & UserList(UserIndex).Stats.MaxMAN
        DAT(8).Caption = "Mapa " & UserList(UserIndex).Pos.MAP & " " & UserList(UserIndex).Pos.x & "," & UserList(UserIndex).Pos.y
        DAT(9).Caption = IIf(Criminal(UserIndex) = True, "CRIMINAL", "CIUDADANO")
        If UserList(UserIndex).flags.PuedeAtacar = False Then
            DAT(10).Caption = "ATACANDO"
        ElseIf UserList(UserIndex).flags.PuedeLanzarSpell = False Then
            DAT(10).Caption = "HECHIZANDO"
        ElseIf UserList(UserIndex).flags.PuedeMoverse = False Then
            DAT(10).Caption = "MOVIENDOSE"
        ElseIf UserList(UserIndex).flags.PuedeTrabajar = False Then
            DAT(10).Caption = "TRABAJANDO"
        End If
        MAP.Cls
        XX.Left = UserList(UserIndex).Pos.x
        XX.Top = UserList(UserIndex).Pos.y
        Vigilando.Interval = 1000
    Else
        Call SEG_Click
        Dialogos.AddItem Time & " - DESCONECTADO"
        Conectado = False
    End If
End If
End Sub
