VERSION 5.00
Begin VB.Form frmG_ValCliente 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GSS >> Validación de Cliente Propio"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6105
   Icon            =   "frmG_ValCliente.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3570
   ScaleWidth      =   6105
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FF00&
      Caption         =   "&Cerrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3120
      Width           =   5850
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "&Autorizar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2640
      Width           =   1770
   End
   Begin VB.CommandButton cmdActualizar 
      BackColor       =   &H0000FF00&
      Caption         =   "&Quitar Autorizacion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2640
      Width           =   2370
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   1980
      Left            =   1680
      Pattern         =   "*.EXE"
      TabIndex        =   2
      Top             =   120
      Width           =   4335
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   2565
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3120
      TabIndex        =   4
      Top             =   2280
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&MD5 Autorizado:"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   2280
      Width           =   1575
   End
End
Attribute VB_Name = "frmG_ValCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdActualizar_Click()
AUTORIZADO = ""
Label2.Caption = ""
Call WriteVar(App.Path & "\Opciones.ini", "ANTI-CHITS", "ClienteValido", "")
End Sub

Private Sub Command1_Click()
If Len(File1.FileNamE) > 0 Then
    AUTORIZADO = txtOffset(hexMd52Asc(MD5File(File1.Path & "\" & File1.FileNamE)), 53)
    Call WriteVar(App.Path & "\Opciones.ini", "ANTI-CHITS", "ClienteValido", Mohamed(AUTORIZADO))
    Label2.Caption = MD5String(Mohamed(AUTORIZADO))
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Dir1_Change()
On Error GoTo fallo
File1.Path = Dir1.Path
fallo:
Dir1.Refresh
File1.Refresh
End Sub

Private Sub Drive1_Change()
On Error GoTo fallo
Dir1.Path = Drive1.Drive
fallo:
Drive1.Drive = "C:"
End Sub

Private Sub Form_Load()
If Len(AUTORIZADO) > 5 Then
    Label2.Caption = MD5String(Mohamed(AUTORIZADO))
Else
    Label2.Caption = ""
End If
Me.Left = 0
Me.Top = 0
End Sub
