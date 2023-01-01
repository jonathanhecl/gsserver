VERSION 5.00
Begin VB.Form frmConfVar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configurar Opciones"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6165
   Icon            =   "frmConfVar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   6165
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text17 
      Height          =   285
      Left            =   5520
      MaxLength       =   2
      TabIndex        =   41
      Text            =   "0"
      Top             =   3600
      Width           =   375
   End
   Begin VB.TextBox Text16 
      Height          =   285
      Left            =   4080
      MaxLength       =   2
      TabIndex        =   40
      Text            =   "0"
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Aplicar cambios"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   35
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Volver a cargar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   36
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cerrar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   34
      Top             =   4080
      Width           =   1455
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   5400
      MaxLength       =   3
      TabIndex        =   27
      Text            =   "0"
      Top             =   3000
      Width           =   495
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   3960
      MaxLength       =   3
      TabIndex        =   26
      Text            =   "0"
      Top             =   3000
      Width           =   495
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   5160
      MaxLength       =   9
      TabIndex        =   22
      Text            =   "0"
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   5160
      MaxLength       =   9
      TabIndex        =   20
      Text            =   "0"
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   5160
      MaxLength       =   9
      TabIndex        =   18
      Text            =   "0"
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   5160
      MaxLength       =   9
      TabIndex        =   16
      Text            =   "0"
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   4200
      MaxLength       =   9
      TabIndex        =   14
      Text            =   "0"
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   4200
      MaxLength       =   16
      TabIndex        =   12
      Text            =   "0"
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   4200
      MaxLength       =   9
      TabIndex        =   10
      Text            =   "0"
      Top             =   120
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Facciones"
      Height          =   2655
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   2895
      Begin VB.TextBox Text15 
         Height          =   285
         Left            =   1680
         MaxLength       =   9
         TabIndex        =   33
         Text            =   "0"
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox Text14 
         Height          =   285
         Left            =   1680
         MaxLength       =   9
         TabIndex        =   32
         Text            =   "0"
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox Text13 
         Height          =   285
         Left            =   1680
         MaxLength       =   9
         TabIndex        =   31
         Text            =   "0"
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   1680
         MaxLength       =   9
         TabIndex        =   30
         Text            =   "0"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   1680
         MaxLength       =   16
         TabIndex        =   29
         Text            =   "0"
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   1680
         MaxLength       =   16
         TabIndex        =   28
         Text            =   "0"
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Recompensa Caos:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Recompensa Real:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Para Ejercito Caos:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Para Armada Real:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Exp. Recompensa:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Exp. Al Enlistar:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CheckBox Res 
      Caption         =   "Utilizar Resto en los Niveles ALERTA: No funciona bien!!"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2415
   End
   Begin VB.CheckBox Pub 
      Caption         =   "Bloquear Publicidades"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.Line Line11 
      X1              =   0
      X2              =   6120
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line10 
      X1              =   0
      X2              =   0
      Y1              =   3960
      Y2              =   0
   End
   Begin VB.Line Line9 
      X1              =   6120
      X2              =   6120
      Y1              =   3960
      Y2              =   0
   End
   Begin VB.Line Line8 
      X1              =   3100
      X2              =   3100
      Y1              =   3960
      Y2              =   1080
   End
   Begin VB.Line Line7 
      X1              =   2640
      X2              =   3120
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line6 
      X1              =   120
      X2              =   2640
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line5 
      X1              =   3120
      X2              =   6120
      Y1              =   3345
      Y2              =   3345
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   6120
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line3 
      X1              =   3120
      X2              =   6120
      Y1              =   2730
      Y2              =   2730
   End
   Begin VB.Line Line2 
      X1              =   3120
      X2              =   6120
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line1 
      X1              =   2640
      X2              =   2640
      Y1              =   0
      Y2              =   1080
   End
   Begin VB.Label Label19 
      Caption         =   "Max. Dados:"
      Height          =   255
      Left            =   4560
      TabIndex        =   39
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label18 
      Caption         =   "Min. Dados:"
      Height          =   255
      Left            =   3150
      TabIndex        =   38
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label17 
      Caption         =   "Al crear un nuevo PJ..."
      Height          =   255
      Left            =   3150
      TabIndex        =   37
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label16 
      Caption         =   "MaxSkills:"
      Height          =   255
      Left            =   4560
      TabIndex        =   25
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label15 
      Caption         =   "MinSkills:"
      Height          =   255
      Left            =   3150
      TabIndex        =   24
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label14 
      Caption         =   "Al subir de lvl..."
      Height          =   255
      Left            =   3150
      TabIndex        =   23
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label Label13 
      Caption         =   "Maximo de ST:"
      Height          =   255
      Left            =   3840
      TabIndex        =   21
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label12 
      Caption         =   "Maximo de MP:"
      Height          =   255
      Left            =   3840
      TabIndex        =   19
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label11 
      Caption         =   "Maximo de HP:"
      Height          =   255
      Left            =   3840
      TabIndex        =   17
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label10 
      Caption         =   "Nivel Maximo:"
      Height          =   255
      Left            =   3840
      TabIndex        =   15
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label9 
      Caption         =   "Maximo de Oro:"
      Height          =   255
      Left            =   2880
      TabIndex        =   13
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "Maximo de Exp:"
      Height          =   255
      Left            =   2880
      TabIndex        =   11
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "Maximo de Obj. en el Inventario:"
      Height          =   495
      Left            =   2760
      TabIndex        =   9
      Top             =   20
      Width           =   1335
   End
End
Attribute VB_Name = "frmConfVar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
If Not Numeric(Text1.Text) Or Not Numeric(Text2.Text) Or Not Numeric(Text3.Text) Or Not Numeric(Text4.Text) Or Not Numeric(Text5.Text) Or Not Numeric(Text6.Text) Or Not Numeric(Text7.Text) Or Not Numeric(Text8.Text) Or Not Numeric(Text9.Text) Or Not Numeric(Text10.Text) Or Not Numeric(Text11.Text) Or Not Numeric(Text12.Text) Or Not Numeric(Text13.Text) Or Not Numeric(Text14.Text) Or Not Numeric(Text15.Text) Or Not Numeric(Text16.Text) Or Not Numeric(Text17.Text) Then
    MsgBox "Alguno de los campos ingresados, no contiene un valor numerico."
    Exit Sub
End If
Call WriteVar(IniPath & "Opciones.ini", "OPCIONES", "BloqPublicidad", Pub.Value)
Call WriteVar(IniPath & "Opciones.ini", "OPCIONES", "UsarResto", Pub.Value)
Call WriteVar(IniPath & "Opciones.ini", "OPCIONES", "ParaCaos", Text13.Text)
Call WriteVar(IniPath & "Opciones.ini", "OPCIONES", "ParaArmada", Text12.Text)
Call WriteVar(IniPath & "Opciones.ini", "OPCIONES", "RecompensaArmada", Text14.Text)
Call WriteVar(IniPath & "Opciones.ini", "OPCIONES", "RecompensaCaos", Text15.Text)
Call WriteVar(IniPath & "Opciones.ini", "OPCIONES", "EnlistarExp", Text10.Text)
Call WriteVar(IniPath & "Opciones.ini", "OPCIONES", "RecompensaExp", Text11.Text)
Call WriteVar(IniPath & "Opciones.ini", "OPCIONES", "MaxObjInventario", Text1.Text)
Call WriteVar(IniPath & "Opciones.ini", "OPCIONES", "MaxEXP", Text2.Text)
Call WriteVar(IniPath & "Opciones.ini", "OPCIONES", "MaxORO", Text3.Text)
Call WriteVar(IniPath & "Opciones.ini", "OPCIONES", "MaxLVL", Text4.Text)
Call WriteVar(IniPath & "Opciones.ini", "OPCIONES", "MaxHP", Text5.Text)
Call WriteVar(IniPath & "Opciones.ini", "OPCIONES", "MaxMAN", Text6.Text)
Call WriteVar(IniPath & "Opciones.ini", "OPCIONES", "MaxST", Text7.Text)
Call WriteVar(IniPath & "Opciones.ini", "OPCIONES", "MinSKILL", Text8.Text)
Call WriteVar(IniPath & "Opciones.ini", "OPCIONES", "MaxSKILL", Text9.Text)
Call WriteVar(IniPath & "Opciones.ini", "OPCIONES", "MinAtrib", Text16.Text)
Call WriteVar(IniPath & "Opciones.ini", "OPCIONES", "MaxAtrib", Text17.Text)
LoadOpcsINI
Unload Me
End Sub

Private Sub Command3_Click()
On Error Resume Next
DoEvents
If Publicidad = True Then
    Pub.Value = 1
Else
    Pub.Value = 0
End If
If UsarResto = True Then
    Res.Value = 1
Else
    Res.Value = 0
End If
Text10.Text = ExpAlUnirse
Text11.Text = ExpX100
Text12.Text = ParaArmada
Text13.Text = ParaCaos
Text14.Text = RecompensaXArmada
Text15.Text = RecompensaXCaos
Text1.Text = MAX_INVENTORY_OBJS
Text2.Text = MAXEXP
Text3.Text = MAXORO
Text4.Text = STAT_MAXELV
Text5.Text = STAT_MAXHP
Text6.Text = STAT_MAXMAN
Text7.Text = STAT_MAXSTA
Text8.Text = MINSKILL_G
Text9.Text = MAXSKILL_G
Text16.Text = MINATTRB
Text17.Text = MAXATTRB
DoEvents
End Sub

Private Sub Form_Load()
Call Command3_Click
End Sub

Private Sub Text10_LostFocus()
If Not Numeric(Text10.Text) Then
    MsgBox "Debe ingresar un valor numerico!"
    Text10.SetFocus
End If
End Sub

Private Sub Text11_LostFocus()
If Not Numeric(Text11.Text) Then
    MsgBox "Debe ingresar un valor numerico!"
    Text11.SetFocus
End If
End Sub

Private Sub Text15_LostFocus()
If Not Numeric(Text15.Text) Then
    MsgBox "Debe ingresar un valor numerico!"
    Text15.SetFocus
End If
End Sub

Private Sub Text16_LostFocus()
If Not Numeric(Text16.Text) Then
    MsgBox "Debe ingresar un valor numerico!"
    Text16.SetFocus
End If
End Sub

Private Sub Text17_LostFocus()
If Not Numeric(Text17.Text) Then
    MsgBox "Debe ingresar un valor numerico!"
    Text17.SetFocus
End If
End Sub

Private Sub Text9_LostFocus()
If Not Numeric(Text9.Text) Then
    MsgBox "Debe ingresar un valor numerico!"
    Text9.SetFocus
End If
End Sub
Private Sub Text8_LostFocus()
If Not Numeric(Text8.Text) Then
    MsgBox "Debe ingresar un valor numerico!"
    Text8.SetFocus
End If
End Sub
Private Sub Text7_LostFocus()
If Not Numeric(Text7.Text) Then
    MsgBox "Debe ingresar un valor numerico!"
    Text7.SetFocus
End If
End Sub
Private Sub Text6_LostFocus()
If Not Numeric(Text6.Text) Then
    MsgBox "Debe ingresar un valor numerico!"
    Text6.SetFocus
End If
End Sub
Private Sub Text5_LostFocus()
If Not Numeric(Text5.Text) Then
    MsgBox "Debe ingresar un valor numerico!"
    Text5.SetFocus
End If
End Sub
Private Sub Text4_LostFocus()
If Not Numeric(Text4.Text) Then
    MsgBox "Debe ingresar un valor numerico!"
    Text4.SetFocus
End If
End Sub
Private Sub Text3_LostFocus()
If Not Numeric(Text3.Text) Then
    MsgBox "Debe ingresar un valor numerico!"
    Text3.SetFocus
End If
End Sub
Private Sub Text2_LostFocus()
If Not Numeric(Text2.Text) Then
    MsgBox "Debe ingresar un valor numerico!"
    Text2.SetFocus
End If
End Sub
Private Sub Text1_LostFocus()
If Not Numeric(Text1.Text) Then
    MsgBox "Debe ingresar un valor numerico!"
    Text1.SetFocus
End If
End Sub
Private Sub Text12_LostFocus()
If Not Numeric(Text12.Text) Then
    MsgBox "Debe ingresar un valor numerico!"
    Text12.SetFocus
End If
End Sub
Private Sub Text13_LostFocus()
If Not Numeric(Text13.Text) Then
    MsgBox "Debe ingresar un valor numerico!"
    Text13.SetFocus
End If
End Sub
Private Sub Text14_LostFocus()
If Not Numeric(Text14.Text) Then
    MsgBox "Debe ingresar un valor numerico!"
    Text14.SetFocus
End If
End Sub
