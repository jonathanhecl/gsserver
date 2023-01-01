VERSION 5.00
Begin VB.Form EscaneadorDePJs 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GSS >> Escaneador de PJs"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10065
   Icon            =   "escanadordepjs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   10065
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check2 
      BackColor       =   &H00004000&
      Caption         =   "&De Prueba"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   8760
      TabIndex        =   56
      Top             =   4560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00004000&
      Caption         =   "&ACTIVADO"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   8760
      TabIndex        =   55
      Top             =   4200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command20 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "&Quitar Todas las Prohibiciones"
      Height          =   435
      Left            =   2880
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   5655
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command13 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "&Cerrar"
      Height          =   435
      Left            =   5400
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   6120
      Width           =   4575
   End
   Begin VB.CommandButton Command16 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "BORRAR PJ"
      Height          =   375
      Left            =   8760
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   45
      ToolTipText     =   "Banea/desbanea el PJ"
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "UN/&BAN PJ"
      Height          =   375
      Left            =   8760
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Banea/desbanea el PJ"
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "Reset &OROs"
      Height          =   375
      Left            =   8760
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Resetear el Oro en el Banco y en la Billetera"
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton Command19 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "&Quitar Prohibicion seleccionada"
      Height          =   435
      Left            =   2880
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   5250
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command18 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "&Agregar Nueva Prohibicion >>"
      Height          =   435
      Left            =   2880
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   4845
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command17 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "&Guardar Nueva Configuración"
      Height          =   435
      Left            =   2880
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   4365
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.ComboBox lstCP 
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   315
      ItemData        =   "escanadordepjs.frx":1042
      Left            =   2880
      List            =   "escanadordepjs.frx":1049
      TabIndex        =   50
      Text            =   "Prohibir Todo"
      Top             =   4080
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.ListBox CP 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   2565
      Left            =   5400
      TabIndex        =   47
      Top             =   3360
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CheckBox Admin 
      BackColor       =   &H0000FF00&
      Caption         =   "&ADMINISTRADOR"
      Height          =   435
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   6120
      Width           =   2655
   End
   Begin VB.CommandButton Command15 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "BORRARDOR de PJ's"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2880
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   6120
      Width           =   2415
   End
   Begin VB.FileListBox SECRET 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   225
      Hidden          =   -1  'True
      Left            =   8760
      TabIndex        =   43
      Top             =   3480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command12 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "Quitar ITEM a todos los PJs"
      Height          =   375
      Left            =   2880
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   4440
      Width           =   2415
   End
   Begin VB.CommandButton Command14 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "Restear los CLANES de todos"
      Height          =   255
      Left            =   2880
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   4200
      Width           =   2415
   End
   Begin VB.CommandButton Command11 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "Inventario PJ"
      Height          =   315
      Left            =   2880
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   3720
      Width           =   2415
   End
   Begin VB.CommandButton Command10 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "Inventario Banco"
      Height          =   315
      Left            =   2880
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   3360
      Width           =   2415
   End
   Begin VB.CommandButton Command9 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "Siguiente >>"
      Height          =   375
      Left            =   4200
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "Siguiente >>"
      Height          =   375
      Left            =   4200
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "Siguiente >>"
      Height          =   375
      Left            =   4200
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "Reset IT&EMs"
      Height          =   375
      Left            =   8760
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Resetear los Items en el inventario del Banco y en el PJ"
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "Buscar por I&tem"
      Height          =   375
      Left            =   3480
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "Buscar por &Mail"
      Height          =   375
      Left            =   3480
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "Buscar por &IP"
      Height          =   375
      Left            =   3480
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   4920
      Width           =   1335
   End
   Begin VB.ListBox inv 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   2565
      Left            =   5400
      TabIndex        =   27
      Top             =   3360
      Width           =   3255
   End
   Begin VB.FileListBox archivos 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   5490
      Hidden          =   -1  'True
      Left            =   120
      Pattern         =   "*.chr"
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackColor       =   &H00004000&
      Caption         =   "Acciones:"
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
      Height          =   195
      Left            =   8760
      TabIndex        =   61
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "Baneado:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7440
      TabIndex        =   60
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackColor       =   &H00004000&
      Caption         =   "Detalles del Char Seleccionado...."
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
      Height          =   195
      Left            =   4560
      TabIndex        =   59
      Top             =   120
      Width           =   2925
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H00004000&
      Caption         =   "Listado de Char's"
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
      Height          =   195
      Left            =   600
      TabIndex        =   58
      Top             =   120
      Width           =   1485
   End
   Begin VB.Label ONLINE 
      AutoSize        =   -1  'True
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "OFFLINE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   8520
      TabIndex        =   57
      Top             =   120
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Label Label15 
      BackColor       =   &H00004000&
      Caption         =   "Configuraciones:"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3480
      TabIndex        =   49
      Top             =   3840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label14 
      BackColor       =   &H00004000&
      Caption         =   "Comandos prohibidos:"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2880
      TabIndex        =   48
      Top             =   3360
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   2640
      Left            =   5385
      Top             =   3345
      Width           =   3285
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   5895
      Left            =   105
      Top             =   120
      Width           =   2685
   End
   Begin VB.Label ban 
      BackColor       =   &H00004000&
      Caption         =   "-"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   8280
      TabIndex        =   40
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Email 
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4080
      TabIndex        =   26
      Top             =   2880
      Width           =   2775
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2880
      TabIndex        =   25
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2880
      TabIndex        =   24
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label desc 
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4080
      TabIndex        =   23
      Top             =   2520
      Width           =   5295
   End
   Begin VB.Label clase 
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4080
      TabIndex        =   22
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "Clase:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2880
      TabIndex        =   21
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label genero 
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4080
      TabIndex        =   20
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "Genero:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2880
      TabIndex        =   19
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "Raza:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6240
      TabIndex        =   18
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label raza 
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   7440
      TabIndex        =   17
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label IP 
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   7440
      TabIndex        =   16
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "Ultimo IP:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6240
      TabIndex        =   15
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label mana 
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   7440
      TabIndex        =   14
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "Mana:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6240
      TabIndex        =   13
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "Vida:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2880
      TabIndex        =   12
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label vida 
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4080
      TabIndex        =   11
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label banco 
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   7440
      TabIndex        =   10
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label oro 
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4080
      TabIndex        =   9
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label exp 
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   6240
      TabIndex        =   8
      Top             =   720
      Width           =   3615
   End
   Begin VB.Label nivel 
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4080
      TabIndex        =   7
      Top             =   720
      Width           =   735
   End
   Begin VB.Label nombre 
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4080
      TabIndex        =   6
      Top             =   360
      Width           =   5775
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "Banco:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6240
      TabIndex        =   5
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "Oro:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2880
      TabIndex        =   4
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "Exp:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5400
      TabIndex        =   3
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "Nivel:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2880
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2880
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00004000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   3150
      Left            =   2865
      Top             =   120
      Width           =   7080
   End
End
Attribute VB_Name = "EscaneadorDePJs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public BuscarEMail As String
Public Ult As Integer
Public BuscarIP As String
Public BuscarItem As Integer


Private Sub Admin_Click()
If nombre = "" Or Admin.Value = 0 Then
    ' new menus
    Label14.Visible = False
    Label15.Visible = False
    lstCP.Visible = False
    CP.Visible = False
    Command17.Visible = False
    Command18.Visible = False
    Command19.Visible = False
    Command20.Visible = False
    Check2.Visible = False
    Check1.Visible = False
    ' Modo user :P
    Command1.Visible = True
    Command2.Visible = True
    Command3.Visible = True
    Command4.Visible = True
    Command5.Visible = True
    Command6.Visible = True
    Command7.Visible = True
    Command8.Visible = True
    Command9.Visible = True
    Command10.Visible = True
    Command11.Visible = True
    Command12.Visible = True
    Command14.Visible = True
    Command15.Visible = True
    Command16.Visible = True
    inv.Visible = True
    Admin.Value = 0
    Exit Sub
End If
If Admin.Value = 1 Then
    ' Modo admin
    Command1.Visible = False
    Command2.Visible = False
    Command3.Visible = False
    Command4.Visible = False
    Command5.Visible = False
    Command6.Visible = False
    Command7.Visible = False
    Command8.Visible = False
    Command9.Visible = False
    Command10.Visible = False
    Command11.Visible = False
    Command12.Visible = False
    Command14.Visible = False
    Command15.Visible = False
    Command16.Visible = False
    inv.Visible = False
    ' new menus
    Call CargarListaCPConfigurable
    Label14.Visible = True
    Label15.Visible = True
    lstCP.Visible = True
    CP.Visible = True
    CP.Clear
    Command17.Visible = True
    Command18.Visible = True
    Command19.Visible = True
    Command19.Enabled = False
    Command20.Visible = True
    Check1.Visible = True
    'Check1.Value = 0
    Check2.Visible = True
    'Check2.Value = 1
    Call CargarAdminUser
End If
End Sub

Sub CargarAdminUser()
Dim tempStr As String
Dim TempINT, i As Integer
If val(GetVar(archivos.Path & "\" & archivos.FileNamE, "ADMINISTRACION", "Activado")) = 1 Then
    Check1.Value = 1
Else
    Check1.Value = 0
End If
If val(GetVar(archivos.Path & "\" & archivos.FileNamE, "ADMINISTRACION", "EnPrueba")) = 1 Then
    Check2.Value = 1
Else
    Check2.Value = 0
End If
' Nombre de la configuracion :S si tiene
tempStr = GetVar(archivos.Path & "\" & archivos.FileNamE, "ADMINISTRACION", "Config")
If Len(tempStr) > 0 Then
    lstCP.Text = tempStr
Else
    lstCP.Text = "-(No guardada)-"
End If
' Prohibiciones
TempINT = val(GetVar(archivos.Path & "\" & archivos.FileNamE, "ADMINISTRACION", "CP"))
If TempINT > 0 Then
    For i = 1 To TempINT
        tempStr = GetVar(archivos.Path & "\" & archivos.FileNamE, "ADMINISTRACION", "CP" & i)
        If Len(tempStr) > 0 Then
            CP.AddItem tempStr
        End If
    Next
Else
    If lstCP.Text = "" Then
        lstCP.Text = "Prohibir Todo"
        CP.AddItem "/*"
        If Check1.Value = 1 Then Exit Sub
        Check1.Value = 0
        Check2.Value = 1
    End If
End If

End Sub

Sub CargarListaCPConfigurable()
Dim ConfigPro, i As Integer
Dim NamePro As String
lstCP.Clear
ConfigPro = val(GetVar(App.Path & "\Config-Priv.ini", "INIT", "MaxProhibicion"))
'lstCP.AddItem "Prohibir Todo"
For i = 1 To ConfigPro
    NamePro = GetVar(App.Path & "\Config-Priv.ini", "PROHIBICION" & i, "Nombre")
    If NamePro <> "" Then
        lstCP.AddItem NamePro
    End If
Next
End Sub

Private Sub archivos_Click()
On Error Resume Next
' ORO
If Len(archivos.FileNamE) = 0 Then Exit Sub

nombre = Left(archivos.FileNamE, Len(archivos.FileNamE) - 4)

For LoopC = 1 To LastUser
    If (UCase(UserList(LoopC).Name) = UCase(nombre)) Then
        ONLINE.Visible = True
    Else
        ONLINE.Visible = False
    End If
Next LoopC


nivel = GetVar(archivos.Path & "\" & archivos.FileNamE, "STATS", "ELV")
exp = GetVar(archivos.Path & "\" & archivos.FileNamE, "STATS", "EXP") & "/" & GetVar(archivos.Path & "\" & archivos.FileNamE, "STATS", "ELU")
oro = GetVar(archivos.Path & "\" & archivos.FileNamE, "STATS", "GLD")
banco = GetVar(archivos.Path & "\" & archivos.FileNamE, "STATS", "BANCO")
vida = GetVar(archivos.Path & "\" & archivos.FileNamE, "STATS", "MINHP") & "/" & GetVar(archivos.Path & "\" & archivos.FileNamE, "STATS", "MAXHP")
mana = GetVar(archivos.Path & "\" & archivos.FileNamE, "STATS", "MINMAN") & "/" & GetVar(archivos.Path & "\" & archivos.FileNamE, "STATS", "MAXMAN")
genero = GetVar(archivos.Path & "\" & archivos.FileNamE, "INIT", "Genero")
raza = GetVar(archivos.Path & "\" & archivos.FileNamE, "INIT", "Raza")
clase = GetVar(archivos.Path & "\" & archivos.FileNamE, "INIT", "Clase")
IP = GetVar(archivos.Path & "\" & archivos.FileNamE, "INIT", "LastIP")
desc = GetVar(archivos.Path & "\" & archivos.FileNamE, "INIT", "Desc")
Email = GetVar(archivos.Path & "\" & archivos.FileNamE, "CONTACTO", "Email")
If val(GetVar(archivos.Path & "\" & archivos.FileNamE, "FLAGS", "Ban")) = 1 Then
    ban = "PJ baneado"
Else
    ban = "PJ desbaneado"
End If

If val(GetVar(archivos.Path & "\" & archivos.FileNamE, "ADMINISTRACION", "Activado")) = 1 Then
    Admin.Value = 1
    Call Admin_Click
Else
    Admin.Value = 0
    If ONLINE.Visible = True Then
        Admin.Enabled = False
    Else
        Call Admin_Click
    End If
End If

Call Command11_Click
End Sub

Private Sub Check1_Click()
Dim LoopC As Integer
If Check1.Value = 0 Then
    ' lo desactivaron y estaba online
    If ONLINE.Visible = True Then
        For LoopC = 1 To LastUser
            If (UCase(UserList(LoopC).Name) = UCase(nombre)) Then
                ' lo desconecto
                Call SendData(ToIndex, LoopC, 0, "ERRHas sido removido de la administración.")
                Call SendData(ToIndex, LoopC, 0, "FINOK")
                Call CloseUser(LoopC)
            End If
        Next LoopC
    End If
End If
Call WriteVar(archivos.Path & "\" & archivos.FileNamE, "ADMINISTRACION", "Activado", val(Check1.Value))
End Sub

Private Sub Check2_Click()
Dim LoopC As Integer
'If Check2.Value = 0 Then
    ' Lo pongo a prueba y esta online
    If ONLINE.Visible = True Then
        For LoopC = 1 To LastUser
            If (UCase(UserList(LoopC).Name) = UCase(nombre)) Then
                ' lo desconecto
                Call SendData(ToIndex, LoopC, 0, "ERRTú configuración de Administrador ha cambiado.")
                Call SendData(ToIndex, LoopC, 0, "FINOK")
                Call CloseUser(LoopC)
            End If
        Next LoopC
    End If
'End If
Call WriteVar(archivos.Path & "\" & archivos.FileNamE, "ADMINISTRACION", "EnPrueba", val(Check2.Value))
End Sub

Private Sub Command17_Click()
Dim ConfigPro, i, j, k, sV As Integer
Dim RespMSG
If Len(lstCP.Text) > 1 And AsciiValidos(lstCP.Text) Then
     sV = 0
     For i = 0 To lstCP.ListCount - 1
        If UCase(lstCP.List(i)) = UCase(lstCP.Text) Then
            RespMSG = MsgBox(lstCP.Text & " ya existe, desea reemplazarlo?", vbYesNo + vbCritical)
            If RespMSG = vbYes Then
                ConfigPro = val(GetVar(App.Path & "\Config-Priv.ini", "INIT", "MaxProhibicion"))
                For j = 1 To ConfigPro
                    If UCase(GetVar(App.Path & "\Config-Priv.ini", "PROHIBICION" & j, "Nombre")) = UCase(lstCP.Text) Then
                        For k = 0 To CP.ListCount - 1
                            If Len(CP.List(k)) > 0 Then
                                Call WriteVar(App.Path & "\Config-Priv.ini", "PROHIBICION" & j, "CP" & (k + 1), UCase(CP.List(k)))
                                sV = sV + 1
                            End If
                        Next
                        Call WriteVar(App.Path & "\Config-Priv.ini", "PROHIBICION" & j, "CP", val(sV))
                        Call FrmMensajes.msg("Nota", "Nueva configuracion guardada.")
                        Exit Sub
                    End If
                Next
                'Call FrmMensajes.MSG("ERROR", "Error al guardar.")
                'Exit Sub
            Else
                ' No quiere sobre escribir
                Exit Sub
            End If
        End If
    Next
    ' guardar nuevo
    ConfigPro = val(GetVar(App.Path & "\Config-Priv.ini", "INIT", "MaxProhibicion"))
    ConfigPro = ConfigPro + 1
    Call WriteVar(App.Path & "\Config-Priv.ini", "INIT", "MaxProhibicion", val(ConfigPro))
    Call WriteVar(App.Path & "\Config-Priv.ini", "PROHIBICION" & ConfigPro, "Nombre", lstCP.Text)
    For j = 0 To CP.ListCount - 1
        If Len(CP.List(j)) > 0 Then
            Call WriteVar(App.Path & "\Config-Priv.ini", "PROHIBICION" & ConfigPro, "CP" & (j + 1), UCase(CP.List(j)))
            sV = sV + 1
        End If
    Next
    Call WriteVar(App.Path & "\Config-Priv.ini", "PROHIBICION" & ConfigPro, "CP", val(sV))
    Call FrmMensajes.msg("Nota", "Nueva configuracion guardada.")
    For i = 0 To lstCP.ListCount - 1
        If UCase(lstCP.List(i)) = UCase(lstCP.Text) Then Exit Sub
    Next
    lstCP.AddItem lstCP.Text
    Exit Sub
Else
    Call FrmMensajes.msg("ERROR", "Nombre de configuración invalida.")
End If
End Sub

Function Buscar(ByVal Que As String, Tipo As Integer, ComoQue As String)
On Error Resume Next
Dim ipo As String
If Tipo = -1 Then Tipo = 0
For i = Tipo To archivos.ListCount - 1
    If Que = "IP" Then
        ipo = GetVar(archivos.Path & "\" & archivos.List(i), "INIT", "LastIP")
    ElseIf Que = "EMAIL" Then
        ipo = GetVar(archivos.Path & "\" & archivos.List(i), "CONTACTO", "Email")
    ElseIf Que = "ITEM" Then
        ' En el banco
        Dim Cantidaddeitems As Integer
        Cantidaddeitems = val(GetVar(archivos.Path & "\" & archivos.List(i), "BancoInventory", "CantidadItems"))
        For j = 1 To Cantidaddeitems
            If val(ReadField(1, GetVar(archivos.Path & "\" & archivos.List(i), "BancoInventory", "Obj" & j), Asc("-"))) = val(ComoQue) Then
                MsgBox "Item: " & ComoQue & " Cantidad: " & ReadField(2, GetVar(archivos.Path & "\" & archivos.List(i), "BancoInventory", "Obj" & j), Asc("-")) & ", en el personaje " & Left(archivos.List(i), Len(archivos.List(i)) - 4)
                archivos.ListIndex = i
                Ult = i + 1
                Exit Function
            End If
        Next
        Cantidaddeitems = val(GetVar(archivos.Path & "\" & archivos.List(i), "Inventory", "CantidadItems"))
        For j = 1 To Cantidaddeitems
            If val(ReadField(1, GetVar(archivos.Path & "\" & archivos.List(i), "Inventory", "Obj" & j), Asc("-"))) = val(ComoQue) Then
                MsgBox "Item: " & ComoQue & " Cantidad: " & ReadField(2, GetVar(archivos.Path & "\" & archivos.List(i), "Inventory", "Obj" & j), Asc("-")) & ", en el personaje " & Left(archivos.List(i), Len(archivos.List(i)) - 4)
                archivos.ListIndex = i
                Ult = i + 1
                Exit Function
            End If
        Next
        GoTo salta:
    End If
    
        If Len(ipo) > 2 Then
            For j = 1 To Len(ipo)
                If Mid(ipo, j, Len(ComoQue)) = ComoQue Then
                    archivos.ListIndex = i
                    Ult = i + 1
                    Exit Function
                End If
            Next
        End If
salta:
Next
MsgBox "Fin de la Busqueda"
End Function

Private Sub Command1_Click()
On Error Resume Next
BuscarIP = InputBox("Ingrese el IP a buscar, o parte del IP")
If BuscarIP <> "" Then
    Call Buscar("IP", -1, BuscarIP)
End If
End Sub

Private Sub Command10_Click()
On Error Resume Next
inv.Clear
If Len(archivos.FileNamE) = 0 Then Exit Sub
Dim Cantidaddeitems As Integer
Cantidaddeitems = val(GetVar(archivos.Path & "\" & archivos.FileNamE, "BancoInventory", "CantidadItems"))
For j = 1 To Cantidaddeitems
    inv.AddItem j & "- (" & ReadField(1, GetVar(archivos.Path & "\" & archivos.FileNamE, "BancoInventory", "Obj" & j), Asc("-")) & ") Cantidad: " & ReadField(2, GetVar(archivos.Path & "\" & archivos.FileNamE, "BancoInventory", "Obj" & j), Asc("-")) & " = " & ObjData(ReadField(1, GetVar(archivos.Path & "\" & archivos.FileNamE, "BancoInventory", "Obj" & j), Asc("-"))).Name
Next
If Cantidaddeitems = 0 Then
    inv.AddItem "-nada-"
End If
End Sub

Private Sub Command11_Click()
On Error Resume Next
inv.Clear
If Len(archivos.FileNamE) = 0 Then Exit Sub
Cantidaddeitems = val(GetVar(archivos.Path & "\" & archivos.FileNamE, "Inventory", "CantidadItems"))
For j = 1 To Cantidaddeitems
    inv.AddItem j & "- (" & ReadField(1, GetVar(archivos.Path & "\" & archivos.FileNamE, "Inventory", "Obj" & j), Asc("-")) & ") Cantidad: " & ReadField(2, GetVar(archivos.Path & "\" & archivos.FileNamE, "Inventory", "Obj" & j), Asc("-")) & " = " & ObjData(ReadField(1, GetVar(archivos.Path & "\" & archivos.FileNamE, "Inventory", "Obj" & j), Asc("-"))).Name
Next
If Cantidaddeitems = 0 Then
    inv.AddItem "-nada-"
End If
End Sub

Private Sub Command12_Click()
On Error Resume Next
Dim ItemABorrar As Integer
ItemABorrar = InputBox("Numero del Item a quitar?")
If IsNumeric(ItemABorrar) = False Then Exit Sub
If ItemABorrar > 0 Then
    For i = 0 To archivos.ListCount - 1
        Dim Cantidaddeitems As Integer
        Cantidaddeitems = val(GetVar(archivos.Path & "\" & archivos.List(i), "BancoInventory", "CantidadItems"))
        For j = 1 To Cantidaddeitems
            If ItemABorrar = val(ReadField(1, GetVar(archivos.Path & "\" & archivos.List(i), "BancoInventory", "Obj" & j), Asc("-"))) Then
                Call WriteVar(archivos.Path & "\" & archivos.List(i), "BancoInventory", "Obj" & j, "0-0")
            End If
        Next
        Cantidaddeitems = val(GetVar(archivos.Path & "\" & archivos.List(i), "Inventory", "CantidadItems"))
        For j = 1 To Cantidaddeitems
            If ItemABorrar = val(ReadField(1, GetVar(archivos.Path & "\" & archivos.List(i), "Inventory", "Obj" & j), Asc("-"))) Then
                Call WriteVar(archivos.Path & "\" & archivos.List(i), "Inventory", "Obj" & j, "0-0-0")
            End If
        Next
    Next
End If
MsgBox "Fin"
End Sub

Private Sub Command13_Click()
On Error Resume Next
If EDPJ_Borrado.Visible = True Then Unload EDPJ_Borrado
Unload Me
End Sub

Private Sub Command14_Click()
On Error Resume Next
Dim SIoNO
SIoNO = MsgBox("¿Esta seguro que desea Resetear los Clanes de todos los Pjs?", vbCritical + vbYesNoCancel)
If SIoNO = vbYes Then
    ' Borrar clanes
    For i = 0 To archivos.ListCount - 1
        Call WriteVar(archivos.Path & "\" & archivos.List(i), "GUILD", "EsGuildLeader", "0")
        Call WriteVar(archivos.Path & "\" & archivos.List(i), "GUILD", "Echadas", "0")
        Call WriteVar(archivos.Path & "\" & archivos.List(i), "GUILD", "Solicitudes", "0")
        Call WriteVar(archivos.Path & "\" & archivos.List(i), "GUILD", "SolicitudesRechazadas", "0")
        Call WriteVar(archivos.Path & "\" & archivos.List(i), "GUILD", "VecesFueGuildLeader", "0")
        Call WriteVar(archivos.Path & "\" & archivos.List(i), "GUILD", "YaVoto", "0")
        Call WriteVar(archivos.Path & "\" & archivos.List(i), "GUILD", "GuildName", "")
        Call WriteVar(archivos.Path & "\" & archivos.List(i), "GUILD", "ClanFundado", "")
        Call WriteVar(archivos.Path & "\" & archivos.List(i), "GUILD", "ClanesParticipo", "0")
        Call WriteVar(archivos.Path & "\" & archivos.List(i), "GUILD", "GuildPts", "0")
    Next
    ' La carpeta GUILDS

'<GuildsInfo.inf>
'[INIT]
'NroGuilds = 0
    SECRET.Path = App.Path & "\Guilds\"
    For i = 0 To SECRET.ListCount - 1
        Call BorrarArchivo(SECRET.Path & "\" & SECRET.List(i))
    Next
    Call WriteVar(SECRET.Path & "\GuildsInfo.inf", "INIT", "NroGuilds", "0")
    Call MsgBox("Listo, todos los PJ ya no tienen ni pertenecen a ningun clan.", vbInformation)
End If
End Sub

Private Sub Command15_Click()
Unload Me
EDPJ_Borrado.Show
End Sub

Private Sub Command16_Click()
On Error Resume Next
If Len(archivos.FileNamE) = 0 Then Exit Sub
Call MatarPersonaje(ReadField(1, archivos.FileNamE, Asc(".")))
archivos.Refresh
End Sub



Private Sub Command18_Click()
Dim tempStr As String
Dim i As Integer
Dim Yaa As Integer
tempStr = InputBox("Que desea prohibir?" & vbCrLf & "Ejemplos:" & vbCrLf & "/p* - Prohibe todos los comandos que comienzen en /p" & vbCrLf & "*567 - Prohibe todos los comandos que terminen en 567" & vbCrLf & "/c*23 - Prohibe todos los comandos que comiencen en /c y terminen en 23")
If Len(tempStr) > 0 Then
    ' Valido?
    If Left(tempStr, 1) = " " Then GoTo Invalido
    If Right(tempStr, 1) = " " Then GoTo Invalido
    For i = 1 To Len(tempStr)
        If Mid(tempStr, i, 1) <> " " Then GoTo ok
    Next
    GoTo Invalido
ok:
    Yaa = 0
    If Left(tempStr, 1) = "*" Then Yaa = Yaa + 1
    If Right(tempStr, 1) = "*" Then Yaa = Yaa + 1
    If Len(tempStr) > 3 Then
        For i = 2 To (Len(tempStr) - 1)
            If Mid(tempStr, i, 1) = "*" Then Yaa = Yaa + 1
        Next
    End If
    If Yaa > 1 Then
        Call FrmMensajes.msg("ERROR", "Solo se puede usar un solo * por comando.")
        Exit Sub
    End If
    For i = 0 To CP.ListCount - 1
        If UCase(CP.List(i)) = UCase(tempStr) Then Exit Sub
    Next
    CP.AddItem UCase(tempStr)
    GuardarProh
Else
    If Len(tempStr) = 0 Then Exit Sub
    GoTo Invalido
End If
Exit Sub
Invalido:
Call FrmMensajes.msg("ERROR", "La prohibicion no esta bien definida o es invalida.")
End Sub

Private Sub Command19_Click()
If CP.ListIndex > -1 Then
    CP.RemoveItem (CP.ListIndex)
    CP.Refresh
    GuardarProh
    Command19.Enabled = False
End If
End Sub

Private Sub Command2_Click()
On Error Resume Next
BuscarEMail = InputBox("Ingrese el Email a buscar, o parte del Email")
If BuscarEMail <> "" Then
    Call Buscar("EMAIL", -1, BuscarEMail)
End If
End Sub

Private Sub Command20_Click()
lstCP.Text = ""
CP.Clear
Call WriteVar(archivos.Path & "\" & archivos.FileNamE, "ADMINISTRACION", "Config", "")
Call WriteVar(archivos.Path & "\" & archivos.FileNamE, "ADMINISTRACION", "CP", 0)
End Sub

Private Sub Command3_Click()
On Error Resume Next
BuscarItem = InputBox("Ingrese el numero del Item")
If IsNumeric(BuscarItem) = False Then Exit Sub
If BuscarItem <> 0 Then
    Call Buscar("ITEM", -1, str(BuscarItem))
End If
End Sub

Private Sub Command4_Click()
On Error Resume Next
If Len(archivos.FileNamE) = 0 Then Exit Sub
If val(GetVar(archivos.Path & "\" & archivos.FileNamE, "FLAGS", "Ban")) = 1 Then
    Call WriteVar(archivos.Path & "\" & archivos.FileNamE, "FLAGS", "Ban", "0")
Else
    Call WriteVar(archivos.Path & "\" & archivos.FileNamE, "FLAGS", "Ban", "1")
End If
Call archivos_Click
End Sub

Private Sub Command5_Click()
On Error Resume Next
If Len(archivos.FileNamE) = 0 Then Exit Sub
Call WriteVar(archivos.Path & "\" & archivos.FileNamE, "STATS", "GLD", "0")
Call WriteVar(archivos.Path & "\" & archivos.FileNamE, "STATS", "BANCO", "0")
Call archivos_Click
End Sub

Private Sub Command6_Click()
On Error Resume Next
If Len(archivos.FileNamE) = 0 Then Exit Sub
        Dim Cantidaddeitems As Integer
        Cantidaddeitems = val(GetVar(archivos.Path & "\" & archivos.FileNamE, "BancoInventory", "CantidadItems"))
        For j = 1 To Cantidaddeitems
            Call WriteVar(archivos.Path & "\" & archivos.FileNamE, "BancoInventory", "Obj" & j, "0-0")
        Next
        Call WriteVar(archivos.Path & "\" & archivos.FileNamE, "BancoInventory", "CantidadItems", "0")

        Cantidaddeitems = val(GetVar(archivos.Path & "\" & archivos.FileNamE, "Inventory", "CantidadItems"))
        For j = 1 To Cantidaddeitems
            Call WriteVar(archivos.Path & "\" & archivos.FileNamE, "Inventory", "Obj" & j, "0-0-0")
        Next
        Call WriteVar(archivos.Path & "\" & archivos.FileNamE, "Inventory", "CantidadItems", "0")
        Call archivos_Click
End Sub

Private Sub Command7_Click()
On Error Resume Next
If BuscarIP <> "" And Ult <> 0 Then
    Call Buscar("IP", Ult, BuscarIP)
End If
End Sub

Private Sub Command8_Click()
On Error Resume Next
If BuscarEMail <> "" And Ult <> 0 Then
    Call Buscar("EMAIL", Ult, BuscarEMail)
End If
End Sub

Private Sub Command9_Click()
On Error Resume Next
If BuscarItem <> 0 And Ult <> 0 Then
    Call Buscar("ITEM", Ult, str(BuscarItem))
End If
End Sub

Private Sub CP_Click()
If CP.ListIndex > -1 Then
    Command19.Enabled = True
Else
    Command19.Enabled = False
End If
End Sub

Private Sub CP_LostFocus()
'Command19.Enabled = False
End Sub

Private Sub Email_Click()
On Error Resume Next
BuscarEMail = InputBox("Ingrese el Email a buscar, o parte del Email", , Email.Caption)
If BuscarEMail <> "" Then
    Call Buscar("EMAIL", -1, BuscarEMail)
End If
End Sub

Private Sub Form_Load()
On Error Resume Next

archivos.Path = App.Path & "\Charfile\"
Admin.Value = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
Unload Me
End Sub

Private Sub IP_Click()
On Error Resume Next
BuscarIP = InputBox("Ingrese el IP a buscar, o parte del IP", , IP.Caption)
If BuscarIP <> "" Then
    Call Buscar("IP", -1, BuscarIP)
End If
End Sub

Private Sub lstCP_Click()
Call LeerConfig(lstCP.Text)
End Sub

Sub LeerConfig(ByVal Config As String)
Dim ConfigPro, i, j, CPx As Integer
CP.Clear
ConfigPro = val(GetVar(App.Path & "\Config-Priv.ini", "INIT", "MaxProhibicion"))
For i = 1 To ConfigPro
    If UCase(GetVar(App.Path & "\Config-Priv.ini", "PROHIBICION" & i, "Nombre")) = UCase(Config) Then
        CPx = val(GetVar(App.Path & "\Config-Priv.ini", "PROHIBICION" & i, "CP"))
        For j = 1 To CPx
            CP.AddItem UCase(GetVar(App.Path & "\Config-Priv.ini", "PROHIBICION" & i, "CP" & j))
        Next
        GuardarProh
        Exit Sub
    End If
Next

End Sub

Sub GuardarProh()
If Check2.Visible = False Then Exit Sub
Dim i, sV As Integer
sV = 0
For i = 0 To CP.ListCount - 1
    If Len(CP.List(i)) > 0 Then
        sV = sV + 1
        Call WriteVar(archivos.Path & "\" & archivos.FileNamE, "ADMINISTRACION", "CP" & sV, CP.List(i))
    End If
Next
Call WriteVar(archivos.Path & "\" & archivos.FileNamE, "ADMINISTRACION", "CP", val(sV))
Call WriteVar(archivos.Path & "\" & archivos.FileNamE, "ADMINISTRACION", "Config", lstCP.Text)
End Sub
