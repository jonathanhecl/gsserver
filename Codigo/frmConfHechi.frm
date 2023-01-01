VERSION 5.00
Begin VB.Form frmConfHechi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configurar Hechizos"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10470
   Icon            =   "frmConfHechi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   10470
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text24 
      Height          =   285
      Left            =   9720
      TabIndex        =   59
      Text            =   "0"
      Top             =   2760
      Width           =   735
   End
   Begin VB.TextBox Text23 
      Height          =   285
      Left            =   9720
      TabIndex        =   57
      Text            =   "0"
      Top             =   3120
      Width           =   735
   End
   Begin VB.TextBox Text22 
      Height          =   285
      Left            =   6840
      TabIndex        =   56
      Text            =   "Text1"
      Top             =   3840
      Width           =   1815
   End
   Begin VB.TextBox Text21 
      Height          =   285
      Left            =   6840
      TabIndex        =   55
      Text            =   "Text1"
      Top             =   3480
      Width           =   1815
   End
   Begin VB.TextBox Text20 
      Height          =   285
      Left            =   6840
      TabIndex        =   54
      Text            =   "Text1"
      Top             =   3120
      Width           =   1815
   End
   Begin VB.TextBox Text19 
      Height          =   285
      Left            =   6840
      TabIndex        =   53
      Text            =   "Text1"
      Top             =   2760
      Width           =   1815
   End
   Begin VB.TextBox Text18 
      Height          =   285
      Left            =   6840
      TabIndex        =   52
      Text            =   "Text1"
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox Text17 
      Height          =   285
      Left            =   6840
      TabIndex        =   51
      Text            =   "Text1"
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox Text16 
      Height          =   285
      Left            =   6840
      TabIndex        =   50
      Text            =   "Text1"
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox Text15 
      Height          =   285
      Left            =   6840
      TabIndex        =   49
      Text            =   "Text1"
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox Text14 
      Height          =   285
      Left            =   6840
      TabIndex        =   48
      Text            =   "Text1"
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox Text13 
      Height          =   285
      Left            =   6840
      TabIndex        =   47
      Text            =   "Text1"
      Top             =   600
      Width           =   1815
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   6840
      TabIndex        =   46
      Text            =   "Text1"
      Top             =   240
      Width           =   1815
   End
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   3480
      TabIndex        =   45
      Text            =   "Text1"
      Top             =   3840
      Width           =   1815
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   3480
      TabIndex        =   44
      Text            =   "Text1"
      Top             =   3480
      Width           =   1815
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   3480
      TabIndex        =   43
      Text            =   "Text1"
      Top             =   3120
      Width           =   1815
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   3480
      TabIndex        =   42
      Text            =   "Text1"
      Top             =   2760
      Width           =   1815
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   3480
      TabIndex        =   41
      Text            =   "Text1"
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   3480
      TabIndex        =   40
      Text            =   "Text1"
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   3480
      TabIndex        =   39
      Text            =   "Text1"
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   3480
      TabIndex        =   38
      Text            =   "Text1"
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3480
      TabIndex        =   37
      Text            =   "Text1"
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3480
      TabIndex        =   36
      Text            =   "Text1"
      Top             =   600
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3480
      TabIndex        =   35
      Text            =   "Text1"
      Top             =   240
      Width           =   1815
   End
   Begin VB.CheckBox Check7 
      Caption         =   "Revive"
      Height          =   255
      Left            =   8880
      TabIndex        =   33
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CheckBox Check6 
      Caption         =   "Envenena"
      Height          =   255
      Left            =   8880
      TabIndex        =   32
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Estupidez"
      Height          =   255
      Left            =   8880
      TabIndex        =   31
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Ciega"
      Height          =   255
      Left            =   8880
      TabIndex        =   30
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Paraliza"
      Height          =   255
      Left            =   8880
      TabIndex        =   29
      Top             =   960
      Width           =   1455
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Invisibilidad"
      Height          =   255
      Left            =   8880
      TabIndex        =   28
      Top             =   600
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Cura Veneno"
      Height          =   255
      Left            =   8880
      TabIndex        =   19
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Aplicar y &Guardar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Aplicar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Re-cargar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   4320
      Width           =   1575
   End
   Begin VB.ListBox ListadoHechi 
      Height          =   4155
      ItemData        =   "frmConfHechi.frx":000C
      Left            =   120
      List            =   "frmConfHechi.frx":000E
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Proximamente..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6840
      TabIndex        =   60
      Top             =   4320
      Width           =   3495
   End
   Begin VB.Label Label24 
      Caption         =   "Invoca:"
      Height          =   255
      Left            =   8880
      TabIndex        =   58
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label Label23 
      Caption         =   "NPC num:"
      Height          =   255
      Left            =   8880
      TabIndex        =   34
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label22 
      Caption         =   "Sube Sed:"
      Height          =   255
      Left            =   5400
      TabIndex        =   27
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label21 
      Caption         =   "Sube Hambre:"
      Height          =   255
      Left            =   5400
      TabIndex        =   26
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label20 
      Caption         =   "Sube Stamina:"
      Height          =   255
      Left            =   5400
      TabIndex        =   25
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label19 
      Caption         =   "Sube Mana:"
      Height          =   255
      Left            =   5400
      TabIndex        =   24
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label18 
      Caption         =   "Max. HP de daño:"
      Height          =   255
      Left            =   5400
      TabIndex        =   23
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label17 
      Caption         =   "Min. HP de daño:"
      Height          =   255
      Left            =   5400
      TabIndex        =   22
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label16 
      Caption         =   "Resistencia:"
      Height          =   255
      Left            =   5400
      TabIndex        =   21
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label15 
      Caption         =   "Sube HP:"
      Height          =   255
      Left            =   5400
      TabIndex        =   20
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label14 
      Caption         =   "Tipo de Objetivo:"
      Height          =   255
      Left            =   5400
      TabIndex        =   18
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label13 
      Caption         =   "Mana Requerido:"
      Height          =   255
      Left            =   5400
      TabIndex        =   17
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label12 
      Caption         =   "Min. Skill:"
      Height          =   255
      Left            =   5400
      TabIndex        =   16
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label11 
      Caption         =   "Rep. Animacion:"
      Height          =   255
      Left            =   2040
      TabIndex        =   15
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label10 
      Caption         =   "Grafico FX:"
      Height          =   255
      Left            =   2040
      TabIndex        =   14
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label9 
      Caption         =   "Nro. de Sonido:"
      Height          =   255
      Left            =   2040
      TabIndex        =   13
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "Tipo de Hechizo:"
      Height          =   255
      Left            =   2040
      TabIndex        =   12
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "Clase Exclusiva:"
      Height          =   255
      Left            =   2040
      TabIndex        =   11
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Msg. Propio:"
      Height          =   255
      Left            =   2040
      TabIndex        =   10
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Msg. al Target:"
      Height          =   255
      Left            =   2040
      TabIndex        =   9
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Msg. al Hechizero:"
      Height          =   255
      Left            =   2040
      TabIndex        =   8
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Palabras Magicas:"
      Height          =   255
      Left            =   2040
      TabIndex        =   7
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Descripcion:"
      Height          =   255
      Left            =   2040
      TabIndex        =   6
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre:"
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmConfHechi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function LoadHechiMem()
ListadoHechi.Clear
Dim i As Integer
For i = 1 To NumeroHechizos
    ListadoHechi.AddItem Hechizos(i).Nombre
Next
End Function

Private Sub ListadoHechi_Click()
On Error Resume Next
If ListadoHechi.SelCount = 1 Then
    Text1.Text = Hechizos(ListadoHechi.ListIndex + 1).Nombre
    Text2.Text = Hechizos(ListadoHechi.ListIndex + 1).Desc
    Text3.Text = Hechizos(ListadoHechi.ListIndex + 1).PalabrasMagicas
    Text4.Text = Hechizos(ListadoHechi.ListIndex + 1).HechizeroMsg
    Text5.Text = Hechizos(ListadoHechi.ListIndex + 1).TargetMsg
    Text6.Text = Hechizos(ListadoHechi.ListIndex + 1).PropioMsg
    Text7.Text = Hechizos(ListadoHechi.ListIndex + 1).ExclusivoClase
    Text8.Text = Hechizos(ListadoHechi.ListIndex + 1).Tipo
    Text9.Text = Hechizos(ListadoHechi.ListIndex + 1).WAV
    Text10.Text = Hechizos(ListadoHechi.ListIndex + 1).FXgrh
    Text11.Text = Hechizos(ListadoHechi.ListIndex + 1).loops
    Text12.Text = Hechizos(ListadoHechi.ListIndex + 1).MinSkill
    Text13.Text = Hechizos(ListadoHechi.ListIndex + 1).ManaRequerido
    Text14.Text = Hechizos(ListadoHechi.ListIndex + 1).Target
    Text15.Text = Hechizos(ListadoHechi.ListIndex + 1).Resis
    Text16.Text = Hechizos(ListadoHechi.ListIndex + 1).MinHP
    Text17.Text = Hechizos(ListadoHechi.ListIndex + 1).MaxHP
    Text18.Text = Hechizos(ListadoHechi.ListIndex + 1).SubeHP
    Text19.Text = Hechizos(ListadoHechi.ListIndex + 1).SubeMana
    Text20.Text = Hechizos(ListadoHechi.ListIndex + 1).SubeSta
    Text21.Text = Hechizos(ListadoHechi.ListIndex + 1).SubeHam
    Text22.Text = Hechizos(ListadoHechi.ListIndex + 1).SubeSed
    Text23.Text = Hechizos(ListadoHechi.ListIndex + 1).NumNpc
    Text24.Text = Hechizos(ListadoHechi.ListIndex + 1).Invoca
    Check1.Value = Hechizos(ListadoHechi.ListIndex + 1).CuraVeneno
    Check2.Value = Hechizos(ListadoHechi.ListIndex + 1).Invisibilidad
    Check3.Value = Hechizos(ListadoHechi.ListIndex + 1).Paraliza
    Check4.Value = Hechizos(ListadoHechi.ListIndex + 1).Ceguera
    Check5.Value = Hechizos(ListadoHechi.ListIndex + 1).Estupidez
    Check6.Value = Hechizos(ListadoHechi.ListIndex + 1).Envenena
    Check7.Value = Hechizos(ListadoHechi.ListIndex + 1).Revivir
End If
End Sub
