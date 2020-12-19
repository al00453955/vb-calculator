VERSION 4.00
Begin VB.Form Form1 
   BackColor       =   &H00808080&
   Caption         =   "Super Calculadora"
   ClientHeight    =   4710
   ClientLeft      =   1275
   ClientTop       =   1860
   ClientWidth     =   6555
   BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Height          =   5400
   Icon            =   "Calculadora.frx":0000
   Left            =   1215
   LinkTopic       =   "Form1"
   MouseIcon       =   "Calculadora.frx":030A
   ScaleHeight     =   4710
   ScaleWidth      =   6555
   Top             =   1230
   Width           =   6675
   Begin VB.TextBox Text10 
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   5400
      TabIndex        =   36
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox Text9 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   480
      TabIndex        =   35
      Top             =   4080
      Width           =   3135
   End
   Begin VB.CommandButton Command27 
      Caption         =   "MR"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   34
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton Command26 
      Caption         =   "Min"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   33
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton Command25 
      Caption         =   "M-"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   32
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton Command24 
      Caption         =   "M+"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   31
      Top             =   3480
      Width           =   615
   End
   Begin VB.TextBox Text8 
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   30
      Top             =   3840
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.TextBox Text7 
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   29
      Top             =   3240
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.TextBox Text6 
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   28
      Top             =   2640
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.TextBox Text5 
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   27
      Top             =   1800
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.CommandButton Command23 
      Caption         =   "1/x"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   26
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton Command22 
      Caption         =   "p"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Symbol"
         Size            =   18
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3960
      TabIndex        =   25
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton Command21 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Raíz"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   24
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton Command20 
      BackColor       =   &H00C0C0C0&
      Caption         =   "+/-"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   23
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Command19 
      BackColor       =   &H00C0C0C0&
      Caption         =   "!"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   22
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Borrar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   21
      Top             =   4080
      Width           =   2055
   End
   Begin VB.CommandButton Command17 
      BackColor       =   &H00C0C0C0&
      Caption         =   "^"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   20
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H00C0C0C0&
      Caption         =   "/"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   19
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H00C0C0C0&
      Caption         =   "x"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   18
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H00C0C0C0&
      Caption         =   "-"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   17
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox Text4 
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   16
      Top             =   1320
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00C0C0C0&
      Caption         =   "="
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   15
      Top             =   3480
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H000080FF&
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   6600
      TabIndex        =   14
      Top             =   840
      Visible         =   0   'False
      Width           =   3732
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   6600
      TabIndex        =   13
      Top             =   240
      Visible         =   0   'False
      Width           =   3732
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00C0C0C0&
      Caption         =   "+"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   12
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton Command11 
      Caption         =   "."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2640
      TabIndex        =   11
      Top             =   2520
      Width           =   972
   End
   Begin VB.CommandButton Command10 
      Caption         =   "0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   10
      Top             =   2520
      Width           =   2055
   End
   Begin VB.CommandButton Command9 
      Caption         =   "9"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2640
      TabIndex        =   9
      Top             =   1080
      Width           =   972
   End
   Begin VB.CommandButton Command8 
      Caption         =   "8"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1560
      TabIndex        =   8
      Top             =   1080
      Width           =   972
   End
   Begin VB.CommandButton Command7 
      Caption         =   "7"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   480
      TabIndex        =   7
      Top             =   1080
      Width           =   972
   End
   Begin VB.CommandButton Command6 
      Caption         =   "6"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2640
      TabIndex        =   6
      Top             =   1560
      Width           =   972
   End
   Begin VB.CommandButton Command5 
      Caption         =   "5"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1560
      TabIndex        =   5
      Top             =   1560
      Width           =   972
   End
   Begin VB.CommandButton Command4 
      Caption         =   "4"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   480
      TabIndex        =   4
      Top             =   1560
      Width           =   972
   End
   Begin VB.CommandButton Command3 
      Caption         =   "3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2640
      TabIndex        =   3
      Top             =   2040
      Width           =   972
   End
   Begin VB.CommandButton Command2 
      Caption         =   "2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1560
      TabIndex        =   2
      Top             =   2040
      Width           =   972
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   480
      TabIndex        =   1
      Top             =   2040
      Width           =   972
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   4815
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   480
      X2              =   3600
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Menu Opciones 
      Caption         =   "&Opciones"
      Begin VB.Menu OpcionesBorrar 
         Caption         =   "&Borrar"
         Shortcut        =   ^B
      End
      Begin VB.Menu Separador 
         Caption         =   "-"
      End
      Begin VB.Menu OpcionesSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu Creditos 
      Caption         =   "&Creditos"
      Begin VB.Menu Acercade 
         Caption         =   "Acerca de la Supercalculadora"
         Begin VB.Menu Acerca 
            Caption         =   "&Acerca de...."
            Shortcut        =   ^A
         End
         Begin VB.Menu Separador2 
            Caption         =   "-"
         End
         Begin VB.Menu Ofayra 
            Caption         =   "&Ofayra Ruiz R."
            Shortcut        =   ^O
         End
         Begin VB.Menu Isabel 
            Caption         =   "&Isabel Guerrero"
            Shortcut        =   ^I
         End
         Begin VB.Menu Gabriel 
            Caption         =   "&Gabriel Téllez M."
            Shortcut        =   ^G
         End
         Begin VB.Menu Cynthia 
            Caption         =   "&Cynthia De Pando G."
            Shortcut        =   ^C
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub Acerca_Click()
Form6.Show
End Sub

Private Sub Command1_Click()
'Esta pone el número 1 como texto, no como valor,
'debido a que no es posible hacerlo con dos o más
'dígitos si lleva un valor numérico. Así después
'la tecla de =, toma este texto como valores
'numéricos para que se puedan realizar operaciones
'con muchos números
Text1.Text = Text1.Text + "1"

End Sub


Private Sub Command10_Click()
'Esta pone el 0 como texto, no como valor,
'debido a que no es posible hacerlo con dos o más
'dígitos si lleva un valor numérico. Así después
'la tecla de =, toma este texto como valores
'numéricos para que se puedan realizar operaciones
'con muchos números

Text1.Text = Text1.Text + "0"


End Sub

Private Sub Command11_Click()
'Esta el pto decimal como texto, no como valor,
'debido a que no es posible hacerlo con dos o más
'dígitos si lleva un valor numérico. Así después
'la tecla de =, toma este texto como valores
'numéricos para que se puedan realizar operaciones
'con muchos números

Text1.Text = Text1.Text + "."


End Sub


Private Sub Command12_Click()
'Este botón asegura la operación que se
'quiera realizar, este manda el dato
'"SUMA" al text4.text, para que después
'el boton de =, cumpla las condiciones
'que allí dentro tiene

If Text4.Text = "RESTA" Or Text4.Text = "DIVI" Or Text4.Text = "MULT" Or Text4.Text = "EXPO" Then
Command13 = True
End If

Text10.Text = ""
Text10.Text = Text10.Text + "+"

Text4.Text = "SUMA"
text2.Text = Text1.Text
Text1.Text = ""
text5.Text = Val(text2.Text) + Val(text5.Text)

End Sub


Private Sub Command13_Click()
'Este comando es uno de los más importantes, ya que
'Calcula las operaciones básicas de una calculadora
'suma, resta, multiplicación, división y exponente,
'el factorial no esta aqui porque decidimos que no
'era necesario oprimir la tecla de =, mejor cuando
'nadamas le oprimes la tecla de factorial, te lo saca
'automaticamente

'estas variables se usan para la división
Text3.Text = Text1.Text
a = Val(text2.Text)
b = Val(Text3.Text)

If Text4.Text = "SUMA" Then
Text1.Text = Val(text5.Text) + Val(Text1.Text)
text2.Text = ""
text5.Text = ""
Text4.Text = ""
End If

If Text4.Text = "RESTA" Then
Text1.Text = Val(text5.Text) - Val(Text1.Text)
text2.Text = ""
text5.Text = ""
Text4.Text = ""
End If

If Text4.Text = "MULT" Then
Text8.Text = Val(Text7.Text) * Val(Text1.Text)
Text1.Text = Text8.Text
Text6.Text = "1"
Text7.Text = "1"
End If

If Text4.Text = "DIVI" Then
If b > 0 Then
Text1.Text = a / b
Else
MsgBox "División entre cero: Número Infinito"
End If
End If

If Text4.Text = "EXPO" Then
Text3.Text = Text1.Text
If Text3.Text = 0 Then
Text1.Text = 1
Else
xp = 1
base = 0
Do
base = base + 1
xp = xp * Val(text2.Text)
Loop Until base = Val(Text3.Text)
Text1.Text = xp
End If
End If
Text10.Text = ""

End Sub


Private Sub Command14_Click()
'Este botón asegura la operación que se
'quiera realizar, este manda el dato
'"RESTA" al text4.text, para que después
'el boton de =, cumpla las condiciones
'que allí dentro tiene

If Text4.Text = "SUMA" Or Text4.Text = "DIVI" Or Text4.Text = "MULT" Or Text4.Text = "EXPO" Then
Command13 = True
End If

Text10.Text = ""
Text10.Text = Text10.Text + "-"

Text4.Text = "RESTA"
text2.Text = Text1.Text
Text1.Text = ""
text5.Text = Val(text2.Text) - Val(text5.Text)
End Sub

Private Sub Command15_Click()
'Este botón asegura la operación que se
'quiera realizar, este manda el dato
'"MULT" al text4.text, para que después
'el boton de =, cumpla las condiciones
'que allí dentro tiene

If Text4.Text = "RESTA" Or Text4.Text = "DIVI" Or Text4.Text = "SUMA" Or Text4.Text = "EXPO" Then
Command13 = True
End If

Text10.Text = ""
Text10.Text = Text10.Text + "x"

Text4.Text = "MULT"
Text6.Text = Text1.Text
Text7.Text = Val(Text6.Text) * Val(Text7.Text)
Text1.Text = ""

End Sub


Private Sub Command16_Click()
'Este botón asegura la operación que se
'quiera realizar, este manda el dato
'"DIVI" al text4.text, para que después
'el boton de =, cumpla las condiciones
'que allí dentro tiene

If Text4.Text = "RESTA" Or Text4.Text = "SUMA" Or Text4.Text = "MULT" Or Text4.Text = "EXPO" Then
Command13 = True
End If

Text10.Text = ""
Text10.Text = Text10.Text + "/"

Text4.Text = "DIVI"
text2.Text = Text1.Text
Text1.Text = ""

End Sub


Private Sub Command17_Click()
'Este botón asegura la operación que se
'quiera realizar, este manda el dato
'"EXPO" al text4.text, para que después
'el boton de =, cumpla las condiciones
'que allí dentro tiene

If Text4.Text = "RESTA" Or Text4.Text = "DIVI" Or Text4.Text = "MULT" Or Text4.Text = "SUMA" Then
Command13 = True
End If

Text10.Text = ""
Text10.Text = Text10.Text + "^"

Text4.Text = "EXPO"
text2.Text = Text1.Text
Text1.Text = ""
End Sub

Private Sub Command18_Click()
'Con esta tecla, se borran los datos de todas
'las cajas de texto, esto sirve para volver
'a empezar a hacer una nueva operación, sin
'tener complicaciones con los datos antes
'ingresados
Text1.Text = novalue
text2.Text = novalue
Text3.Text = novalue
Text4.Text = novalue
text5.Text = novalue
Text6.Text = 1
Text7.Text = 1
Text10.Text = novalue

End Sub

Private Sub Command19_Click()
'Esta tecla despliega el factorial, después
'de ingresar la variable "a". Este factorial
'tiene varias restricciones, ya que si el
'número es menor a cero, te manda un msgbox
'diciendote de que no existe el factorial, o
'si el número es igual a cero, el factorial es
'igual a 1.
'Esta muy claro que al llegar a determinado número
'el valor del factorial es muy grande, y entonces
'el programa sufre un tremendo overflow ...
'y le manda al usuario una de esas chocantes msgbox.

Text10.Text = ""
a = Val(Text1.Text)


fac = 1
x = 0

If a < 0 Then
MsgBox "No existe el factorial de números menores a cero"
Text1.Text = ""
End If

If a >= 171 Then
MsgBox "Número muy grande, ¿te imaginas cuantos ceros tendría este número...?"
Text1.Text = ""
End If

If a = 0 Then
Text1.Text = 1
End If

If a > 0 And a <= 170 Then
Do
x = x + 1
fac = fac * x
Loop Until x = a
Text1.Text = fac
End If

End Sub

Private Sub Command2_Click()
'Esta pone el número 2 como texto, no como valor,
'debido a que no es posible hacerlo con dos o más
'dígitos si lleva un valor numérico. Así después
'la tecla de =, toma este texto como valores
'numéricos para que se puedan realizar operaciones
'con muchos números

Text1.Text = Text1.Text + "2"

End Sub


Private Sub Command20_Click()
'Esta tecla cambia los valores de signo
'creemos que es importante en una calculadora
'tener este tipo de teclas, ya que es vital para
'hacer una operación
f = Val(Text1.Text)

If f < 0 Then
Text1.Text = f * (-1)
Else
Text1.Text = f * (-1)
End If

End Sub

Private Sub Command21_Click()
'saca la raíz de un número, pero
'no puede haber raíces negativas
'las raíces negativas dan como resultado
'números imaginarios o complejos por eso le
'pusimos una msgbox para que cuando le pongas
'números negativos te diga que no puede hacerlo
Text10.Text = ""

r = Val(Text1.Text)

If r >= 0 Then
r = r ^ (0.5)
Else
MsgBox "Número Imaginario: No se puede sacar raíz a un número negativo"
End If

Text1.Text = r

End Sub


Private Sub Command22_Click()
'Despliega Pi en text1.text
Text1.Text = 3.141592654

Text10.Text = ""
Text10.Text = Text10.Text + "pi"

End Sub

Private Sub Command23_Click()
'Esta tecla da el Inverso de la cantidad dada en el
'text1.text y viceversa

Text10.Text = ""
Text10.Text = Text10.Text + "1/x"

a = Val(Text1.Text)
inv = 1 / a
Text1.Text = inv
End Sub

Private Sub Command24_Click()
Text9.Text = Val(Text1.Text) + Val(Text9.Text)
End Sub

Private Sub Command25_Click()
Text9.Text = Val(Text9.Text) - Val(Text1.Text)
End Sub

Private Sub Command26_Click()
Text9.Text = novalue
End Sub


Private Sub Command27_Click()
Text1.Text = Val(Text9.Text)
End Sub


Private Sub Command28_Click()
Text1.Text = Text1.Text - Text10.Text
End Sub

Private Sub Command3_Click()
'Esta pone el número 3 como texto, no como valor,
'debido a que no es posible hacerlo con dos o más
'dígitos si lleva un valor numérico. Así después
'la tecla de =, toma este texto como valores
'numéricos para que se puedan realizar operaciones
'con muchos números

Text1.Text = Text1.Text + "3"


End Sub


Private Sub Command4_Click()
'Esta pone el número 4 como texto, no como valor,
'debido a que no es posible hacerlo con dos o más
'dígitos si lleva un valor numérico. Así después
'la tecla de =, toma este texto como valores
'numéricos para que se puedan realizar operaciones
'con muchos números

Text1.Text = Text1.Text + "4"


End Sub


Private Sub Command5_Click()
'Esta pone el número 5 como texto, no como valor,
'debido a que no es posible hacerlo con dos o más
'dígitos si lleva un valor numérico. Así después
'la tecla de =, toma este texto como valores
'numéricos para que se puedan realizar operaciones
'con muchos números

Text1.Text = Text1.Text + "5"

End Sub


Private Sub Command6_Click()
'Esta pone el número 6 como texto, no como valor,
'debido a que no es posible hacerlo con dos o más
'dígitos si lleva un valor numérico. Así después
'la tecla de =, toma este texto como valores
'numéricos para que se puedan realizar operaciones
'con muchos números

Text1.Text = Text1.Text + "6"


End Sub


Private Sub Command7_Click()
'Esta pone el número 7 como texto, no como valor,
'debido a que no es posible hacerlo con dos o más
'dígitos si lleva un valor numérico. Así después
'la tecla de =, toma este texto como valores
'numéricos para que se puedan realizar operaciones
'con muchos números

Text1.Text = Text1.Text + "7"


End Sub


Private Sub Command8_Click()
'Esta pone el número 8 como texto, no como valor,
'debido a que no es posible hacerlo con dos o más
'dígitos si lleva un valor numérico. Así después
'la tecla de =, toma este texto como valores
'numéricos para que se puedan realizar operaciones
'con muchos números

Text1.Text = Text1.Text + "8"


End Sub



Private Sub Command9_Click()
'Esta pone el número 9 como texto, no como valor,
'debido a que no es posible hacerlo con dos o más
'dígitos si lleva un valor numérico. Así después
'la tecla de =, toma este texto como valores
'numéricos para que se puedan realizar operaciones
'con muchos números

Text1.Text = Text1.Text + "9"


End Sub


Private Sub Cynthia_Click()
Form5.Show
End Sub

Private Sub Form_Load()
Text6.Text = 1
Text7.Text = 1
Text8.Text = 1
End Sub

Private Sub gabriel_Click()
Form4.Show

End Sub

Private Sub Isabel_Click()
Form3.Show

End Sub

Private Sub Ofayra_Click()
Form2.Show

End Sub

Private Sub OpcionesBorrar_Click()
'Con esta parte del menu, se borran los datos de todas
'las cajas de texto, esto sirve para volver
'a empezar a hacer una nueva operación, sin
'tener complicaciones con los datos antes
'ingresados
Text1.Text = novalue
text2.Text = novalue
Text3.Text = novalue
Text4.Text = novalue
text5.Text = novalue
Text6.Text = 1
Text7.Text = 1

End Sub

Private Sub OpcionesSalir_Click()
'Esta parte del menu, sirve para dejar de correr
'el programa, el Unload Me, lo encontró Gabriel
'en el libro "El Camino Fácil a Visual Basic 4"

Unload Me
End
End Sub


Private Sub Text1_Change()
'Esta es la pantalla de la calculadora, aqui se van
'a ingresar los valores dados y aqui mismo se van a
'dar los resultados. Cuando un valor sea ingresado
'en el text1.text inmediatamente pasara al text2.text
'y al oprimir un boton de cualquier operación borrara
'automaticamente el valor ingresado sera borrado
'para dar paso al ingreso del segundo número, el cual
'pasara al text3.text inmediatamente después de
'oprimir el boton de =, y este podrá hacer lo que quiera
'con los datos del text2.text y del text3.text, y así,
'al terminar de hacer la operación requerida, pondrá
'el resultado en el text1.text
End Sub

Private Sub text2_Change()
'Aqui se va a acumular la variable "a", para que
'después la tecla de =, lo tome y haga con el y con
'la variable "b", la operación que el usuario pida

End Sub

Private Sub Text3_Change()
'Aqui se va a acumular la variable "b", para que
'después la tecla de =, lo tome y haga con el y con
'la variable "a", la operación que el usuario pida
End Sub

