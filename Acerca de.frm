VERSION 4.00
Begin VB.Form Form6 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acerca de..."
   ClientHeight    =   2850
   ClientLeft      =   4485
   ClientTop       =   4650
   ClientWidth     =   4350
   Height          =   3255
   Icon            =   "Acerca de.frx":0000
   Left            =   4425
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   4350
   Top             =   4305
   Width           =   4470
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "18 de Marzo de 1999"
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   705
      Left            =   0
      Picture         =   "Acerca de.frx":030A
      Top             =   120
      Width           =   750
   End
   Begin VB.Label Label4 
      Caption         =   "Desarrollada para Patricia Chávez  por:    Ofayra Ruiz, Isabel Guerrero, Gabriel Téllez y Cynthia de Pando."
      Height          =   615
      Left            =   960
      TabIndex        =   4
      Top             =   1320
      Width           =   3375
   End
   Begin VB.Label Label3 
      Caption         =   "Copyright 1999 - ITESM CEM PS 95400 10"
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   960
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "Versión 1.0"
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Supercalculadora"
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "Form6"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form6.Hide
End Sub


