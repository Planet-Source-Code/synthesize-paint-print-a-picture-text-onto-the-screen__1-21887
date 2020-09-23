VERSION 5.00
Begin VB.Form frmRefresh 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "This form is used to refresh the desktop"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00404040&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblKeep 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"frmRefresh.frx":0000
      ForeColor       =   &H000000FF&
      Height          =   1200
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Visible         =   0   'False
      Width           =   4635
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmRefresh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

