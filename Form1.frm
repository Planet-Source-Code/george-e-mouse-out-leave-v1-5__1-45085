VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H000000FF&
   Caption         =   "Form1"
   ClientHeight    =   4920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2730
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   2730
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   735
      Left            =   480
      TabIndex        =   6
      Top             =   4080
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   3480
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   240
      ScaleHeight     =   435
      ScaleWidth      =   2235
      TabIndex        =   4
      Top             =   2760
      Width           =   2295
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H000000FF&
      Caption         =   "Command4"
      Height          =   375
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H000000FF&
      Caption         =   "Command3"
      Height          =   375
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "Command2"
      Height          =   375
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Caption         =   "Command1"
      Height          =   375
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   1800
      Top             =   4080
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub Form_Load()

  Dim Control As Object

    For Each Control In Me.Controls
        If TypeOf Control Is Shape Then
        Else
            Hook Control
            DoEvents
        End If
    Next Control

End Sub


Private Sub Form_Unload(Cancel As Integer)

  Dim Control As Object

    UnHook Me
    For Each Control In Me.Controls
        If TypeOf Control Is Shape Then
        Else
            UnHook Control
            DoEvents
        End If
    Next Control

End Sub

