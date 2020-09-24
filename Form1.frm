VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5268
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   3744
   LinkTopic       =   "Form1"
   ScaleHeight     =   5268
   ScaleWidth      =   3744
   StartUpPosition =   3  'Windows Default
   Begin Project1.UserControl1 UserControl11 
      Height          =   1608
      Left            =   252
      TabIndex        =   0
      Top             =   2016
      Width           =   3120
      _ExtentX        =   5503
      _ExtentY        =   2836
      ProcessEnter    =   -1  'True
      ConstituentControl=   -1  'True
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "OK/Cancel"
      Default         =   -1  'True
      Height          =   348
      Left            =   924
      TabIndex        =   1
      Top             =   4872
      Width           =   1944
   End
   Begin VB.Frame Frame1 
      Caption         =   """Yellow"" control options"
      Height          =   1776
      Left            =   84
      TabIndex        =   7
      Top             =   84
      Width           =   3456
      Begin VB.CheckBox Check1 
         Caption         =   "Process Tab"
         Height          =   264
         Left            =   252
         TabIndex        =   2
         Top             =   336
         Width           =   3036
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Process Enter"
         Height          =   264
         Left            =   252
         TabIndex        =   3
         Top             =   672
         Width           =   3036
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Constituent control"
         Height          =   264
         Left            =   252
         TabIndex        =   5
         Top             =   1344
         Width           =   3036
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Process Escape"
         Height          =   264
         Left            =   252
         TabIndex        =   4
         Top             =   1008
         Width           =   3036
      End
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":0000
      Height          =   1032
      Left            =   252
      TabIndex        =   6
      Top             =   3864
      Width           =   3372
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
    UserControl11.ProcessTab = (Check1.Value = vbChecked)
End Sub

Private Sub Check2_Click()
    UserControl11.ProcessEnter = (Check2.Value = vbChecked)
End Sub

Private Sub Check3_Click()
    UserControl11.ProcessEscape = (Check3.Value = vbChecked)
End Sub

Private Sub Check4_Click()
    UserControl11.ConstituentControl = (Check4.Value = vbChecked)
End Sub

Private Sub Command1_Click()
    MsgBox "Button pressed!"
End Sub

Private Sub Form_Load()
    Check1.Value = Abs(UserControl11.ProcessTab)
    Check2.Value = Abs(UserControl11.ProcessEnter)
    Check3.Value = Abs(UserControl11.ProcessEscape)
    Check4.Value = Abs(UserControl11.ConstituentControl)
End Sub
