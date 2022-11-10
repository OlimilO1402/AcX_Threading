VERSION 5.00
Begin VB.Form frmHelp 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Timer tmrDelay 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   705
      Top             =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Server"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Instance As MyThread

Public Property Set Instance(ByRef Value As MyThread)
    Set m_Instance = Value
End Property

Private Sub tmrDelay_Timer()
    tmrDelay.Enabled = False
    m_Instance.DoAction
End Sub
