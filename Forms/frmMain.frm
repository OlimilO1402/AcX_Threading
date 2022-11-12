VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "ThreadingClient"
   ClientHeight    =   9405
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8175
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9405
   ScaleWidth      =   8175
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnUnregister 
      Caption         =   "UnregServer"
      Height          =   375
      Left            =   6480
      TabIndex        =   5
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton BtnClearList 
      Caption         =   "Clear List"
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton BtnStartThreads 
      Caption         =   "Start Threads"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton BtnCreateThreads 
      Caption         =   "Create Threads"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Rechts
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "100"
      Top             =   120
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   4740
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   8175
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IThreading

Private m_Threads As Collection
Private m_ThrdSrv_pfn   As String
Private m_ThrdSrv_libID As String
Private m_ThrdSrv_clsID As String

Private Sub Form_Load()
    
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
    Set m_Threads = New Collection
    m_ThrdSrv_libID = "PThreadServer"
    m_ThrdSrv_clsID = "MyThread"
    m_ThrdSrv_pfn = App.Path & "\" & Mid(m_ThrdSrv_libID, 2) & ".exe"
    
End Sub

Private Sub Form_Resize()
    Dim L As Single
    Dim T As Single: T = List1.Top
    Dim W As Single: W = Me.ScaleWidth
    Dim H As Single: H = Me.ScaleHeight - T
    If W > 0 And H > 0 Then List1.Move L, T, W, H
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Dim mt As MyThread, v
    For Each v In m_Threads
        
        Set mt = v
        mt.Terminate
        
    Next
    
End Sub

Private Sub BtnCreateThreads_Click()
Try: On Error GoTo Catch
    If Not CheckAndRegister(m_ThrdSrv_libID, m_ThrdSrv_clsID, m_ThrdSrv_pfn) Then Exit Sub
    
    Dim mp As MousePointerConstants: mp = Me.MousePointer
    Me.MousePointer = MousePointerConstants.vbArrowHourglass
    
    Dim i As Long, c As Long, n As Long: n = Num_Parse(Text1.Text)
    If n <= 0 Then Exit Sub
    
    For i = 0 To n - 1
        c = m_Threads.Count
        m_Threads.Add MNew.MyThread(Me, c, "192.168.178." & c)
    Next
    
    Me.MousePointer = mp
    Exit Sub
Catch:
    MsgBox "Error : " & Err.Number & " in " & TypeName(Me) & "::" & "BtnCreateThreads_Click" & vbCrLf & Err.Description & vbCrLf & Err.LastDllError
End Sub

Private Function Num_Parse(s As String) As Long
    
    If IsNumeric(s) Then Num_Parse = CLng(s)
    
End Function

Private Sub BtnStartThreads_Click()
    
    Dim mt As MyThread, v
    For Each v In m_Threads
        Set mt = v
        mt.Action
    Next
    
End Sub

Private Sub BtnClearList_Click()
    List1.Clear
End Sub

Private Sub BtnUnregister_Click()
    MRegAcX.Unregister m_ThrdSrv_libID, m_ThrdSrv_clsID, m_ThrdSrv_pfn
End Sub


' v ############################## v '    Implements IThreading    ' v ############################## v '

Private Sub IThreading_ActionStarted(ByVal Index As Long)
    Dim mt As MyThread: Set mt = m_Threads.Item(Index + 1)
    Dim std As Date: std = mt.StartDate
    List1.AddItem "Started " & std & " " & Index
End Sub

Private Sub IThreading_ActionCompleted(ByVal Index As Long)
    Dim mt As MyThread: Set mt = m_Threads.Item(Index + 1)
    Dim dur As String: dur = mt.DurationMs & " ms"
    List1.AddItem "Completed in " & dur & "; Index: " & Index & "; IP: " & mt.IP & "; Name: " & mt.Name
End Sub

