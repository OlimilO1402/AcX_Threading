VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "ThreadingClient"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9735
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   9735
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnUnregister 
      Caption         =   "UnregServer"
      Height          =   375
      Left            =   8040
      TabIndex        =   5
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton BtnClearList 
      Caption         =   "Clear List"
      Height          =   375
      Left            =   6480
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton BtnSetNewIPAddresses 
      Caption         =   "New IP-address"
      Height          =   375
      Left            =   4920
      TabIndex        =   6
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
      Alignment       =   2  'Zentriert
      Height          =   345
      Left            =   135
      TabIndex        =   1
      Text            =   "192.168.178"
      Top             =   135
      Width           =   1545
   End
   Begin VB.ListBox List1 
      Height          =   6780
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   9615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IThreading

Private m_Threads       As Collection
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
    
    'Dim i As Long, n As Long: n = Num_Parse(Text1.Text)
    'If n <= 0 Then Exit Sub
    Dim tcp As String:  tcp = Text1.Text
    If Not IPBase_TryParse(tcp) Then
        MsgBox tcp & ": please give a valid tcp-address in the form: [0-255].[0-255].[0-255]"
        Exit Sub
    End If
    
    Dim i As Long, ip As String
    For i = 0 To 255 'n - 1
        ip = tcp & CStr(i)
        ThreadsAdd(MNew.MyThread(Me, ip)).Action
    Next
    
    Me.MousePointer = mp
    Exit Sub
Catch:
    MsgBox "Error : " & Err.Number & " in " & TypeName(Me) & "::" & "BtnCreateThreads_Click" & vbCrLf & Err.Description & vbCrLf & Err.LastDllError
End Sub

Function ThreadsAdd(mt As MyThread) As MyThread
    m_Threads.Add mt, mt.ip
    Set ThreadsAdd = mt
End Function

'Private Function Num_Parse(s As String) As Long
'
'    If IsNumeric(s) Then Num_Parse = CLng(s)
'
'End Function

Private Function IPBase_TryParse(s_inout As String) As Boolean
Try: On Error GoTo Catch
    Dim sa() As String: sa = Split(s_inout, ".")
    Dim i As Long, b(0 To 2) As String
    For i = 0 To Min(UBound(sa), 2)
        b(i) = CStr(CByte(sa(i)))
    Next
    s_inout = Join(b, ".") & "."
    IPBase_TryParse = True
Catch:
End Function

Function Min(V1, V2)
    If V1 < V2 Then Min = V1 Else Min = V2
End Function

Private Sub BtnStartThreads_Click()
    
    Dim mt As MyThread, v
    For Each v In m_Threads
        Set mt = v
        mt.Action
    Next
    
End Sub

Private Sub BtnSetNewIPAddresses_Click()
    
    Dim mt As MyThread, c As Long, v
    For Each v In m_Threads
        Set mt = v
        mt.ip = "192.168.2." & c
        c = c + 1
    Next
    
End Sub

Private Sub BtnClearList_Click()
    List1.Clear
End Sub

Private Sub BtnUnregister_Click()
    MRegAcX.Unregister m_ThrdSrv_libID, m_ThrdSrv_clsID, m_ThrdSrv_pfn
End Sub


' v ############################## v '    Implements IThreading    ' v ############################## v '

Private Sub IThreading_ActionStarted(IndexKey)
    Dim key As String: key = CStr(IndexKey)
    Dim mt As MyThread: Set mt = m_Threads.Item(key)
    Dim std As Date: std = mt.StartDate
    List1.AddItem "Started " & std & " IP: " & mt.ip
End Sub

Private Sub IThreading_ActionCompleted(IndexKey)
    Dim key As String: key = CStr(IndexKey)
    Dim mt As MyThread: Set mt = m_Threads.Item(key)
    Dim dur As String: dur = mt.DurationMs & " ms"
    List1.AddItem "Completed in " & dur & "; IP: " & mt.ip & "; Name: " & mt.Name
End Sub

