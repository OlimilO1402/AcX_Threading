VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MyThread"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliSeconds As Long)
Private mIThreading As IThreading
Private mFrmHelp    As frmHelp
Private mIndex      As Long
Private mStartDate  As Date
Private mStartTime  As Single
Private mEndTime    As Single
Private mIP         As String
Private mName       As String
Private mErr        As String

Private Const WSADescription_Len As Long = 256
Private Const WSASYS_Status_Len  As Long = 128

Private Type WSADATA
   nVersion         As Integer
   nHighVersion     As Integer
   yDescription(0 To WSADescription_Len)   As Byte
   ySystemStatus(0 To WSASYS_Status_Len)   As Byte
   nMaxSockets      As Integer
   nMaxUdp          As Integer
   lVvenderInfo   As Long
End Type

Private Type sockaddr_in
    sin_family  As Integer
    sin_port    As Integer
    sin_addr    As Long
    sin_zero(0 To 8)    As Byte
End Type

Private Declare Function WSAGetLastError Lib "Ws2_32" () As Long

Private Declare Function WSAStartup2 Lib "Ws2_32" Alias "WSAStartup" (ByVal VersionReq As Long, WSADataReturn As WSADATA) As Long
Private Declare Function WSACleanup2 Lib "Ws2_32" Alias "WSACleanup" () As Long
Private Declare Function inet_addr2 Lib "Ws2_32" Alias "inet_addr" (ByVal s As String) As Long
Private Declare Function getnameinfo Lib "Ws2_32" (ByRef sa As sockaddr_in, ByVal salen As Long, ByVal Host As String, ByVal hostlen As Long, ByVal serv As String, ByVal servlen As Long, ByVal flags As Long) As Long
Private Declare Sub ZeroMemory Lib "KERNEL32" Alias "RtlZeroMemory" (dest As Any, ByVal numBytes As Long)

Private Sub Class_Initialize()
    Set mFrmHelp = New frmHelp
    Load mFrmHelp
    'Load frmHelp
End Sub
Private Sub Class_Terminate()
    Unload mFrmHelp
End Sub

Public Sub New_(ith As IThreading, ByVal aIndex As Long, ByVal sIP As String)
    Set mIThreading = ith: mIndex = aIndex: mIP = sIP
End Sub

Public Sub Terminate()
    Class_Terminate
End Sub

Public Sub Action()
    Set mFrmHelp.Instance = Me
    mFrmHelp.tmrDelay.Enabled = True
End Sub

Public Property Get StartDate() As Date
    StartDate = mStartDate
End Property

Public Property Get MyError() As String
    MyError = mErr
End Property

Public Property Get IP() As String
    IP = mIP
End Property
Public Property Let IP(Value As String)
    mIP = Value
End Property

Public Property Get Name() As String
    Name = mName
End Property

Public Property Get DurationMs() As Single
    DurationMs = (mEndTime - mStartTime) * 1000
End Property

Friend Sub DoAction()
Try: On Error GoTo Catch
    mStartDate = Now
    mStartTime = Timer
    mIThreading.ActionStarted mIndex
    
    'Dim z As Long, v As Double
    'For z = 1 To 10 ^ 7
    '    v = Rnd(1)
    'Next z
    'Sleep 1000
    mName = GetNameFromIP(mIP)
    
    mEndTime = Timer
    mIThreading.ActionCompleted mIndex
    Exit Sub
Catch:
    MsgBox TypeName(Me) & "::DoAction" & mErr & " " & Err.Number & " " & Err.Description & vbCrLf & WSAGetLastError
End Sub

 
Private Function GetNameFromIP(ByVal strIP As String) As String
Try: On Error GoTo Catch
    Const WS_VERSION_REQD2          As Long = &H202
    Const SOCKET_ERROR              As Long = -1
    Const IP_SUCCESS                As Long = 0&
    Const WSADescription_Len        As Long = 255
    Const WSASYS_Status_Len         As Long = 127
    Const AF_INET                   As Integer = 2

    Dim llIP            As Long
    Dim ltWSAD          As WSADATA
    Dim ltSockAddr_in   As sockaddr_in
    Dim lsHostname      As String
    Dim llReturnCode    As Long
 
    If WSAStartup2(WS_VERSION_REQD2, ltWSAD) = IP_SUCCESS Then
        
        llIP = inet_addr2(strIP)
        
        If llIP <> SOCKET_ERROR Then
            ZeroMemory ltSockAddr_in, CLng(Len(ltSockAddr_in))
            ' structure
            ltSockAddr_in.sin_family = AF_INET
            ltSockAddr_in.sin_addr = llIP
            lsHostname = String$(1025, 0)
            
            'Debug.Print Timer
            llReturnCode = getnameinfo(ltSockAddr_in, Len(ltSockAddr_in), lsHostname, Len(lsHostname), 0, 0, 0)
            'Debug.Print "end: " & Timer
            
            GetNameFromIP = lsHostname
            Exit Function
        Else
            mErr = "Invalid IP"
        End If
        WSACleanup2
    Else
        mErr = "Error Startup"
    End If
    Exit Function
Catch:
    MsgBox TypeName(Me) & "::GetNameFromIP " & mErr & " " & Err.Number & " " & Err.Description & vbCrLf & WSAGetLastError
End Function

