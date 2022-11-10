Attribute VB_Name = "MNew"
Option Explicit

Public Function MyThread(ith As IThreading, ByVal Index As Long, ByVal sIP As String) As MyThread
    Set MyThread = New MyThread: MyThread.New_ ith, Index, sIP
End Function

