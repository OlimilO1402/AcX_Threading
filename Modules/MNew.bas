Attribute VB_Name = "MNew"
Option Explicit

Public Function MyThread(ith As IThreading, ByVal sIP As String) As MyThread
    Set MyThread = New MyThread: MyThread.New_ ith, sIP
End Function

