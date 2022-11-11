Attribute VB_Name = "MRegAcX"
Option Explicit

Public Function CheckAndRegister(ByVal libID As String, ByVal clsID As String, ByVal pfn As String)
    
    Dim libclsid As String: libclsid = libID & "." & clsID
    CheckAndRegister = IsAcXRegistered(libclsid)
    
    If CheckAndRegister Then Exit Function
    
    Dim mess As String: mess = "Could not create instance of class: " & clsID & "." & vbCrLf & "Maybe class not registered from library: " & libID
    Dim btn As VbMsgBoxStyle: btn = vbOKOnly
    
    If IsAdmin Then
        mess = mess & vbCrLf & "Register now?"
        btn = vbYesNo
    Else
        mess = mess & vbCrLf & "Restart app and run as admin to register here!"
    End If
    
    If MsgBox(mess, btn) = vbYes Then
        
        If Not FileExists(pfn) Then
            MsgBox "File not found: " & vbCrLf & pfn
            Exit Function
        End If
        
        'OK now we try to register it
        Dim rv As Double
        rv = Shell(pfn & " /RegServer", vbNormalFocus)
        
        CheckAndRegister = IsAcXRegistered(libclsid)
        If CheckAndRegister Then
            MsgBox "RegServer successful!"
        Else
            MsgBox "Failed to register!"
        End If
    End If
    
End Function

Public Function FileExists(ByVal pfn As String) As Boolean
    FileExists = Not CBool(GetAttr(pfn) And (vbDirectory Or vbVolume))
End Function

Public Function IsAcXRegistered(ByVal ClassID As String) As Boolean
Try: On Error GoTo Catch
    Dim obj As Object: Set obj = CreateObject(ClassID)
    IsAcXRegistered = Not (obj Is Nothing)
Catch:
End Function

Public Sub Unregister(ByVal libID As String, ByVal clsID As String, pfn As String)
    Dim libclsid As String: libclsid = libID & "." & clsID
    If Not IsAcXRegistered(libclsid) Then
        MsgBox "Nothing to unregister, maybe already unregistered."
        Exit Sub
    End If
    If Not FileExists(pfn) Then
        MsgBox "File not found, unable to unregister!" & vbCrLf & pfn
        Exit Sub
    End If
    If Not IsAdmin Then
        MsgBox "Restart app and run as admin to unregister!"
        Exit Sub
    End If
    Dim rv As Double: rv = Shell(pfn & " /UnregServer", vbNormalFocus)
    If Not IsAcXRegistered(libclsid) Then
        MsgBox "UnregServer successful!"
    Else
        MsgBox "Failed to unregister!"
    End If
End Sub
