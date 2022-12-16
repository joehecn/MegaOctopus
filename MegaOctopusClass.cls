VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MegaOctopusClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function getErr Lib "MegaOctopusSelect" Alias "GetErr" (ByVal Err As Double) As Double
Private Declare Function openPort Lib "MegaOctopusSelect" Alias "OpenPort" (ByVal portnum As Double) As Double
Private Declare Function closePort Lib "MegaOctopusSelect" Alias "ClosePort" (ByVal portnum As Double) As Double
Private Declare Function versionCheck Lib "MegaOctopusSelect" Alias "VersionCheck" (ByVal portnum As Double, _
                                    ByVal pFileName As String) As Double
Private Declare Function readIDm Lib "MegaOctopusSelect" Alias "ReadIDm" (ByVal portnum As Double, ByVal str As String) As Double

Public Type RES
    code As Integer
    message As String
    version As String
    cardNo As String
End Type

'获取 Octopus DLL 版本号
Public Function fnGetVersion() As String
fnGetVersion = "v4.4.4"
End Function

'打开端口
Public Function fnOpenPort(portnum As String) As RES
    Dim r As RES
    
    r.code = 10001
    r.message = ""
    r.version = ""
    r.cardNo = ""
    
    On Error GoTo ErrorHandler
    
    Dim Ret As Integer
    
    Ret = openPort(portnum)
    
    r.code = Ret
    
    If (Ret <> 0) Then
        Dim msg As Double
        msg = getErr(Ret)
        r.message = str(msg)
    End If
    
    fnOpenPort = r

    Exit Function
   
ErrorHandler:
    r.code = Err.Number
    r.message = Err.Description

    fnOpenPort = r
End Function

'关闭端口
Public Function fnClosePort(portnum As String) As RES
    Dim r As RES
    
    r.code = 10001
    r.message = ""
    r.version = ""
    r.cardNo = ""

    On Error GoTo ErrorHandler
    
    Dim Ret As Integer
    
    Ret = closePort(portnum)
    
    r.code = Ret
    
    If (Ret <> 0) Then
        Dim msg As Double
        msg = getErr(Ret)
        r.message = str(msg)
    End If
    
    fnClosePort = r
    
    Exit Function
   
ErrorHandler:
    r.code = Err.Number
    r.message = Err.Description

    fnClosePort = r
End Function

'获取 MegaOctopusSelect 版本
Public Function fnVerChk(portnum As String) As RES
    Dim r As RES
    
    r.code = 10001
    r.message = ""
    r.version = ""
    r.cardNo = ""
    
    On Error GoTo ErrorHandler
    
    Dim sVer As String
    sVer = String(50, 0)
    
    Dim Ret As Integer
    
    Ret = versionCheck(portnum, sVer)
    
    r.code = Ret
    
    If Ret <> 0 Then
        Dim msg As Double
        msg = getErr(Ret)
        r.message = str(msg)
    Else
        r.version = Left(sVer, 17)
    End If
    
    fnVerChk = r
    
    Exit Function
   
ErrorHandler:
    r.code = Err.Number
    r.message = Err.Description

    fnVerChk = r
End Function

'获取卡号
Public Function fnReadIDm(portnum As String) As RES
    Dim r As RES
    
    r.code = 10001
    r.message = ""
    r.version = ""
    r.cardNo = ""

    On Error GoTo ErrorHandler
    
    Dim sIDm As String
    sIDm = "00 00 00 00-00 00 00 00 "
    
    Dim Ret As Integer
    Ret = readIDm(portnum, sIDm)
    
    r.code = Ret
    
    If (Ret <> 0) Then
        Dim msg As Double
        msg = getErr(Ret)
        r.message = str(msg)
    Else
        r.cardNo = sIDm
    End If
    
    fnReadIDm = r
    
    Exit Function
   
ErrorHandler:
    r.code = Err.Number
    r.message = Err.Description

    fnReadIDm = r
End Function




