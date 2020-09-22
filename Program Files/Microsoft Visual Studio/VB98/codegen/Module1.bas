Attribute VB_Name = "modPiracy"
Public Declare Function InternetGetConnectedState Lib "wininet" (lpdwFlags As Long, ByVal dwReserved As Long) As Boolean

Public Const INTERNET_CONNECTION_MODEM = 1
Public Const INTERNET_CONNECTION_LAN = 2
Public Const INTERNET_CONNECTION_PROXY = 4
Public Const INTERNET_CONNECTION_MODEM_BUSY = 8

Public seed1 As String
Public val11 As String
Public val21 As Variant
Public val31 As String

Public Function createCode(val1, val3, seed)
    On Error Resume Next
    createcode2 = seed * val1 / 4 * 200 * val3 / 5 * 137
    
    createCode = seed & "-" & val1 & "-" & Hex(createcode2) & "-" & val3
    
    seed1 = seed
    val11 = val1
    val21 = CStr(Hex(createcode2))
    val31 = val3
    
End Function

Public Function validateCode(seed, val1, val2, val3) As Boolean
    'convert val2 to value
    On Error Resume Next
    val2 = Val("&H" & val2)
    On Error Resume Next
    Temp1 = seed * val1 / 4 * 200 * val3 / 5 * 137
    
    If Temp1 = val2 Then validateCode = True Else validateCode = False
    
End Function
