VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "KTK License Registration"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7065
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picDone 
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   120
      ScaleHeight     =   3615
      ScaleWidth      =   6855
      TabIndex        =   22
      Top             =   1320
      Visible         =   0   'False
      Width           =   6855
      Begin VB.Label Label14 
         Caption         =   $"Form1.frx":000C
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   120
         TabIndex        =   25
         Top             =   1560
         Width           =   6615
      End
      Begin VB.Label Label13 
         Caption         =   $"Form1.frx":00B6
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   24
         Top             =   480
         Width           =   6735
      End
      Begin VB.Label Label12 
         Caption         =   "Thankyou"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   0
         TabIndex        =   23
         Top             =   120
         Width           =   6615
      End
   End
   Begin VB.PictureBox picWait 
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   120
      ScaleHeight     =   3615
      ScaleWidth      =   6855
      TabIndex        =   19
      Top             =   1320
      Visible         =   0   'False
      Width           =   6855
      Begin VB.Label Label11 
         Caption         =   "License Registration is now linking this product to your computer. This process may take a few minutes to complete"
         Height          =   855
         Left            =   1080
         TabIndex        =   21
         Top             =   1560
         Width           =   5775
      End
      Begin VB.Label Label10 
         Caption         =   "Please Wait..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   20
         Top             =   1200
         Width           =   2055
      End
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1200
      TabIndex        =   16
      Top             =   3840
      Width           =   5775
   End
   Begin VB.TextBox txtCode 
      Height          =   285
      Left            =   1200
      TabIndex        =   14
      Top             =   3480
      Width           =   5775
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   4680
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      RequestTimeout  =   30
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next >"
      Height          =   375
      Left            =   5040
      TabIndex        =   13
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6120
      TabIndex        =   12
      Top             =   5160
      Width           =   855
   End
   Begin VB.TextBox txtVal3 
      Height          =   285
      Left            =   4005
      TabIndex        =   11
      Top             =   5160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtVal2 
      Height          =   285
      Left            =   2445
      TabIndex        =   10
      Top             =   5160
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtVal1 
      Height          =   285
      Left            =   1485
      TabIndex        =   9
      Top             =   5160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtSeed 
      Height          =   285
      Left            =   405
      TabIndex        =   8
      Top             =   5160
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   6360
      Picture         =   "Form1.frx":0172
      ScaleHeight     =   525
      ScaleWidth      =   525
      TabIndex        =   2
      Top             =   360
      Width           =   525
   End
   Begin VB.Label lblHardwareFingerPrint 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label10"
      Height          =   255
      Left            =   0
      TabIndex        =   18
      Top             =   960
      Width           =   6975
   End
   Begin VB.Label Label9 
      Caption         =   "NOTE: The serial code will be provided with the software. If you have not received a serial code you should contact KTK"
      Height          =   495
      Left            =   240
      TabIndex        =   17
      Top             =   4440
      Width           =   6735
   End
   Begin VB.Label Label8 
      Caption         =   "Name:"
      Height          =   255
      Left            =   600
      TabIndex        =   15
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   $"Form1.frx":1078
      Height          =   855
      Left            =   240
      TabIndex        =   7
      Top             =   2400
      Width           =   6735
   End
   Begin VB.Label Label6 
      Caption         =   "How Do Serial Codes Work?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2160
      Width           =   6735
   End
   Begin VB.Label Label5 
      Caption         =   "We use serial codes to verify that the software has not been obtained illegally."
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Width           =   6735
   End
   Begin VB.Label Label4 
      Caption         =   "Why Use Serial Codes?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   6615
   End
   Begin VB.Label Label3 
      Caption         =   "Serial Code:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Serial Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   6495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "KTK License Registration Wizard"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6855
   End
   Begin VB.Image Image2 
      Height          =   30
      Left            =   -3840
      Picture         =   "Form1.frx":11BF
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   10935
   End
   Begin VB.Image Image1 
      Height          =   30
      Left            =   -2040
      Picture         =   "Form1.frx":1239
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   9135
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   2295
      Left            =   -1920
      Top             =   -1080
      Width           =   9495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const productid = "File_Explorer_3"
Const id = "we"

Const EncryptName = "Putencyptstringhere "

Private Declare Function GetVolumeInformation Lib "kernel32.dll" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Integer, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

'String to hold Registry Computer Name
Public SysInfoPath As String


' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

Const gREGKEYSYSINFOLOC = "System\CurrentControlSet\Control\ComputerName\ComputerName"
Const gREGKEYSYSINFO = "System\CurrentControlSet\Control\ComputerName\ComputerName"
Const gREGVALSYSINFO = "ComputerName"
Const RegKey = "Reg"
Public Register As String

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

'Put your project name here
'This is an entry in the registry that is created
Const RegPath = "SOFTWARE\KTK File Explorer 3.00"

Dim header As String
Dim licenseNo As String
Dim licenseHolder As String
Dim hardwareFingerPrint As String
Dim evalDays As String
Dim dateRegistered As String
Dim dateLastUsed As String
Dim footer As String


Private Sub cmdCancel_Click()
    End
End Sub

Private Sub cmdNext_Click()
    On Error Resume Next
    picWait.Visible = True
    cmdNext.Visible = False: cmdCancel.Visible = False
    'we must split up the code
    
    'get the location of the '-'
    firstdash = InStr(1, txtCode, "-") - 1
    txtSeed = Left(txtCode, firstdash)
    
    'get the location of the 2nd '-'
    seconddash = InStr(firstdash + 2, txtCode, "-")
    txtVal1 = Mid(txtCode, firstdash + 2, seconddash - firstdash - 2)
    
    'get the location of the 3rd '-'
    thirddash = InStr(seconddash + 1, txtCode, "-")
    txtVal2 = Mid(txtCode, seconddash + 1, thirddash - seconddash - 1)
    
    'get the location of the 3rd '-'
    fourthdash = Len(txtSeed) + Len(txtVal1) + Len(txtVal2)
    txtVal3 = Mid(txtCode, fourthdash + 4)
    On Error GoTo 0
    'Exit Sub
    
    'MsgBox fingerPrint
    
    If txtSeed = "" Or txtVal1 = "" Or txtVal2 = "" Or txtVal3 = "" Then
        MsgBox "You must enter the serial code to continue!", vbCritical, "KTK License Registration Wizard"
        picWait.Visible = False
        cmdNext.Visible = True: cmdCancel.Visible = True
        Exit Sub
    ElseIf Not (IsNumeric(txtSeed.Text)) Or Not (IsNumeric(txtVal1.Text)) Or Not (IsNumeric(txtVal3.Text)) Then
        MsgBox "The serial code you have entered is invalid. Try again!", vbCritical, "KTK License Registration Wizard"
        picWait.Visible = False
        cmdNext.Visible = True: cmdCancel.Visible = True
        Exit Sub
    End If
    
    If validateCode(txtSeed, txtVal1, txtVal2, txtVal3) = False Then
        MsgBox "The serial code you have entered is invalid. Try again!", vbCritical, "KTK License Registration Wizard"
        picWait.Visible = False
        cmdNext.Visible = True: cmdCancel.Visible = True
        Exit Sub
    End If
    
    'the code is correct.
    
    'check if we are connected to the Internet first
    If InternetGetConnectedState(flags, 0) = False Then
        MsgBox "The following error was encountered while performing online lookup:" & Chr(13) & Chr(13) & "Active Internet connection not found!", vbCritical, "KTK License Registration Wizard"
        picWait.Visible = False
        cmdNext.Visible = True: cmdCancel.Visible = True
        Exit Sub
    End If
    
    'MsgBox InternetGetConnectedState(flags, 0)
    
    'check to see if this serial code has already been issued
    
    f = FreeFile
    
    Open App.Path & "\test.dat" For Output As #f
    'Debug.Print "members.lycos.co.uk/ktkpiracy/" & productid & "/" & txtSeed & "_" & txtVal1 & "_" & txtVal2 & "_" & txtVal3 & ".txt"
    d = Inet1.OpenURL("members.lycos.co.uk/ktkpiracy/" & productid & "/" & txtSeed & "-" & txtVal1 & "-" & txtVal2 & "-" & txtVal3 & ".txt")
    Print #f, d
    Close #f
    
    'MsgBox d
    
    extractData App.Path & "\test.dat"
    
    'check to see if this license has been registered
    
    If header = "KTK" Then
        If licenseHolder = "" Then
            'not registered
            s = 0
        ElseIf fingerPrint = hardwareFingerPrint Then
            s = 0 'this user has already this product
        ElseIf Not (fingerPrint = hardwareFingerPrint) Then
            'we must do a day lookup. The computer can be changed once
            'every 30 days!
            If DateDiff("d", CDate(dateLastUsed), Date, vbUseSystemDayOfWeek, vbUseSystem) > 30 Then
                MsgBox "This product has already been registered on another computer!" & Chr(13) & "License Registration restricts the usage of licenses already registered on another computer. However, this license has been inactive for over 30 days so you can re-register it.", vbInformation, "KTK License Registration Wizard"
                s = 0
            Else
                s = 1
                MsgBox "This product has already been registered on another computer!" & Chr(13) & "License Registration restricts the usage of licenses already registered on another computer. The license you are trying to use must be inactive for atleast 30 days before it can be re-registered." & Chr(13) & Chr(13) & "You will be able to re-register this product on " & Format(CDate(dateLastUsed) + 31, "dddd dd mmmm yyyy", vbUseSystemDayOfWeek, vbUseSystem), vbInformation, "KTK License Registered Wizard"
            End If
        Else
            s = 1
            'has been registered
        End If
    Else
        'not registered
        s = 0
    End If
    
    If s = 1 Then
        MsgBox "We could not activate this product because it has already been activated!"
        picWait.Visible = False
        cmdNext.Visible = True: cmdCancel.Visible = True
        Exit Sub
    End If
    
    Inet1.OpenURL "members.lycos.co.uk/ktkpiracy/writeFile.php?productID=" & productid & "&licenseNo=" & txtSeed & "-" & txtVal1 & "-" & txtVal2 & "-" & txtVal3 & "&licenseHolder=" & txtName.Text & "&licenseHardwareKey=" & fingerPrint & "&daysEval=0&dateRegistered=" & Date & "&lastUsed=" & Date
    Me.picDone.Visible = True
    cmdCancel.Visible = True
    cmdCancel.Caption = "Close"
End Sub

Sub extractData(file)
    On Error Resume Next
    f = FreeFile
    Open file For Input As #f
    Input #f, header
    Input #f, licenseNo
    Input #f, licenseHolder
    Input #f, hardwareFingerPrint
    Input #f, evalDays
    Input #f, dateRegistered
    Input #f, dateLastUsed
    Input #f, footer
    Close #f
End Sub

Function fingerPrint()
    Dim TempStr As String
    Dim RegStr As String
    Dim I As Integer
    Dim SerialNumber As Long
    
    'Get The Computer Name in the registry
    'StartSysInfo
    SerialNumber = GetSerialNumber("C:\")
    SysInfoPath = Str(SerialNumber)
    
    'For encrypting purposes make the length
    'of it no more than 20 character
    If Len(SysInfoPath) > 20 Then
        SysInfoPath = Left$(SysInfoPath, 20)
    End If
    'invert the computer name
    InvertIt
    EncryptIt
    EncipherIt
    GetSubKey
    
    fingerPrint = SysInfoPath
    
End Function

'GetSerialNumber Procedure - Put this in the module or form where it is called.
Function GetSerialNumber(strDrive As String) As Long
    Dim SerialNum As Long
    Dim Res As Long
    Dim Temp1 As String
    Dim Temp2 As String
    'initialise the strings
    Temp1 = String$(255, Chr$(0))
    Temp2 = String$(255, Chr$(0))
    'call the API function
    Res = GetVolumeInformation(strDrive, Temp1, Len(Temp1), SerialNum, 0, 0, Temp2, Len(Temp2))
    GetSerialNumber = SerialNum
    
End Function

Sub EncipherIt()
    Dim Temp As Integer
    Dim Hold As String
    Dim I As Integer
    Dim J As Integer
    Dim TempStr As String
    Dim Temp1 As String
    
    TempStr = ""
    For I = 1 To Len(SysInfoPath)
        Temp = Asc(Mid$(SysInfoPath, I, 1))
        Temp1 = Hex(Temp)
        If Len(Temp1) = 1 Then
            Temp1 = "0" & Temp1
        End If
        For J = 1 To 2
            Hold = Mid$(Temp1, J, 1)
            Select Case Hold
                Case "0"
                    TempStr = TempStr + "7"
                Case "1"
                    TempStr = TempStr + "B"
                Case "2"
                    TempStr = TempStr + "F"
                Case "3"
                    TempStr = TempStr + "D"
                Case "4"
                    TempStr = TempStr + "1"
                Case "5"
                    TempStr = TempStr + "9"
                Case "6"
                    TempStr = TempStr + "3"
                Case "7"
                    TempStr = TempStr + "A"
                Case "8"
                    TempStr = TempStr + "6"
                Case "9"
                    TempStr = TempStr + "5"
                Case "A"
                    TempStr = TempStr + "E"
                Case "B"
                    TempStr = TempStr + "8"
                Case "C"
                    TempStr = TempStr + "0"
                Case "D"
                    TempStr = TempStr + "C"
                Case "E"
                    TempStr = TempStr + "2"
                Case "F"
                    TempStr = TempStr + "4"
            End Select
        Next J
    Next I
    SysInfoPath = TempStr
End Sub

Sub EncryptIt()
    Dim Temp As Integer
    Dim Temp1 As Integer
    Dim Hold As Integer
    Dim I As Integer
    Dim J As Integer
    Dim TempStr As String

    TempStr = ""
    For I = 1 To Len(EncryptName)
        Hold = 0
        Temp = Asc(Mid$(EncryptName, I, 1))
        For J = 1 To Len(SysInfoPath)
            Temp1 = Asc(Mid$(SysInfoPath, J, 1))
            Hold = Temp Xor Temp1
         Next J
        TempStr = TempStr + Chr(Hold)
    Next I
    
    SysInfoPath = TempStr
End Sub

Sub InvertIt()
    Dim Temp As Integer
    Dim Hold As Integer
    Dim I As Integer
    Dim TempStr As String
        
    TempStr = ""
    For I = 1 To Len(SysInfoPath)
        Temp = Asc(Mid$(SysInfoPath, I, 1))
        Hold = 0
Top:
    Select Case Temp
        Case Is > 127
            Hold = Hold + 1
            Temp = Temp - 128
            GoTo Top
        Case Is > 63
            Hold = Hold + 2
            Temp = Temp - 64
            GoTo Top
        Case Is > 31
            Hold = Hold + 4
            Temp = Temp - 32
            GoTo Top
        Case Is > 15
            Hold = Hold + 8
            Temp = Temp - 16
            GoTo Top
        Case Is > 7
            Hold = Hold + 16
            Temp = Temp - 8
            GoTo Top
        Case Is > 3
            Hold = Hold + 32
            Temp = Temp - 4
            GoTo Top
        Case Is > 1
            Hold = Hold + 64
            Temp = Temp - 2
            GoTo Top
        Case Is = 1
            Hold = Hold + 128
            
    End Select
        Temp = 255 Xor Hold
        TempStr = TempStr + Chr(Temp)
    Next I
    
    SysInfoPath = TempStr
End Sub

Public Sub GetSubKey()

    If Not GetKeyValue(HKEY_LOCAL_MACHINE, RegPath, RegKey, Register) Then
        'Rem Not in registry
        
    End If
    
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim I As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For I = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, I, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    End Select
    
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occured...
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function

Private Sub Form_Load()
    lblHardwareFingerPrint.Caption = "Computer Reference: " & fingerPrint
End Sub

Private Sub txtCode_Change()
    On Error Resume Next
    'we must split up the code
    
    'get the location of the '-'
    firstdash = InStr(1, txtCode, "-") - 1
    txtSeed = Left(txtCode, firstdash)
    
    'get the location of the 2nd '-'
    seconddash = InStr(firstdash + 2, txtCode, "-")
    txtVal1 = Mid(txtCode, firstdash + 2, seconddash - firstdash - 2)
    
    'get the location of the 3rd '-'
    thirddash = InStr(seconddash + 1, txtCode, "-")
    txtVal2 = Mid(txtCode, seconddash + 1, thirddash - seconddash - 1)
    
    'get the location of the 3rd '-'
    fourthdash = Len(txtSeed) + Len(txtVal1) + Len(txtVal2)
    txtVal3 = Mid(txtCode, fourthdash + 4)
    On Error GoTo 0
End Sub
