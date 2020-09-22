VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3480
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSeed 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Text            =   "Enter seed value"
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generate Numbers"
      Height          =   615
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "All generated keycodes are saved to c:\serial numbers.txt"
      Height          =   615
      Left            =   600
      TabIndex        =   7
      Top             =   960
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "Test Serial:3000-1-539E40-1"
      Height          =   495
      Left            =   1440
      TabIndex        =   5
      Top             =   2520
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    'this sub generate a load of codes and will
    'be able to see the code rejection rate
    
    Open "c:\Serial Numbers.txt" For Output As #1
    
    For x = 1 To 391
        createCode 1, x, 3000
    
        If validateCode(seed1, val11, val21, val31) = True Then
            success = success + 1
            Print #1, createCode(1, x, 3000)
        Else
            failure = failure + 1
        End If
    Next x
    
    For x = 1 To 188
        createCode 2, x, 3120
        
        If validateCode(seed1, val11, val21, val31) = True Then
            success = success + 1
            Print #1, createCode(2, x, 3120)
        Else
            failure = failure + 1
        End If
    Next x

    For x = 1 To 391
        createCode 1, x, 3004
    
        If validateCode(seed1, val11, val21, val31) = True Then
            success = success + 1
            Print #1, createCode(1, x, 3004)
        Else
            failure = failure + 1
        End If
    Next x
    
    For x = 1 To 188
        createCode 2, x, 3124
        
        If validateCode(seed1, val11, val21, val31) = True Then
            success = success + 1
            Print #1, createCode(2, x, 3124)
        Else
            failure = failure + 1
        End If
    Next x
    
    Close #1
    
    MsgBox "Total: " & success + failure & vbCrLf & "Success: " & success & vbCrLf & "Failure: " & failure
    
End Sub

Private Sub Text1_Change()
    If validateCode(Text1, Text2, Text3, Text4) = True Then
        Text1.BackColor = vbGreen
        Text2.BackColor = vbGreen
        Text3.BackColor = vbGreen
        Text4.BackColor = vbGreen
    Else
        Text1.BackColor = vbRed
        Text2.BackColor = vbRed
        Text3.BackColor = vbRed
        Text4.BackColor = vbRed
    End If
End Sub

Private Sub Text2_Change()
    Call Text1_Change
End Sub

Private Sub Text3_Change()
    Call Text1_Change
End Sub

Private Sub Text4_Change()
    Call Text1_Change
End Sub
