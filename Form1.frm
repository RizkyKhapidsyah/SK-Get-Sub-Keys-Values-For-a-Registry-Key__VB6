VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10755
   LinkTopic       =   "Form1"
   ScaleHeight     =   5850
   ScaleWidth      =   10755
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List2 
      Height          =   2790
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   10125
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   435
      Left            =   9330
      TabIndex        =   2
      Top             =   5280
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   10125
   End
   Begin VB.Label Label1 
      Caption         =   "Your installed applications"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1995
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'Label1 = GetSettingString(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", "")

Dim SubKeys As Variant
Dim KeyLoop As Integer
Dim sDispName As String
SubKeys = GetAllKeys(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")

If VarType(SubKeys) = vbArray + vbString Then
    For KeyLoop = 0 To UBound(SubKeys)
        sDispName = GetSettingString(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & SubKeys(KeyLoop), "DisplayName")
        If sDispName > "" And Left(SubKeys(KeyLoop), 1) <> "{" Then
            List1.AddItem sDispName & " / " & SubKeys(KeyLoop)
        End If
    Next
End If
End Sub


Sub CallCodeForGetAllValuesInAKey(ByVal sIn As String)
Dim Values As Variant
Dim KeyLoop As Integer
Dim RegPath As String
Dim HKCU As Long
Dim sLine As String
HKCU = HKEY_LOCAL_MACHINE 'to save typing
RegPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & sIn

Values = GetAllValues(HKCU, RegPath)

If VarType(Values) = vbArray + vbVariant Then

For KeyLoop = 0 To UBound(Values)
    sLine = Values(KeyLoop, 0) & " = "
    
    Select Case Values(KeyLoop, 1)
    Case REG_DWORD
        sLine = sLine & GetSettingLong(HKCU, RegPath, _
        CStr(Values(KeyLoop, 0)))
    Case REG_BINARY
        sLine = sLine & GetSettingByte(HKCU, RegPath, _
        Hex$(Values(KeyLoop, 0)))(0)
    Case REG_SZ
        sLine = sLine & GetSettingString(HKCU, RegPath, _
        CStr(Values(KeyLoop, 0)))
    End Select
    List2.AddItem sLine
Next KeyLoop

End If

End Sub

Private Sub List1_Click()
    If List1.ListIndex < 0 Then
        Exit Sub
    End If
    List2.Clear
    CallCodeForGetAllValuesInAKey Mid(List1.Text, InStr(List1.Text, " / ") + 3)
    
End Sub
