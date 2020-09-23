VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Main 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NIS Update"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Check 
      Caption         =   "Check"
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
   Begin MSComctlLib.ImageList Images 
      Left            =   5400
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":059A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0B34
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":10CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1668
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView trvUpdates 
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   3625
      _Version        =   393217
      Indentation     =   441
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "Images"
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3720
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "ucd"
      DialogTitle     =   "Open Update File..."
      FileName        =   "*.ini"
      Filter          =   "INI|*.ini"
   End
   Begin VB.Label Label2 
      Caption         =   $"Main.frx":1C02
      Height          =   855
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5295
   End
   Begin VB.Label Label1 
      Caption         =   "These are come from the INI file!"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   2775
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Function ReadINI(Section, KeyName, inifile As String) As String
    Dim sRet As String
    sRet = String(255, Chr(0))
    ReadINI = Left(sRet, GetPrivateProfileString(Section, ByVal KeyName, "", sRet, Len(sRet), inifile))
End Function
Private Sub Check_Click()
On Error GoTo diaexit
CommonDialog1.InitDir = App.Path
CommonDialog1.FileName = ""
CommonDialog1.ShowOpen
ReadUcd CommonDialog1.FileName, "001", App.Major, App.Minor, App.Revision
diaexit:
End Sub
Private Function ReadUcd(inUCD As String, inProductID As String, invermax As String, invermin As String, inverrev As String)
trvUpdates.Nodes.Clear
readprodid = ReadINI("Update Information", "Product", CommonDialog1.FileName)
If readprodid <> inProductID Then
MsgBox "Invaild file."
Exit Function
End If
updatemaj = ReadINI("Update Information", "UpdateVersionMajor", CommonDialog1.FileName)
updatemin = ReadINI("Update Information", "UpdateVersionMinor", CommonDialog1.FileName)
updaterev = ReadINI("Update Information", "UpdateVersionRevis", CommonDialog1.FileName)
updatever = updatemaj + "." + updatemin + "." + updaterev
curver = invermax & "." & invermin & "." & inverrev
If updatemaj <> invermax Then
MsgBox "Sorry, you can not upgrade from version " & curver & " to version " & updatever & "."
Exit Function
End If
If curver >= updatever Then
MsgBox "Sorry, you already have the latest version or a newer version."
Exit Function
End If
trvUpdates.Nodes.Add , , , "Visual CyBasic update to: " + updatever, 4
trvUpdates.Nodes.Add 1, tvwChild, , "GUI Updates", 3
trvUpdates.Nodes.Add 1, tvwChild, , "Editor Updates", 3
trvUpdates.Nodes.Add 1, tvwChild, , "Compiler Updates", 2
trvUpdates.Nodes.Add 1, tvwChild, , "Help Updates", 1
x = 1
Do Until countstring = "(None)"
countstring = ReadINI("GUI Updates", x, CommonDialog1.FileName)
If countstring = "(None)" And x <> 1 Then Exit Do
x = x + 1
trvUpdates.Nodes.Add 2, tvwChild, , countstring, 5
Loop
x = 1
Do Until countstring2 = "(None)"
countstring2 = ReadINI("Help Updates", x, CommonDialog1.FileName)
If countstring2 = "(None)" And x <> 1 Then Exit Do
x = x + 1
trvUpdates.Nodes.Add 5, tvwChild, , countstring2, 5
Loop
End Function

