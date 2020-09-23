VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTest 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Print Tree Control"
   ClientHeight    =   4185
   ClientLeft      =   1860
   ClientTop       =   2400
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   9270
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6660
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Frame fraOptions 
      Height          =   2295
      Left            =   6120
      TabIndex        =   3
      Top             =   960
      Width           =   3015
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1995
         Left            =   120
         ScaleHeight     =   1995
         ScaleWidth      =   2775
         TabIndex        =   4
         Top             =   180
         Width           =   2775
         Begin VB.OptionButton optChoice 
            Caption         =   "Print complete tree"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   10
            Top             =   420
            Value           =   -1  'True
            Width           =   1995
         End
         Begin VB.OptionButton optChoice 
            Caption         =   "Print complete tree && Write to file"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   9
            Top             =   660
            Width           =   2715
         End
         Begin VB.OptionButton optChoice 
            Caption         =   "Write complete tree to file"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   8
            Top             =   900
            Width           =   2715
         End
         Begin VB.OptionButton optChoice 
            Caption         =   "Printselected node && children"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   7
            Top             =   1140
            Width           =   2715
         End
         Begin VB.OptionButton optChoice 
            Caption         =   "Print selected node && Write to file"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   6
            Top             =   1380
            Width           =   2715
         End
         Begin VB.OptionButton optChoice 
            Caption         =   "Write selected node to file"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   5
            Top             =   1620
            Width           =   2715
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Select an option then press GO button"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   0
            TabIndex        =   11
            Top             =   120
            Width           =   2775
         End
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   1
      Left            =   8580
      TabIndex        =   2
      Top             =   3600
      Width           =   510
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Go"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   0
      Left            =   7920
      TabIndex        =   1
      Top             =   3600
      Width           =   570
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   7011
      _Version        =   393217
      Style           =   7
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdlPrinter 
      Left            =   6060
      Top             =   3540
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Print MS Treeview"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6420
      TabIndex        =   13
      Top             =   240
      Width           =   2595
   End
   Begin VB.Label lblAuthor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Kenneth Ives"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   7140
      TabIndex        =   12
      Top             =   660
      Width           =   975
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ***************************************************************
' Print Microsoft Treeview Demo
' by Kenneth Ives  kenaso@tx.rr.com
'
' use this form for testing only - builds a tree using Microsoft
' Treeview control
' ***************************************************************
Option Explicit
  
' ***************************************************************
' Constants
' ***************************************************************
  Private Const PGM_NAME     As String = "Print MS Treeview"
  Private Const AUTHOR_EMAIL As String = "kenaso@tx.rr.com"
  
' ***************************************************************
' API Declares
' ***************************************************************
  ' The ShellExecute function opens or prints a specified file.  The file
  ' can be an executable file or a document file.
  Private Declare Function ShellExecute Lib "shell32.dll" _
          Alias "ShellExecuteA" (ByVal hwnd As Long, _
          ByVal lpOperation As String, ByVal lpFile As String, _
          ByVal lpParameters As String, ByVal lpDirectory As String, _
          ByVal nShowCmd As Long) As Long

' ***************************************************************
' Module variables
' ***************************************************************
  Private mintNodeID     As Integer
  Private mintChoice     As Integer
  Private mstrReportPath As String
  
Private Sub Command1_Click(Index As Integer)
    
    Dim objTreePrint As clsTreePrint
    
    If TreeView1.Nodes.Count < 1 Then
        Exit Sub
    End If
                
    Select Case Index
           Case 0
                Set objTreePrint = New clsTreePrint   ' instantiate class
                
                Select Case mintChoice
                       Case 0     ' Print complete tree
                            objTreePrint.ReportTitle = "Print complete tree"
                            objTreePrint.PrintAllNodes = True  ' set property flag
                            mintNodeID = 1                     ' top level node
                            
                            If GetPrinterInfo Then
                                objTreePrint.PrintTreeview TreeView1, mintNodeID
                            End If
                            
                            InfoMsg "Finished printing complete tree"
                       
                       Case 1     ' print tree and save to file
                            objTreePrint.ReportTitle = "Print complete tree and save to file"
                            objTreePrint.ReportFile = mstrReportPath    ' pass the report destination
                            objTreePrint.PrintAllNodes = True     ' set property flag
                            mintNodeID = 1                        ' top level node
                            
                            If GetPrinterInfo Then
                                objTreePrint.PrintTreeview TreeView1, mintNodeID
                            End If
                            
                            objTreePrint.SaveToFile TreeView1, mintNodeID
                            InfoMsg "Finished creating " & mstrReportPath & " and printing."
                
                       Case 2    ' Save complete tree to file
                            objTreePrint.ReportTitle = "Save complete tree to file"
                            objTreePrint.ReportFile = mstrReportPath    ' pass the report destination
                            objTreePrint.PrintAllNodes = True     ' set property flag
                            mintNodeID = 1                        ' top level node
                        
                            objTreePrint.SaveToFile TreeView1, mintNodeID
                            InfoMsg "Finished creating " & mstrReportPath
                
                       Case 3    ' Print selected node
                            objTreePrint.ReportTitle = "Print selected tree node"
                            objTreePrint.PrintAllNodes = False  ' set property flag
                            
                            If GetPrinterInfo Then
                                objTreePrint.PrintTreeview TreeView1, mintNodeID
                            End If
                            
                            InfoMsg "Finished printing"
                
                       Case 4    ' Print selected node and save to file
                            objTreePrint.ReportTitle = "Print selected tree node and save to file"
                            objTreePrint.ReportFile = mstrReportPath     ' pass the report destination
                            objTreePrint.PrintAllNodes = False     ' set property flag
                            
                            If GetPrinterInfo Then
                                objTreePrint.PrintTreeview TreeView1, mintNodeID
                            End If
                
                            objTreePrint.SaveToFile TreeView1, mintNodeID
                            InfoMsg "Finished creating " & mstrReportPath & " and printing."
                
                       Case 5    ' Save selected node to file
                            objTreePrint.ReportTitle = "Save selected tree node to file"
                            objTreePrint.ReportFile = mstrReportPath     ' pass the report destination
                            objTreePrint.PrintAllNodes = False     ' set property flag
                            
                            objTreePrint.SaveToFile TreeView1, mintNodeID
                            InfoMsg "Finished creating " & mstrReportPath
                    
                End Select
                
           Case Else  ' Terminate program
                Unload Me
    End Select
    
    Set objTreePrint = Nothing

End Sub

Private Function GetPrinterInfo() As Boolean

    cdlPrinter.CancelError = True    ' Set Cancel to True
    
    On Error GoTo GetPrinterInfo_Error
    
    With cdlPrinter
        .Flags = cdlPDDisablePrintToFile Or cdlPDNoPageNums
        .PrinterDefault = True
        .ShowPrinter              ' Display the Print dialog box
    End With
    
    GetPrinterInfo = True
    Exit Function
    
GetPrinterInfo_Error:
    ' User pressed the Cancel button
    GetPrinterInfo = False
  
End Function

Private Sub Form_Load()
    
    Dim i        As Integer  ' Indexes (i thru o)
    Dim j        As Integer
    Dim k        As Integer
    Dim l        As Integer
    Dim m        As Integer
    Dim n        As Integer
    Dim o        As Integer
    Dim Key1     As String   ' Tree nodes
    Dim Key2     As String
    Dim Key3     As String
    Dim Key4     As String
    Dim Key5     As String
    Dim Key6     As String
    Dim Key7     As String
    Dim TreeText As String   ' Tree text
    Dim tvNode   As Node
    
    mintNodeID = 0
    mstrReportPath = App.Path & "\TreeRpt.txt"
    
    With TreeView1
        
        .Style = tvwTreelinesPlusMinusText   ' display text only
        .LineStyle = tvwRootLines            ' Use root level lines
        .Nodes.Clear                         ' Empty nodes on tree
    
        ' .Nodes.Add( , , Key1, TreeText)
        '            | |   |      |______ TreeText displayed on tree
        '            | |   |_____________ Used internally to track location
        '            | |_________________ Leave blank for root level else designate as child
        '            |___________________ Leave blank for root level else upper level key
        '
        '
        For i = 1 To 2
            Key1 = "Root " & CStr(i)
            TreeText = "Root " & CStr(i)
            
            Set tvNode = .Nodes.Add(, , Key1, TreeText)
            tvNode.Expanded = False
            
            For j = 1 To 2
                Key2 = Key1 & " Child " & CStr(j)
                TreeText = "Node level 2-" & CStr(j)
                Set tvNode = .Nodes.Add(Key1, tvwChild, Key2, TreeText)
                tvNode.Expanded = False
                
                For k = 1 To 2
                    Key3 = Key2 & "GrandChild " & CStr(k)
                    TreeText = "Node level 3-" & CStr(k)
                    Set tvNode = .Nodes.Add(Key2, tvwChild, Key3, TreeText)
                    tvNode.Expanded = False

                    For l = 1 To 2
                        Key4 = Key3 & "Sibling " & CStr(l)
                        TreeText = "Node level 4-" & CStr(l)
                        Set tvNode = .Nodes.Add(Key3, tvwChild, Key4, TreeText)
                        tvNode.Expanded = False
                        
                        ' uncomment these lines if you want to print lots of pages
                       For m = 1 To 2
                           Key5 = Key4 & "Sibling " & CStr(m)
                           TreeText = "Node level 5-" & CStr(m)
                           Set tvNode = .Nodes.Add(Key4, tvwChild, Key5, TreeText)
                           tvNode.Expanded = False

'                           For n = 1 To 2
'                               Key6 = Key5 & "Sibling " & CStr(n)
'                               TreeText = "Node level 6-" & CStr(n)
'                               Set tvNode = .Nodes.Add(Key5, tvwChild, Key6, TreeText)
'                               tvNode.Expanded = False
'
'                               For o = 1 To 2
'                                   Key7 = Key6 & "Sibling " & CStr(o)
'                                   TreeText = "Node level 7-" & CStr(o)
'                                   Set tvNode = .Nodes.Add(Key6, tvwChild, Key7, TreeText)
'                                   tvNode.Expanded = False
'                               Next o
'                           Next n
                       Next m
                    Next l
                Next k
            Next j
        Next i
    
    End With
    
    mintNodeID = 1
    Caption = "Print Microsoft Tree"
    Show
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Unload frmTest               ' deactivate form object
    Set frmTest = Nothing        ' free form object from memory
    
End Sub

Private Sub lblAuthor_Click()
    SendEmail
End Sub

Private Sub optChoice_Click(Index As Integer)
    mintChoice = Index          ' print or save option
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
    mintNodeID = Node.Index     ' Capture the node ID
End Sub

' ***************************************************************************
' Routine:       SendEmail
'
' Description:   When the email hyperlink is clicked, this routine will fire.
'                It will create a new email message with the author's name in
'                the "To:" box and the name and version of the application
'                on the "Subject:" line.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 23-FEB-2000  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Public Sub SendEmail()

    Dim strMail As String

    On Error GoTo SendEmail_Error

    ' Create email heading for user
    strMail = "mailto:" & AUTHOR_EMAIL & "?subject=" & PGM_NAME

    ' Call ShellExecute() API to create an email to the author
    ShellExecute 0&, vbNullString, strMail, _
                 vbNullString, vbNullString, vbNormalFocus

SendEmail_CleanUp:
    On Error GoTo 0
    Exit Sub

SendEmail_Error:
    InfoMsg Err.Number & vbCrLf & Err.Description, "eMail failure"
    Resume SendEmail_CleanUp

End Sub



