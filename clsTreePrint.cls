VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTreePrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ***************************************************************************
' Module:          clsTreePrint
'
' Description:     Print and save tree node data
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 29-APR-2004  Kenneth Ives  kenaso@tx.rr.com
'              Original
' ***************************************************************************
Option Explicit

' ********************************************************************
' Constants
' ********************************************************************
  Private Const MARGIN_LEFT    As Long = 720       ' approx 3/4 inch
  Private Const MARGIN_RIGHT   As Long = 720       ' approx 3/4 inch
  Private Const MARGIN_TOP     As Long = 1440      ' approx 1 inch
  Private Const MARGIN_BOTTOM  As Long = 1440      ' approx 1 inch
  Private Const INDENT_2       As Long = 2         ' tree indent level
  Private Const VERT_BAR_SOLID As String = "|"     ' Vertical bar (Solid)
  Private Const INDENT_LINE    As String = "__"    ' Horizontal line

' ***************************************************************************
' Variables
'
' Variable name:     mstrOutput
' Naming standard:   m str Output
'                    - --- ---------
'                    |  |    |______ Variable subname
'                    |  |___________ Data type (String)
'                    |______________ Global level designator
'
' ***************************************************************************
  Private mblnPrintAllNodes As Boolean   ' process all tree nodes
  Private mintIndentLevel   As Integer   ' Tree indent level
  Private mstrReportTitle   As String    ' report title
  Private mstrReportFile    As String    ' report path and file name
  Private mstrOutput        As String    ' output print string



' ***************************************************************************
' ****                      Properties                                   ****
' ***************************************************************************

Public Property Let PrintAllNodes(ByVal blnData As Boolean)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.PrintAllNodes = True
    mblnPrintAllNodes = blnData
End Property

Public Property Let ReportTitle(ByVal strData As String)
    ' Used when assigning a value to the property, on the left side
    ' of an assignment.  Syntax:  X.ReportTitle = 5
    mstrReportTitle = strData
End Property

Public Property Let ReportFile(ByVal strData As String)
    ' Used when assigning a value to the property, on the left side
    ' of an assignment.  Syntax:  X.ReportFile = 5
    mstrReportFile = strData
End Property



' ***************************************************************************
' ***                           Methods                                   ***
' ***************************************************************************

' ***************************************************************************
' Routine:         PrintTreeview
'
' Description:     Open the printer dialog box and reset the default printer.
'
' Parameters:      objTree - Picture of resource tree
'                  intNodeID - Position on the tree to start at
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 29-APR-2004  Kenneth Ives  kenaso@tx.rr.com
'              Original
' ***************************************************************************
Public Sub PrintTreeview(ByRef objTree As TreeView, _
                Optional ByVal intNodeID As Integer = 1)

    Dim tvNode As Node  ' tree node
    
    ' See if there are any nodes on the tree.
    ' If not, then leave.
    If objTree.Nodes.Count < 1 Then
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    mintIndentLevel = 0
    mstrOutput = ""
    Set tvNode = objTree.Nodes(intNodeID) ' start at the selected tree level
    
    Do
        FormatPrintOutput tvNode
    Loop Until tvNode Is Nothing
        
    ' print the report
    PrintHeaderFooter False   ' print page header/footer but do not eject a page
    PrintTreeData             ' send data to printer
    Screen.MousePointer = vbDefault
        
End Sub

' ***************************************************************************
' Routine:         SaveToFile
'
' Description:     Save the tree data to a text file.
'
' Parameters:      objTree - Picture of resource tree
'                  intNodeID - Position on the tree to start at
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 29-APR-2004  Kenneth Ives  kenaso@tx.rr.com
'              Original
' ***************************************************************************
Public Sub SaveToFile(ByRef objTree As TreeView, _
             Optional ByVal intNodeID As Integer = 1)

    Dim tvNode As Node      ' tree node
    Dim hFile  As Integer   ' File handle
    
    ' See if there are any nodes on the tree.
    ' If not, then leave.
    If objTree.Nodes.Count < 1 Then
        Exit Sub
    End If
    
    ' If no report title was passed the use the default.
    If Len(Trim$(mstrReportTitle)) = 0 Then
        mstrReportTitle = "<Unknown>"
    End If
        
    Screen.MousePointer = vbHourglass
    mintIndentLevel = 0
    
    ' Open up a text file to export the indented TreeView nodes into.
    hFile = FreeFile
    Open mstrReportFile For Output As #hFile   ' create the file
    Print #hFile, "Rpt Name:    " & mstrReportTitle
    Print #hFile, "Created on:  " & Format$(Now(), "dd mmm yyyy   h:mm ampm")
    Print #hFile, String$(60, "-")
    Print #hFile, " "
    
    Set tvNode = objTree.Nodes(intNodeID)   ' start at the root level of the tree
    CreateFile tvNode, hFile
    Close #hFile
    
    Screen.MousePointer = vbDefault
    
End Sub



' ***************************************************************************
' ****              Internal Functions and Procedures                    ****
' ***************************************************************************

' ***************************************************************************
' Routine:         CreateFile
'
' Description:     Create the file copy of the tree data
'
' Parameters:      tvNode - Node to begin processing
'                  hFile - File handle
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 29-APR-2004  Kenneth Ives  kenaso@tx.rr.com
'              Original
' ***************************************************************************
Private Sub CreateFile(ByRef tvNode As Node, _
                       ByVal hFile As Integer)

    Dim strPrintLine As String    ' output print string
    
    ' * Create tree lines in the file
    strPrintLine = FormatTreeLines(tvNode)
    Print #hFile, strPrintLine
    
    ' Check to see if the current node has children
    If tvNode.Children > 0 Then
        mintIndentLevel = mintIndentLevel + 1   ' Increase the indent if children exist
        CreateFile tvNode.Child, hFile
    End If
        
    ' only test if processing the complete tree
    If Not mblnPrintAllNodes Then
        If mintIndentLevel = 0 Then
            Exit Sub
        End If
    End If
    
    Set tvNode = tvNode.Next   ' Set the next node to capture
    
    ' As long as we have not reached the last node in
    ' a branch, continue to call the print routine
    If TypeName(tvNode) <> "Nothing" Then
        CreateFile tvNode, hFile
    Else
        ' If the last node of a branch was reached,
        ' decrease the indentation counter
        mintIndentLevel = mintIndentLevel - 1
    End If

End Sub

' ***************************************************************************
' Routine:         FormatPrintOutput
'
' Description:     Print the tree data without lines
'
' Parameters:      tvNode - Node to begin processing
'
' Returns:         None
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 29-APR-2004  Kenneth Ives  kenaso@tx.rr.com
'              Original
' ***************************************************************************
Private Sub FormatPrintOutput(ByRef tvNode As Node)

    ' Create tree lines
    mstrOutput = mstrOutput & FormatTreeLines(tvNode) & vbCrLf
    
    ' Check to see if the current node has children
    If tvNode.Children > 0 Then
        mintIndentLevel = mintIndentLevel + 1   ' Increase the indent if children exist
        FormatPrintOutput tvNode.Child
    End If
    
    Set tvNode = tvNode.Next                    ' Set the next node to capture
    
    ' As long as we have not reached the last node in
    ' a branch, continue to call the print routine
    If TypeName(tvNode) <> "Nothing" Then
        FormatPrintOutput tvNode
    Else
        ' If the last node of a branch was reached,
        ' decrease the indentation counter
        mintIndentLevel = mintIndentLevel - 1
    End If

End Sub

' ***************************************************************************
' Routine:         PrintTreeData
'
' Description:     Print the tree data with full lines (Version 2)
'
' Parameters:      tvNode - Node to begin processing
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 29-APR-2004  Kenneth Ives  kenaso@tx.rr.com
'              Original
' ***************************************************************************
Private Sub PrintTreeData()

    Dim lngPointer     As Long
    Dim lngIndent      As Long
    Dim strPrintOutput As String
    
    lngPointer = 1
    
    ' Print text, word-wrapping as we go
    Do Until lngPointer > Len(mstrOutput)

        ' Get next word
        strPrintOutput = ""
        Do Until lngPointer > Len(mstrOutput) Or Mid$(mstrOutput, lngPointer, 2) = vbCrLf
            strPrintOutput = strPrintOutput & Mid$(mstrOutput, lngPointer, 1)
            lngPointer = lngPointer + 1
        Loop
        '
        ' Check if word will fit on this line
        If (Printer.CurrentX + Printer.TextWidth(strPrintOutput)) > (Printer.ScaleWidth - MARGIN_RIGHT) Then

            ' Send carriage-return line-feed to printer
            Printer.Print

            ' Check if we need to start a new page
            If Printer.CurrentY > (Printer.ScaleHeight - MARGIN_BOTTOM) Then
                PrintHeaderFooter True
            Else
                Printer.CurrentX = MARGIN_LEFT
            End If
        End If

        ' Print this word
        Printer.Print strPrintOutput;

        ' Process whitespace and any control characters
        Do Until lngPointer > Len(mstrOutput) Or _
                 Mid$(mstrOutput, lngPointer, 1) > " "

            Select Case Asc(Mid$(mstrOutput, lngPointer, 1))
            
                   Case 0, 32  ' null value or blank space
                        Printer.Print " ";
                        
                   Case 9      ' Tab
                        lngIndent = (Printer.CurrentX - MARGIN_LEFT) / Printer.TextWidth("0")
                        lngIndent = lngIndent + (10 - (lngIndent Mod 10))
                        Printer.CurrentX = MARGIN_LEFT + (lngIndent * Printer.TextWidth("0"))
                  
                   Case 10     ' Line-feed
                        Printer.Print   ' Send carriage-return line-feed to printer
  
                        ' Check if we need to start a new page
                        If Printer.CurrentY > (Printer.ScaleHeight - MARGIN_BOTTOM) Then
                            PrintHeaderFooter True
                        Else
                            Printer.CurrentX = MARGIN_LEFT
                        End If
    
                  Case Else   ' Ignore other characters
            End Select

            ' Increment the counter
            lngPointer = lngPointer + 1
        Loop
    Loop
    
    ' Notify printer that we are finished with this document
    Printer.EndDoc

End Sub

' ***************************************************************************
' Routine:         PrintHeaderFooter
'
' Description:     Print the report header and footer
'
' Parameters:      blnEjectPage - True=Eject a page from the printer
'                                 False=Do not eject a page
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 29-APR-2004  Kenneth Ives  kenaso@tx.rr.com
'              Original
' ***************************************************************************
Private Sub PrintHeaderFooter(ByVal blnEjectPage As Boolean)
    
    Dim strTemp As String
    
    ' Start new page if requested
    If blnEjectPage Then
        Printer.NewPage
    End If
    
    ' Define printer font for header/footer
    Printer.Print
    Printer.Font = "Arial"
    Printer.FontSize = 12
    Printer.FontBold = True

    ' Print page header
    Printer.CurrentY = (MARGIN_TOP - Printer.TextHeight(mstrReportTitle)) / 2
    Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(mstrReportTitle)) / 2
    Printer.Print mstrReportTitle;
    
    ' Print page footer
    strTemp = StrConv(Format(Now, "d mmm yyyy  hh:mm:ss"), vbProperCase)
    Printer.CurrentX = MARGIN_LEFT
    Printer.CurrentY = Printer.ScaleHeight - (MARGIN_BOTTOM / 2)
    Printer.Print strTemp;
    
    strTemp = "Page  " & Printer.Page
    Printer.CurrentX = Printer.ScaleWidth - (Printer.TextWidth(strTemp) + MARGIN_RIGHT)
    Printer.CurrentY = Printer.ScaleHeight - (MARGIN_BOTTOM / 2)
    Printer.Print strTemp;
    Printer.Print
    
    ' Reset position to top of page and redefine font for printing the body.
    Printer.Font = "Arial"
    Printer.FontSize = 10
    Printer.FontBold = False
    Printer.CurrentX = MARGIN_LEFT
    Printer.CurrentY = MARGIN_TOP
    Printer.Print

End Sub

' ***************************************************************************
' Routine:         FormatTreeLines
'
' Description:     Format the line for the output data
'
' Parameters:      tvNode - Current node being evaluated
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 29-APR-2004  Kenneth Ives  kenaso@tx.rr.com
'              Original
' ***************************************************************************
Private Function FormatTreeLines(ByRef tvNode As Node) As String

    Dim strVertBars  As String    ' prefixing vertical bars
    Dim strPrintLine As String    ' return data string
    Dim intLevel     As Integer   ' node level
    Dim intIndex     As Integer   ' loop counter
    
    strVertBars = ""
    strPrintLine = ""
    
    ' Format the node that was passed in and account for the node's level
    Select Case mintIndentLevel
            
           Case 0     ' This is the root level
                strPrintLine = tvNode.Text
                
           Case 1     ' This is the second level (one off of root)
                strPrintLine = Space$(INDENT_2) & VERT_BAR_SOLID & INDENT_LINE & tvNode.Text
    
           Case 2     ' This is the Third level (two off of root)
                intLevel = mintIndentLevel - 2
                strVertBars = Space$(INDENT_2) & VERT_BAR_SOLID & INDENT_LINE  ' initialize first indent
                ' Append horizontal line and the node text
                strPrintLine = strVertBars & INDENT_LINE & tvNode.Text
    
           Case Else  ' determine the node level and assign a multipler value
                intLevel = mintIndentLevel - 2
                strVertBars = Space$(INDENT_2) & VERT_BAR_SOLID & INDENT_LINE & INDENT_LINE
        
                ' format the number of vertical bars to prefix a line
                For intIndex = 1 To intLevel
                    strVertBars = strVertBars & VERT_BAR_SOLID & INDENT_LINE & INDENT_LINE
                Next intIndex
        
                ' Append horizontal line and the node text
                strPrintLine = strVertBars & tvNode.Text
    End Select

    FormatTreeLines = strPrintLine    ' return formatted data
    
End Function

Private Sub Class_Initialize()
    PrintAllNodes = True
End Sub
