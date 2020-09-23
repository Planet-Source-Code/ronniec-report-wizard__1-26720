VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmSelect 
   Caption         =   "Report Wizard"
   ClientHeight    =   9345
   ClientLeft      =   2070
   ClientTop       =   795
   ClientWidth     =   10275
   LinkTopic       =   "Form1"
   ScaleHeight     =   9345
   ScaleWidth      =   10275
   Begin VB.Frame progFrame 
      Caption         =   "Progress"
      Height          =   735
      Left            =   150
      TabIndex        =   35
      Top             =   8370
      Visible         =   0   'False
      Width           =   8235
      Begin MSComctlLib.ProgressBar OutPg 
         Height          =   270
         Left            =   30
         TabIndex        =   36
         Top             =   390
         Width           =   8115
         _ExtentX        =   14314
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblpg2 
         Caption         =   "100 %"
         Height          =   225
         Left            =   7635
         TabIndex        =   38
         Top             =   150
         Width           =   510
      End
      Begin VB.Label lblpg1 
         Caption         =   "0 %"
         Height          =   225
         Left            =   60
         TabIndex        =   37
         Top             =   165
         Width           =   510
      End
   End
   Begin VB.TextBox txtNORecs 
      Height          =   345
      Left            =   9360
      TabIndex        =   33
      Top             =   4620
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.CommandButton cmdCriteria 
      Caption         =   "Criteria"
      Enabled         =   0   'False
      Height          =   465
      Left            =   7005
      TabIndex        =   32
      Top             =   3405
      Width           =   1155
   End
   Begin VB.CommandButton cmdOutputRep 
      Caption         =   "Output To Word"
      Enabled         =   0   'False
      Height          =   390
      Left            =   4230
      TabIndex        =   31
      Top             =   4470
      Width           =   2580
   End
   Begin VB.CommandButton CmdViewRes 
      Caption         =   "View Results"
      Height          =   390
      Left            =   1440
      TabIndex        =   30
      Top             =   4470
      Width           =   2580
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset Lists"
      Height          =   645
      Left            =   3165
      TabIndex        =   28
      Top             =   2010
      Width           =   1185
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   645
      Left            =   8580
      TabIndex        =   27
      Top             =   8415
      Width           =   1665
   End
   Begin VB.CommandButton cmdOrdSort 
      Caption         =   "Sort Order"
      Enabled         =   0   'False
      Height          =   465
      Left            =   7005
      TabIndex        =   26
      Top             =   2525
      Width           =   1155
   End
   Begin VB.Frame frameCriteria 
      Caption         =   "Criteria"
      Height          =   2715
      Left            =   4035
      TabIndex        =   12
      Top             =   5055
      Visible         =   0   'False
      Width           =   6165
      Begin VB.TextBox txtCrit 
         Height          =   315
         Index           =   3
         Left            =   3720
         TabIndex        =   25
         Top             =   1470
         Visible         =   0   'False
         Width           =   1965
      End
      Begin VB.TextBox txtCrit 
         Height          =   315
         Index           =   2
         Left            =   3720
         TabIndex        =   24
         Top             =   1090
         Visible         =   0   'False
         Width           =   1965
      End
      Begin VB.TextBox txtCrit 
         Height          =   315
         Index           =   1
         Left            =   3720
         TabIndex        =   23
         Top             =   695
         Visible         =   0   'False
         Width           =   1965
      End
      Begin VB.TextBox txtCrit 
         Height          =   315
         Index           =   0
         Left            =   3720
         TabIndex        =   22
         Top             =   300
         Width           =   1965
      End
      Begin VB.ComboBox lstOper 
         Height          =   315
         Index           =   3
         Left            =   2670
         TabIndex        =   21
         Top             =   1470
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.ComboBox lstOper 
         Height          =   315
         Index           =   2
         Left            =   2670
         TabIndex        =   20
         Top             =   1090
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.ComboBox lstOper 
         Height          =   315
         Index           =   1
         Left            =   2670
         TabIndex        =   19
         Top             =   695
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.ComboBox lstOper 
         Height          =   315
         Index           =   0
         Left            =   2670
         TabIndex        =   18
         Top             =   300
         Width           =   900
      End
      Begin VB.ComboBox lstCrit 
         Height          =   315
         Index           =   3
         Left            =   150
         TabIndex        =   17
         Top             =   1470
         Visible         =   0   'False
         Width           =   2340
      End
      Begin VB.ComboBox lstCrit 
         Height          =   315
         Index           =   2
         Left            =   150
         TabIndex        =   16
         Top             =   1090
         Visible         =   0   'False
         Width           =   2340
      End
      Begin VB.ComboBox lstCrit 
         Height          =   315
         Index           =   1
         Left            =   150
         TabIndex        =   15
         Top             =   695
         Visible         =   0   'False
         Width           =   2340
      End
      Begin VB.ComboBox lstCrit 
         Height          =   315
         Index           =   0
         Left            =   150
         TabIndex        =   14
         Top             =   300
         Width           =   2340
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Reset"
         Height          =   390
         Left            =   105
         TabIndex        =   13
         Top             =   2085
         Width           =   870
      End
   End
   Begin VB.ListBox lstAvail 
      Height          =   3765
      Left            =   645
      TabIndex        =   8
      Top             =   405
      Width           =   2310
   End
   Begin VB.Frame frameOrd 
      Caption         =   "Order By"
      Height          =   2715
      Left            =   285
      TabIndex        =   5
      Top             =   5055
      Visible         =   0   'False
      Width           =   2670
      Begin VB.ComboBox lstOrd 
         Height          =   315
         Index           =   3
         Left            =   135
         TabIndex        =   11
         Top             =   1530
         Visible         =   0   'False
         Width           =   2310
      End
      Begin VB.ComboBox lstOrd 
         Height          =   315
         Index           =   2
         Left            =   135
         TabIndex        =   10
         Top             =   1130
         Visible         =   0   'False
         Width           =   2310
      End
      Begin VB.ComboBox lstOrd 
         Height          =   315
         Index           =   1
         Left            =   135
         TabIndex        =   9
         Top             =   730
         Visible         =   0   'False
         Width           =   2310
      End
      Begin VB.CommandButton cmdOrdReset 
         Caption         =   "Reset"
         Height          =   390
         Left            =   105
         TabIndex        =   7
         Top             =   2085
         Width           =   870
      End
      Begin VB.ComboBox lstOrd 
         Height          =   315
         Index           =   0
         Left            =   135
         TabIndex        =   6
         Top             =   330
         Width           =   2310
      End
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "Move Down"
      Enabled         =   0   'False
      Height          =   465
      Left            =   7005
      TabIndex        =   4
      Top             =   1645
      Width           =   1155
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "Move Up"
      Enabled         =   0   'False
      Height          =   465
      Left            =   7005
      TabIndex        =   3
      Top             =   765
      Width           =   1155
   End
   Begin VB.CommandButton cmdRem 
      Caption         =   "Remove"
      Height          =   495
      Left            =   3135
      TabIndex        =   2
      Top             =   3690
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   495
      Left            =   3135
      TabIndex        =   1
      Top             =   450
      Width           =   1215
   End
   Begin VB.ListBox lstSelected 
      Height          =   3765
      Left            =   4575
      TabIndex        =   0
      Top             =   405
      Width           =   2205
   End
   Begin MSDataGridLib.DataGrid wizGrid 
      Height          =   3240
      Left            =   255
      TabIndex        =   29
      Top             =   5055
      Visible         =   0   'False
      Width           =   9990
      _ExtentX        =   17621
      _ExtentY        =   5715
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   6153
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   6153
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label lblnorecs 
      Caption         =   "Records Found"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7695
      TabIndex        =   34
      Top             =   4680
      Visible         =   0   'False
      Width           =   1590
   End
End
Attribute VB_Name = "frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Dim rs_fields As ADODB.Recordset
Dim strcon As String
Dim RS_REP As ADODB.Recordset
Dim rsord, rscrit As Recordset
Dim objWrd As Word.Application
Dim mboolWordRunning As Boolean

Private Sub cmdAdd_Click()
    AddItem
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdCriteria_Click()
If wizGrid.Visible = True Then
    wizGrid.Visible = False
End If
    lstCritRS
    frameCriteria.Visible = True
    CritLists (0)
End Sub

Private Sub cmdDown_Click()
    Dim strIndex As String   '-- hold the selected index data temporarily for move
    
    Dim i As Integer   '-- holds the index of the item to be moved
    
    i = lstSelected.ListIndex
    
    If i > -1 Then
         
         strIndex = lstSelected.List(i)
        
        '-- Add the item selected to one position above the current position
        lstSelected.AddItem strIndex, (i + 2)
        
        '-- remove it from the current position. Note the current position has changed because the add has moved everything down by 1
        lstSelected.RemoveItem (i)
        
        '-- Reselect the item that was moved.
        lstSelected.Selected(i + 1) = True
    
    End If
End Sub


Private Sub cmdOrdSort_Click()
If wizGrid.Visible = True Then
    wizGrid.Visible = False
End If
    LstOrdRS
    frameOrd.Visible = True
    OrderLists (0)
End Sub

Private Sub cmdRem_Click()
    RemItem
End Sub

Private Sub cmdReset_Click()
lstSelected.Clear
lstAvail.Clear
FillLists
End Sub

Private Sub cmdUp_Click()

    Dim strIndex As String   '-- hold the selected index data temporarily for move
    
    Dim i As Integer   '-- holds the index of the item to be moved
    
    i = lstSelected.ListIndex
    
    If i > -1 Then
         
         strIndex = lstSelected.List(i)
        
        '-- Add the item selected to one position above the current position
        lstSelected.AddItem strIndex, (i - 1)
        
        '-- remove it from the current position. Note the current position has changed because the add has moved everything down by 1
        lstSelected.RemoveItem (i + 1)
        
        '-- Reselect the item that was moved.
        lstSelected.Selected(i - 1) = True
    
    End If
    
End Sub

Private Sub CmdViewRes_Click()
On Error GoTo prev_err
Set RS_REP = New ADODB.Recordset
'Create the report recordset
With RS_REP
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Source = StrSql 'function buildes the string according to the select fields
    .ActiveConnection = con
    .Open
End With

'fills the grid with results
With wizGrid
    Set .DataSource = RS_REP
    .Refresh
    .Visible = True
    .Enabled = True
End With
frameOrd.Visible = False
frameCriteria.Visible = False

cmdOutputRep.Enabled = True
txtNORecs.Visible = True
lblnorecs.Visible = True

txtNORecs = RS_REP.RecordCount
Exit Sub

prev_err:
    MsgBox "Cannot create results set." & _
    "Please review criteria. Contact Administrator" & _
    "if problem Persists.", vbOKOnly, "Rept Wiz"
    
End Sub

Private Sub CmdOutputRep_Click()
'this section outputs the report results set to word
On Error GoTo out_err
Dim X As Integer
Dim newline As String
progFrame.Visible = True
OutPg.Value = 0

Dim xaxis, yaxis As Integer
xaxis = RS_REP.RecordCount
yaxis = RS_REP.Fields.Count

Screen.MousePointer = vbHourglass
'Startup Word if not started, or switch to existing one
Call modGetWord
objWrd.Application.ScreenUpdating = False
objWrd.Visible = False
objWrd.Application.WindowState = wdWindowStateMinimize
objWrd.Documents.Add
newline = Chr$(13) & Chr$(10)

objWrd.Selection.InsertAfter Format(Now(), "dd/mm/yyyy") 'TxtTitle.Text  'set title of report
objWrd.Selection.ParagraphFormat.Alignment = wdAlignParagraphRight 'align title in centre of page
objWrd.Selection.Font.Bold = True 'bold the text
objWrd.Selection.Font.Italic = True
objWrd.Selection.Font.Name = "Arial" 'set font to arial
objWrd.Selection.Font.Size = 12
objWrd.Selection.InsertAfter newline
objWrd.Selection.MoveDown wdLine, 1
objWrd.Selection.InsertAfter InputBox("Please Enter a Name for Your Report", "MMC ReptWiz")   'TxtTitle.Text  'set title of report
objWrd.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter 'align title in centre of page
objWrd.Selection.Font.Bold = True 'bold the text
objWrd.Selection.Font.Underline = wdUnderlineSingle 'underscore
objWrd.Selection.Font.Name = "Arial" 'set font to arial
objWrd.Selection.Font.Size = 18 'set size of font to 18

objWrd.Selection.InsertAfter newline & newline 'insert 2 new lines
objWrd.Selection.MoveDown wdLine, 2  'move down the two lines as word extends when insertion
objWrd.Selection.InsertAfter newline & newline 'insert 2 new lines
objWrd.Selection.MoveDown wdLine, 2 'move down the two lines

objWrd.ActiveDocument.PageSetup.Orientation = wdOrientLandscape 'set the page orientation to landscape
objWrd.ActiveWindow.View.Type = wdPageView 'change view of report to page view
'format table to number of columns in grid
objWrd.ActiveDocument.Tables.Add Range:=Selection.Range, NumRows:=3, NumColumns:=yaxis   'wizGrid.Columns.Count
objWrd.Selection.SelectRow 'select the entire top row
objWrd.Selection.Font.Underline = wdUnderlineNone 'no underscore
objWrd.Selection.Font.Size = 12
    With objWrd.Selection.Cells
        With .Shading
            .Texture = wdTexture25Percent
            .ForegroundPatternColorIndex = wdAuto
            .BackgroundPatternColorIndex = wdWhite
        End With
    End With
objWrd.Selection.Rows.HeadingFormat = wdToggle 'set first row as header so that they show on subsequent pages of report
objWrd.Selection.MoveLeft

'loop round the grid to get the column headers to make the report headers
For i = 0 To wizGrid.Columns.Count - 1
    lstSelected.ListIndex = i
    objWrd.Selection.InsertAfter lstSelected  'wizGrid.Columns(i).Caption
    objWrd.Selection.MoveRight wdCell, 1
Next

objWrd.Selection.SelectRow
objWrd.Selection.Font.Underline = wdUnderlineNone
objWrd.Selection.Font.Bold = False
objWrd.Selection.Font.Italic = False
objWrd.Selection.MoveLeft
RS_REP.MoveFirst
RS_REP.MoveFirst 'move to the first record int he ado control

'loop round the resultset of the contol and insert the values into the table

 '
objWrd.Selection.SelectRow
        objWrd.Selection.Font.Bold = False
        objWrd.Selection.Font.Underline = wdUnderlineNone
        objWrd.Selection.Font.Name = "Arial" 'set font to arial
        objWrd.Selection.Font.Size = 11
        
OutPg.Max = RS_REP.RecordCount

Do While Not RS_REP.EOF
        For X = 0 To RS_REP.Fields.Count - 1
        objWrd.Selection.InsertAfter "" & RS_REP.Fields(X).Value
        objWrd.Selection.MoveRight wdCell, 1 'move to the next cell in the table
        Next
    
    objWrd.Selection.InsertRows 1
    objWrd.Selection.MoveLeft Unit:=wdCharacter, Count:=1

RS_REP.MoveNext 'next record int he ado control
OutPg.Value = OutPg.Value + 1
Loop

'delete any empty rows which may be in the table
Do While objWrd.Selection.Information(wdWithInTable)
     objWrd.Selection.SelectRow
     objWrd.Selection.Rows.Delete
Loop

objWrd.Application.Activate 'activate word
objWrd.ActiveWindow.ActivePane.View.Zoom.Percentage = 75 'zoom the report to 100%
objWrd.Application.ScreenUpdating = True
objWrd.Application.WindowState = wdWindowStateMaximize
objWrd.Visible = True
Set objWrd = Nothing 'clear object variable
Screen.MousePointer = vbDefault

Exit Sub
out_err:
    MsgBox "Cannot create Document. Contact Administrator" & _
    "if problem Persists.", vbOKOnly, "Rept Wiz"
    Set objWrd = Nothing
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Load()
    FillLists
End Sub

Sub AddItem()
'adds item to selected list
lstSelected.AddItem (lstAvail)
lstAvail.RemoveItem (lstAvail.ListIndex)
BttnEnable

End Sub

Sub RemItem()
'removes item from selected list
lstAvail.AddItem (lstSelected)
lstSelected.RemoveItem (lstSelected.ListIndex)
lstAvail.Refresh
BttnEnable
End Sub


Private Sub lstAvail_DblClick()
    AddItem
End Sub

Private Sub lstCrit_Click(Index As Integer)
'this allows you to select the criteria for teh recordset
'once a field is selected it removes from the list
rscrit.MoveFirst
Do
    If rscrit!CritField Like lstCrit(Index).Text Then
        rscrit.Delete adAffectCurrent
    End If
    rscrit.MoveNext
Loop Until rscrit.EOF

'allows for four criteria.  checks to see if control exists before showing
If Exists("lstCrit", Index + 1) Then
    lstCrit(Index + 1).Visible = True
    lstOper(Index + 1).Visible = True
    txtCrit(Index + 1).Visible = True
    CritLists (Index + 1)
End If

End Sub

Private Sub lstOrd_Click(Index As Integer)
'section to order list  i.e. Order By
rsord.MoveFirst
Do
    If rsord!OrdField Like lstOrd(Index).Text Then
        rsord.Delete adAffectCurrent
    End If
    rsord.MoveNext
Loop Until rsord.EOF

If Exists("lstOrd", Index + 1) Then
    lstOrd(Index + 1).Visible = True
    OrderLists (Index + 1)
End If
End Sub

Private Sub lstSelected_DblClick()
    RemItem
End Sub

Sub FillLists()
'opens recordset and fills list with column names from criteria
Set con = New ADODB.Connection
Set rs_fields = New ADODB.Recordset
'the connection string
strcon = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Program Files\Microsoft Visual Studio\VB98\NWIND.MDB;"

With con
    .CursorLocation = adUseClient
    .ConnectionString = strcon
    .Open
End With

'this recordset looks up the column names in the query and adds them to the list
With rs_fields
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    'the attribute for a col name in Msysqueries is '6'. The object Id you will have to find in the
    'msysobjects table.  it is the object id for the query you have created
    .Source = "SELECT EXPRESSION FROM MSYSQUERIES WHERE ATTRIBUTE = 6 AND OBJECTID = -2147483566"
    .ActiveConnection = con
    .Open
End With
'adds fields to the list
With rs_fields
    .MoveFirst
    Do
    lstAvail.AddItem (!EXPRESSION)
    .MoveNext
    Loop Until .EOF
.Close
End With

End Sub

Sub BttnEnable()
If lstSelected.ListCount > 1 Then
    cmdUp.Enabled = True
    cmdDown.Enabled = True
    cmdOrdSort.Enabled = True
    cmdCriteria.Enabled = True
Else
    cmdUp.Enabled = False
    cmdDown.Enabled = False
    cmdOrdSort.Enabled = False
    cmdCriteria.Enabled = False
End If
End Sub

Function StrSql() As String

'this function builds the query string according to what is selected and criteria etc.
Dim StrSel, StrOrd, StrWhere As String

Dim n As Integer
    
    For n = 0 To lstSelected.ListCount - 1
        lstSelected.ListIndex = n
        StrSel = StrSel & lstSelected & ","
    Next n
StrSel = "SELECT " & Left(StrSel, (Len(StrSel) - 1)) & " FROM NEWQRY"

    For n = 0 To lstOrd.Count - 1
        If Not lstOrd(n).Text = "" Then
            StrOrd = StrOrd & lstOrd(n).Text & ","
        End If
    Next n
If StrOrd <> "" Then
    StrOrd = " Order By " & Left(StrOrd, (Len(StrOrd) - 1))
End If

For n = 0 To lstCrit.Count - 1
        If Not lstCrit(n).Text = "" Then
            StrWhere = StrWhere & lstCrit(n).Text & " " & lstOper(n).Text & " " & txtCrit(n).Text & " And "
        End If
    Next n
If StrWhere <> "" Then
    StrWhere = " Where " & Left(StrWhere, (Len(StrWhere) - 4))
End If
If StrOrd <> "" And StrWhere <> "" Then
    StrSql = StrSel & StrWhere & StrOrd
ElseIf StrOrd <> "" And StrWhere = "" Then
    StrSql = StrSel & StrOrd
ElseIf StrOrd = "" And StrWhere <> "" Then
    StrSql = StrSel & StrWhere
Else
    StrSql = StrSel
End If
End Function

Sub OrderLists(OrdIndex As Integer)
rsord.MoveFirst
    Do
        lstOrd(OrdIndex).AddItem rsord!OrdField
        rsord.MoveNext
    Loop Until rsord.EOF
End Sub


Sub LstOrdRS()
Dim X As Integer
Set rsord = New Recordset
With rsord
    .Fields.Append ("OrdField"), adVariant
    .Open
For X = 0 To lstSelected.ListCount - 1
    lstSelected.ListIndex = X
    .AddNew ("OrdField"), lstSelected
Next
End With
End Sub

Sub lstCritRS()

Dim X As Integer
Set rscrit = New Recordset
With rscrit
    .Fields.Append ("CritField"), adVariant
    .Open
For X = 0 To lstSelected.ListCount - 1
    lstSelected.ListIndex = X
    .AddNew ("CritField"), lstSelected
Next
End With
End Sub
Sub CritLists(CritIndex As Integer)
'list for the criteria items
rscrit.MoveFirst
    Do
        lstCrit(CritIndex).AddItem rscrit!CritField
        rscrit.MoveNext
    Loop Until rscrit.EOF
    
    lstOper(CritIndex).AddItem "="
    lstOper(CritIndex).AddItem "<>"
    lstOper(CritIndex).AddItem ">="
    lstOper(CritIndex).AddItem "<="
    lstOper(CritIndex).AddItem "Like"

End Sub


Private Sub modGetWord()
'opens the word object
    On Error Resume Next
    Set objWrd = GetObject(, "Word.Application")

    If Err.Number <> 0 Then
        mboolWordRunning = True
    Else
        mboolWordRunning = False
    End If

    Err.Clear
'    Call modWordClass

    If mboolWordRunning = True Then
        Set objWrd = New Word.Application
    End If
End Sub

Private Sub txtCrit_LostFocus(Index As Integer)
If lstOper(Index) = "Like" Then
    txtCrit(Index).Text = "'%" & txtCrit(Index).Text & "%'"
Else
    txtCrit(Index).Text = "'" & txtCrit(Index).Text & "'"
End If
End Sub


