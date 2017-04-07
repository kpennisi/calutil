VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMain 
   Caption         =   "Print Utility for MS Outlook Calendars"
   ClientHeight    =   3912
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   7020
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3912
   ScaleWidth      =   7020
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrint 
      Caption         =   "  Print Calendar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   5160
      TabIndex        =   7
      ToolTipText     =   "Click here to preview the report."
      Top             =   1920
      Width           =   1275
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview Calendar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   3600
      TabIndex        =   4
      ToolTipText     =   "Click here to preview the report."
      Top             =   1920
      Width           =   1215
   End
   Begin VB.ComboBox cboCals 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Left            =   3300
      TabIndex        =   1
      ToolTipText     =   "Select the Outlook Folder that contains the Calendar."
      Top             =   1200
      Width           =   3435
   End
   Begin MSComCtl2.MonthView mvDateSel 
      Height          =   2256
      Left            =   240
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Select one or more days you want to include in your report."
      Top             =   780
      Width           =   2664
      _ExtentX        =   4699
      _ExtentY        =   3979
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      MaxSelCount     =   31
      MultiSelect     =   -1  'True
      StartOfWeek     =   22872065
      CurrentDate     =   36753
      MinDate         =   29221
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   432
      Left            =   240
      TabIndex        =   6
      Top             =   3240
      Width           =   6192
   End
   Begin VB.Image imgAbout 
      Height          =   264
      Left            =   6600
      Picture         =   "frmMain.frx":0442
      ToolTipText     =   "About ..."
      Top             =   3600
      Width           =   288
   End
   Begin VB.Label lblMain 
      Caption         =   "Select the day(s) to include in the report:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   2
      Left            =   300
      TabIndex        =   5
      Top             =   180
      Width           =   2232
   End
   Begin VB.Label lblMain 
      Caption         =   "** To select multiple days, hold down the shift key while selecting the from/to days."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   3720
      Width           =   6192
   End
   Begin VB.Label lblMain 
      Caption         =   "Select the calendar to report on:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Index           =   0
      Left            =   3300
      TabIndex        =   2
      Top             =   780
      Width           =   3252
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type OlkInfo
   path As String
   olMF As Outlook.MAPIFolder
End Type

Dim currLogin As String
Dim aOutlookPath() As OlkInfo
Dim olApp As Outlook.Application
Dim crApp As CRPEAuto.Application
Dim mDefCal As String
Dim mDefCalIdx As Integer
         
Private Sub printReport(doPrint As Boolean)
   Dim fDate As Date
   Dim tDate As Date
   Dim cID As Long
   Dim rpt As CRPEAuto.Report
   Dim crPrm As CRPEAuto.ParameterFieldDefinition
   
   On Error GoTo Err_Print
   
   lblStatus.Caption = "Opening Report File ..."
   DoEvents
   fDate = mvDateSel.SelStart
   tDate = mvDateSel.SelEnd
'   currLogin = aOutlookPath(cboCals.ListIndex).path
   
'   lblStatus.Caption = "Logging on to Exchange Server ..."
'   DoEvents
   Set crApp = Nothing
   Set crApp = New CRPEAuto.Application
   
'   crApp.LogOnServer "p2soutlk.dll", "Appointment", currLogin, "", ""
   Set rpt = crApp.OpenReport(App.path & "\" & "cal3col.rpt")
   
   rpt.Database.Tables(1).Location = App.path & "\calutil.mdb"
   Set crPrm = rpt.ParameterFields(1)
   crPrm.SetCurrentValue fDate, 10   '10
   Set crPrm = rpt.ParameterFields(2)
   crPrm.SetCurrentValue tDate, 10   'crDateField   '10
   Set crPrm = Nothing
   
   lblStatus.Caption = "Creating Report ..."
   DoEvents
   If (doPrint) Then
      rpt.PrintOut
   Else
      rpt.Preview
   End If
   lblStatus.Caption = ""
   Set rpt = Nothing
   Exit Sub
   
Err_Print:
   If (Err.Number <> 20545) Then    'cancelled printing
      lblStatus.Caption = "Error: " & Err.Description
   Else
      lblStatus.Caption = "Printing cancelled."
   End If
End Sub

Private Sub cboCals_Change()
   cmdPrint.SetFocus
End Sub

Private Sub cmdPreview_Click()
   Debug.Print "Start Time: " & Now
   Call doReport(False)
   Debug.Print "End Time: " & Now
End Sub

Private Sub cmdPrint_Click()
   Call doReport(True)
End Sub

Private Sub Form_Activate()
   mvDateSel.Value = Date
   cboCals.ListIndex = mDefCalIdx
End Sub

Private Sub Form_Load()
   mDefCalIdx = -1
   Call readRegData
   Call getFolders
   Unload frmWait
End Sub

Public Sub getFolders()
   
   On Error GoTo Err_GetFolders
   Set olApp = CreateObject("Outlook.Application")
   
   Erase aOutlookPath
   Call parseFolders(olApp.GetNamespace("MAPI").Folders)
   Exit Sub
   
Err_GetFolders:
   MsgBox "Error: Could not parse Outlook Folders", vbExclamation, "CalUtility"
   lblStatus.Caption = Err.Description
End Sub

Private Sub parseFolders(oFolders As Outlook.Folders)
   Dim olMF As Outlook.MAPIFolder
   Dim fldPath As String
   Static ii As Integer
   Static fullPath As String
   Dim calName As String
   
   On Error GoTo Err_ParseFolders
   
   'save all outlook folders in a table for our combobox to read
   For Each olMF In oFolders
       If (olMF.Class = olFolder) Then
         If (olMF.DefaultItemType = olAppointmentItem) Then
            'outlook will gen an error checking description if it is 0 len
            'so it needs to be non zero!
            calName = getFolderName(olMF.Name, olMF.Description)
            cboCals.AddItem calName
            If (calName = mDefCal) Then
               mDefCalIdx = cboCals.NewIndex
            End If
            ReDim Preserve aOutlookPath(ii)
            aOutlookPath(ii).path = fullPath & olMF.Name
            Set aOutlookPath(ii).olMF = olMF
            ii = ii + 1
         ElseIf (olMF.Folders.Count > 0) Then
            fldPath = fullPath
            fullPath = fullPath & olMF.Name & "@"
            'use recursion to go thru all folders
            Call parseFolders(olMF.Folders)
            'reset folder path to correct level
            fullPath = fldPath
         End If
      End If
   Next olMF
   Set olMF = Nothing
   Exit Sub
   
Err_ParseFolders:
   lblStatus.Caption = "Error: " & Err.Description
End Sub

Private Function getFolderName(oName As String, oDescr As String) As String
   'if this is the default cal, use description value to make unique
   If (StrComp(oName, "Calendar", vbTextCompare) = 0) Then
      getFolderName = oName & " - " & oDescr
   Else
      getFolderName = oName
   End If
End Function

Private Sub Form_Unload(Cancel As Integer)
   Call saveRegData
   Erase aOutlookPath
   Set olApp = Nothing
   Set crApp = Nothing
   Set frmMain = Nothing
End Sub

Private Sub imgAbout_Click()
   frmAbout.Show vbModal, Me
End Sub

Private Sub mvDateSel_DateClick(ByVal DateClicked As Date)
   cmdPrint.SetFocus
End Sub

Private Function getData() As Integer
   Dim myDB As dao.Database
   Dim olMF As Outlook.MAPIFolder
   Dim olAppt As Outlook.AppointmentItem
   Dim rItems As Outlook.Items
   Dim rs As Recordset
   Dim mySQL As String
   Dim eDate As Date
   Dim rCnt As Integer
   
   If (cboCals.ListIndex = -1) Then
      MsgBox "Please Select a Calendar", vbExclamation, "Print Calendar"
      Exit Function
   End If
   
   On Error GoTo Err_GetData
   Set myDB = OpenDatabase(App.path & "\calutil.mdb")
   mySQL = "DELETE * FROM tAppts"
   myDB.Execute mySQL, dbFailOnError
   
   Set rs = myDB.OpenRecordset("tAppts", dbOpenTable)
   
   Set olMF = aOutlookPath(cboCals.ListIndex).olMF
   
   'we will go to next day so we pull all events up till midnight for next day
   eDate = DateAdd("d", 1, mvDateSel.SelEnd)
   Set rItems = olMF.Items.Restrict("[Start]>= '" & mvDateSel.SelStart & "' AND [Start]< '" & eDate & "'")
   'this will generate an error if they don't have permission to view info
   rCnt = rItems.Count
   For Each olAppt In rItems
      rs.AddNew
      rs.Fields("ApptDate") = olAppt.Start
      rs.Fields("StartTime") = olAppt.Start
      rs.Fields("EndTime") = olAppt.End
      rs.Fields("Notes") = olAppt.Body
      rs.Fields("Subject") = olAppt.Subject
      rs.Fields("Location") = olAppt.Location
      rs.Update
   Next olAppt
   rs.Close
   Set rs = Nothing
   Set rItems = Nothing
   Set olAppt = Nothing
   Set olMF = Nothing
   Set myDB = Nothing
   getData = rCnt
   Exit Function
   
Err_GetData:
   lblStatus.Caption = "Error: " & Err.Description
   getData = -1
End Function

Private Sub readRegData()
   mDefCal = GetSetting(App.Title, "Settings", "Calendar", "")
End Sub

Private Sub saveRegData()
   Dim calName As String
   
   If (cboCals.ListIndex <> -1) Then
      calName = cboCals.List(cboCals.ListIndex)
   End If
   Call SaveSetting(App.Title, "Settings", "Calendar", calName)
End Sub

Private Sub doReport(bPrint As Boolean)
   Dim rc As Integer
   
   rc = getData
   If (rc > 0) Then
      Me.MousePointer = vbHourglass
      Call printReport(bPrint)
      Me.MousePointer = vbDefault
   ElseIf (rc = 0) Then
      lblStatus.Caption = "The calendar has no entries for those days."
   End If
End Sub
