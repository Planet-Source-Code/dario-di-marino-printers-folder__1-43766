VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMonitor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Monitor"
   ClientHeight    =   5520
   ClientLeft      =   3090
   ClientTop       =   2550
   ClientWidth     =   6945
   LinkTopic       =   "frmMonitor"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   6945
   Begin VB.Timer tmrRefreshJobs 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   7200
      Top             =   1440
   End
   Begin VB.Timer tmrRefreshPrinter 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   7200
      Top             =   720
   End
   Begin MSComctlLib.ListView lstJobs 
      Height          =   3015
      Left            =   0
      TabIndex        =   1
      Top             =   2520
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   5318
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Document name"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Owner"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Pages"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Kb to Send"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Job Id"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lstPrinters 
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   4471
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Printer Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Jobs"
         Object.Width           =   1482
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Status"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Port"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Driver Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Description"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Comment"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   tmrRefreshPrinter.Enabled = True
End Sub


Private Sub lstPrinters_ItemClick(ByVal Item As MSComctlLib.ListItem)
   Dim lngIdx As Long
   
   ' I Refresh the printer data
   lngIdx = CLng(Mid$(Item.Key, 2))
   plstPrinter(lngIdx).GetInfoPrinter2

   Item.SubItems(1) = plstPrinter(lngIdx).Jobs
   Item.SubItems(2) = plstPrinter(lngIdx).StatusDesc
   Item.SubItems(3) = plstPrinter(lngIdx).PortName
   Item.SubItems(4) = plstPrinter(lngIdx).DriverName
   Item.SubItems(5) = plstPrinter(lngIdx).Description
   Item.SubItems(6) = plstPrinter(lngIdx).Comment
   
   lstJobs.ListItems.Clear
End Sub


Private Sub tmrRefreshJobs_Timer()
   Dim lngIdx As Long
   Dim objListItem As ListItem
      
   If lstPrinters.SelectedItem.Key = "" Then Exit Sub
   ' I Refresh the printer data
   lngIdx = CLng(Mid$(lstPrinters.SelectedItem.Key, 2))
   
   ' Pupulate List Printers
   EnumerateJobs2 plstPrinter(lngIdx)
         
   ' Now I populate form list
   On Error GoTo ErrorArray
   If UBound(plstJob) >= 0 Then
      For lngIdx = LBound(plstJob) To UBound(plstJob)
         If lngIdx > (lstJobs.ListItems.Count - 1) Then
            Set objListItem = lstJobs.ListItems.Add
         Else
            Set objListItem = lstJobs.ListItems.Item(lngIdx + 1)
         End If
            objListItem.Text = plstJob(lngIdx).Document
            objListItem.SubItems(1) = plstJob(lngIdx).StatusDesc
            objListItem.SubItems(2) = plstJob(lngIdx).UserName
            objListItem.SubItems(3) = plstJob(lngIdx).TotalPages
            objListItem.SubItems(4) = Format$(plstJob(lngIdx).Size / 1024, "0.00") & "K"
            objListItem.SubItems(5) = plstJob(lngIdx).JobId
            
            objListItem.EnsureVisible
            lstPrinters.Refresh
      Next lngIdx
   End If
   
   Set objListItem = Nothing
ErrorArray:
End Sub

Private Sub tmrRefreshPrinter_Timer()
   Dim lngIdx As Long
   Dim objListItem As ListItem
   
   tmrRefreshPrinter.Enabled = False
   
   ' Pupulate List Printers
   EnumeratePrinters1
      
   ' Now I populate form list
   For lngIdx = LBound(plstPrinter) To UBound(plstPrinter)
      plstPrinter(lngIdx).GetInfoPrinter2
      
      Set objListItem = lstPrinters.ListItems.Add(, "K" & lngIdx, plstPrinter(lngIdx).PrinterName)
         objListItem.SubItems(1) = plstPrinter(lngIdx).Jobs
         objListItem.SubItems(2) = plstPrinter(lngIdx).StatusDesc
         objListItem.SubItems(3) = plstPrinter(lngIdx).PortName
         objListItem.SubItems(4) = plstPrinter(lngIdx).DriverName
         objListItem.SubItems(5) = plstPrinter(lngIdx).Description
         objListItem.SubItems(6) = plstPrinter(lngIdx).Comment
         
         objListItem.EnsureVisible
         lstPrinters.Refresh
   Next lngIdx
   
   Debug.Print lstPrinters.SelectedItem.Key
   
   Set objListItem = Nothing
   
   tmrRefreshJobs.Enabled = True
End Sub


