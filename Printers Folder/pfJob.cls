VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pfJob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mlngJobId As Long
Private mstrPrinterName As String
Private mstrMachineName As String
Private mstrUserName As String
Private mstrDocument As String
Private mstrNotifyName As String
Private mstrDatatype As String
Private mstrPrintProcessor As String
Private mstrParameters As String
Private mstrDriverName As String
'Private mDevMode As DEVMODE
Private mstrStatusDesc As String
'Private mSecurityDescriptor As SECURITY_DESCRIPTOR
Private mlngStatus As Long
Private mlngPriority As Long
Private mlngPosition As Long
Private mlngStartTime As Long
Private mlngUntilTime As Long
Private mlngTotalPages As Long
Private mlngSize As Long
'Private Submitted As SYSTEMTIME
Private mlngTime As Long
Private mlngPagesPrinted As Long
Public Property Get JobId() As Long
   JobId = mlngJobId
End Property
Public Property Let JobId(ByVal Nuovo_JobId As Long)
   mlngJobId = Nuovo_JobId
End Property
Public Property Get PagesPrinted() As Long
   PagesPrinted = mlngPagesPrinted
End Property
Public Property Let PagesPrinted(ByVal Nuovo_PagesPrinted As Long)
   mlngPagesPrinted = Nuovo_PagesPrinted
End Property
Public Property Get Time() As Long
   Time = mlngTime
End Property
Public Property Let Time(ByVal Nuovo_Time As Long)
   mlngTime = Nuovo_Time
End Property
Public Property Get Size() As Long
   Size = mlngSize
End Property
Public Property Let Size(ByVal Nuovo_Size As Long)
   mlngSize = Nuovo_Size
End Property
Public Property Get TotalPages() As Long
   TotalPages = mlngTotalPages
End Property
Public Property Let TotalPages(ByVal Nuovo_TotalPages As Long)
   mlngTotalPages = Nuovo_TotalPages
End Property
Public Property Get UntilTime() As Long
   UntilTime = mlngUntilTime
End Property
Public Property Let UntilTime(ByVal Nuovo_UntilTime As Long)
   mlngUntilTime = Nuovo_UntilTime
End Property
Public Property Get StartTime() As Long
   StartTime = mlngStartTime
End Property
Public Property Let StartTime(ByVal Nuovo_StartTime As Long)
   mlngStartTime = Nuovo_StartTime
End Property
Public Property Get Position() As Long
   Position = mlngPosition
End Property
Public Property Let Position(ByVal Nuovo_Position As Long)
   mlngPosition = Nuovo_Position
End Property
Public Property Get Priority() As Long
   Priority = mlngPriority
End Property
Public Property Let Priority(ByVal Nuovo_Priority As Long)
   mlngPriority = Nuovo_Priority
End Property
Public Property Get Status() As Long
   Status = mlngStatus
End Property
Public Property Let Status(ByVal Nuovo_Status As Long)
   mlngStatus = Nuovo_Status
End Property
Public Property Get StatusDesc() As String
   If mstrStatusDesc <> "" Then
      StatusDesc = mstrStatusDesc
   Else
      Select Case mlngStatus
         Case JS_PAUSED: StatusDesc = "Paused"
         Case JB_ERROR: StatusDesc = "Error"
         Case JS_DELETING: StatusDesc = "Deleting"
         Case JS_SPOOLING: StatusDesc = "Spooling"
         Case JS_PRINTING: StatusDesc = "Printing"
         Case JS_OFFLINE: StatusDesc = "Off line"
         Case JS_PAPEROUT: StatusDesc = "Paper out"
         Case JS_PRINTED: StatusDesc = "Printed"
         Case JS_USER_INTERVENTION: StatusDesc = "User intervention"
      End Select
   End If
End Property
Public Property Let StatusDesc(ByVal Nuovo_StatusDesc As String)
   mstrStatusDesc = Nuovo_StatusDesc
End Property
Public Property Get DriverName() As String
   DriverName = mstrDriverName
End Property
Public Property Let DriverName(ByVal Nuovo_DriverName As String)
   mstrDriverName = Nuovo_DriverName
End Property
Public Property Get Parameters() As String
   Parameters = mstrParameters
End Property
Public Property Let Parameters(ByVal Nuovo_Parameters As String)
   mstrParameters = Nuovo_Parameters
End Property
Public Property Get PrintProcessor() As String
   PrintProcessor = mstrPrintProcessor
End Property
Public Property Let PrintProcessor(ByVal Nuovo_PrintProcessor As String)
   mstrPrintProcessor = Nuovo_PrintProcessor
End Property
Public Property Get DataType() As String
   DataType = mstrDatatype
End Property
Public Property Let DataType(ByVal Nuovo_DataType As String)
   mstrDatatype = Nuovo_DataType
End Property
Public Property Get NotifyName() As String
   NotifyName = mstrNotifyName
End Property
Public Property Let NotifyName(ByVal Nuovo_NotifyName As String)
   mstrNotifyName = Nuovo_NotifyName
End Property
Public Property Get Document() As String
   Document = mstrDocument
End Property
Public Property Let Document(ByVal Nuovo_Document As String)
   mstrDocument = Nuovo_Document
End Property
Public Property Get UserName() As String
   UserName = mstrUserName
End Property
Public Property Let UserName(ByVal Nuovo_UserName As String)
   mstrUserName = Nuovo_UserName
End Property
Public Property Get MachineName() As String
   MachineName = mstrMachineName
End Property
Public Property Let MachineName(ByVal Nuovo_MachineName As String)
   mstrMachineName = Nuovo_MachineName
End Property
Public Property Get PrinterName() As String
   PrinterName = mstrPrinterName
End Property
Public Property Let PrinterName(ByVal Nuovo_PrinterName As String)
   mstrPrinterName = Nuovo_PrinterName
End Property

