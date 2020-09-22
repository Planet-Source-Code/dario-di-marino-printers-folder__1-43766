VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pfPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mstrServerName As String
Private mstrPrinterName As String
Private mstrShareName As String
Private mstrPortName As String
Private mstrDriverName As String
Private mstrComment As String
Private mstrLocation As String
' Private mDevMode As DEVMODE
Private mstrSepFile As String
Private mstrPrintProcessor As String
Private mstrDatatype As String
Private mstrParameters As String
Private mstrDescription As String
' Private mSecurityDescriptor As SECURITY_DESCRIPTOR
Private mlngAttributes As Long
Private mlngPriority As Long
Private mlngDefaultPriority As Long
Private mlngStartTime As Long
Private mlngUntilTime As Long
Private mlngStatus As Long
Private mlngcJobs As Long
Private mlngAveragePPM As Long
Private mlngFlags As Long
Private mstrIPHost As String
Public Property Get StatusDesc() As String
   Dim strStatusDesc As String
   
   Select Case mlngStatus
      Case enm_PRINTER_STATUS.PS_READY: strStatusDesc = "Ready"
      Case enm_PRINTER_STATUS.PS_BUSY: strStatusDesc = "Busy"
      Case enm_PRINTER_STATUS.PS_DOOR_OPEN: strStatusDesc = "Door open"
      Case enm_PRINTER_STATUS.PS_INITIALIZING: strStatusDesc = "Initializing"
      Case enm_PRINTER_STATUS.PS_MANUAL_FEED: strStatusDesc = "Manual feed"
      Case enm_PRINTER_STATUS.PS_NO_TONER: strStatusDesc = "No toner"
      Case enm_PRINTER_STATUS.PS_NOT_AVAILABLE: strStatusDesc = "Not avaible"
      Case enm_PRINTER_STATUS.PS_OFFLINE: strStatusDesc = "Off line"
      Case enm_PRINTER_STATUS.PS_PAPER_PROBLEM: strStatusDesc = "Paper problem"
      Case enm_PRINTER_STATUS.PS_PAPER_OUT: strStatusDesc = "Paper empty"
      Case enm_PRINTER_STATUS.PS_PAUSED: strStatusDesc = "Paused"
      Case enm_PRINTER_STATUS.PS_PRINTING: strStatusDesc = "Printing"
      Case enm_PRINTER_STATUS.PS_PROCESSING: strStatusDesc = "Processing"
      Case enm_PRINTER_STATUS.PS_TONER_LOW: strStatusDesc = "Low toner"
      Case enm_PRINTER_STATUS.PS_USER_INTERVENTION:  strStatusDesc = "User intervention"
      Case enm_PRINTER_STATUS.PS_WAITING: strStatusDesc = "Waiting"
      Case enm_PRINTER_STATUS.PS_WARMING_UP: strStatusDesc = "Warming up"
      Case Else: strStatusDesc = "Unknow error"
   End Select
   
   StatusDesc = mlngStatus & " " & strStatusDesc
End Property
Public Property Get ServerName() As String
   ServerName = mstrServerName
End Property
Public Property Let ServerName(ByVal Nuovo_ServerName As String)
   mstrServerName = Nuovo_ServerName
End Property
Public Property Get PrinterName() As String
   PrinterName = mstrPrinterName
End Property
Public Property Get ShareName() As String
   ShareName = mstrShareName
End Property
Public Property Let ShareName(ByVal Nuovo_ShareName As String)
   mstrShareName = Nuovo_ShareName
End Property
Public Property Get Comment() As String
   Comment = mstrComment
End Property
Public Property Get Location() As String
   Location = mstrLocation
End Property
Public Property Get SepFile() As String
   SepFile = mstrSepFile
End Property
Public Property Let SepFile(ByVal Nuovo_SepFile As String)
   mstrSepFile = Nuovo_SepFile
End Property
Public Property Get PrintProcessor() As String
   PrintProcessor = mstrPrintProcessor
End Property
Public Property Let PrintProcessor(ByVal Nuovo_PrintProcessor As String)
   mstrPrintProcessor = Nuovo_PrintProcessor
End Property
Public Property Get Datatype() As String
   Datatype = mstrDatatype
End Property
Public Property Get Parameters() As String
   Parameters = mstrParameters
End Property
Public Property Get Description() As String
   Description = mstrDescription
End Property
Public Property Get Attributes() As Long
   Attributes = mlngAttributes
End Property
Public Property Get Priority() As Long
   Priority = mlngPriority
End Property
Public Property Let Priority(ByVal Nuovo_Priority As Long)
   mlngPriority = Nuovo_Priority
End Property
Public Property Get DefaultPriority() As Long
   DefaultPriority = mlngDefaultPriority
End Property
Public Property Let DefaultPriority(ByVal Nuovo_DefaultPriority As Long)
   mlngDefaultPriority = Nuovo_DefaultPriority
End Property
Public Property Get StartTime() As Long
   StartTime = mlngStartTime
End Property
Public Property Let StartTime(ByVal Nuovo_StartTime As Long)
   mlngStartTime = Nuovo_StartTime
End Property
Public Property Let UntilTime(ByVal Nuovo_UntilTime As Long)
   mlngUntilTime = Nuovo_UntilTime
End Property
Public Property Get UntilTime() As Long
   UntilTime = mlngUntilTime
End Property
Public Property Get Status() As Long
   Status = mlngStatus
End Property
Public Property Let Status(ByVal Nuovo_Status As Long)
   mlngStatus = Nuovo_Status
End Property
Public Property Get Jobs() As Long
   Jobs = mlngcJobs
End Property
Public Property Get AveragePPM() As Long
   AveragePPM = mlngAveragePPM
End Property
Public Property Get flags() As Long
   flags = mlngFlags
End Property
Public Property Get PortName() As String
   PortName = mstrPortName
End Property
Public Property Get DriverName() As String
   DriverName = mstrDriverName
End Property
Public Property Let Attributes(ByVal Nuovo_Attributes As Long)
   mlngAttributes = Nuovo_Attributes
End Property
Public Property Let Description(ByVal Nuovo_Description As String)
   mstrDescription = Nuovo_Description
End Property
Public Property Let DriverName(ByVal Nuovo_DriverName As String)
   mstrDriverName = Nuovo_DriverName
End Property
Public Property Let Location(ByVal Nuovo_Location As String)
   mstrLocation = Nuovo_Location
End Property
Public Property Let Parameters(ByVal Nuovo_Parameters As String)
   mstrParameters = Nuovo_Parameters
End Property
Public Property Let PrinterName(ByVal Nuovo_PrinterName As String)
   mstrPrinterName = Nuovo_PrinterName
End Property
Public Property Let PortName(ByVal Nuovo_PortName As String)
   mstrPortName = Nuovo_PortName
End Property
Public Property Let flags(ByVal Nuovo_Flags As Long)
   mlngFlags = Nuovo_Flags
End Property
Public Property Let Jobs(ByVal Nuovo_Jobs As Long)
   mlngcJobs = Nuovo_Jobs
End Property
Public Property Let Comment(ByVal Nuovo_Comment As String)
   mstrComment = Nuovo_Comment
End Property
Public Property Let Datatype(ByVal Nuovo_Datatype As String)
   mstrDatatype = Nuovo_Datatype
End Property
Public Property Let AveragePPM(ByVal Nuovo_AveragePPM As Long)
   mlngAveragePPM = Nuovo_AveragePPM
End Property



Public Sub GetInfoPrinter2()
   EnumeratePrinters2 Me
End Sub
