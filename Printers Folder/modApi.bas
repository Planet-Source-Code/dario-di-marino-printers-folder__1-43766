Attribute VB_Name = "modAPI"
Option Explicit

' Declaration for printer
Private Const PRINTER_ENUM_CONNECTIONS = &H4
Private Const PRINTER_ENUM_LOCAL = &H2
Private Const PRINTER_CONTROL_RESUME = 2
Private Const PRINTER_CONTROL_SET_STATUS = 4
Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32
Private Const PRINTER_ATTRIBUTE_WORK_OFFLINE = &H400

Private Const ERROR_ACCESS_DENIED = 5         ' Access is denied.
Private Const ERROR_BAD_NETPATH = 53          ' The network path was not found
Private Const ERROR_UNEXP_NET_ERR = 59        ' An unexpected network error occurred
Private Const ERROR_INSUFFICIENT_BUFFER = 122 ' The data area passed to a system call is too small
Private Const ERROR_NETWORK_UNREACHABLE = 1231 ' The remote network is not reachable by the transport
Private Const RPC_S_SERVER_UNAVAILABLE = 1722 ' The RPC server is unavailable
Private Const RPC_S_CALL_FAILED = 1726        ' The remote procedure call failed

' Original API. I have changed pPrinterEnum and pJob in long
' Public Declare Function EnumPrinters Lib "winspool.drv" Alias "EnumPrintersA" (ByVal flags As Long, ByVal name As String, ByVal Level As Long, pPrinterEnum As Byte, ByVal cdBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
' Public Declare Function EnumJobs Lib "winspool.drv" Alias "EnumJobsA" (ByVal hPrinter As Long, ByVal FirstJob As Long, ByVal NoJobs As Long, ByVal Level As Long, pJob As Byte, ByVal cdBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
Private Declare Function EnumPrinters Lib "winspool.drv" Alias "EnumPrintersA" (ByVal flags As Long, ByVal name As String, ByVal Level As Long, pPrinterEnum As Long, ByVal cdBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
Private Declare Function EnumJobs Lib "winspool.drv" Alias "EnumJobsA" (ByVal hPrinter As Long, ByVal FirstJob As Long, ByVal NoJobs As Long, ByVal Level As Long, pJob As Long, ByVal cdBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
Private Declare Function PtrToStr Lib "kernel32" Alias "lstrcpyA" (ByVal RetVal As String, ByVal Ptr As Long) As Long
Private Declare Function StrLen Lib "kernel32" Alias "lstrlenA" (ByVal Ptr As Long) As Long
Private Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, pDefault As PRINTER_DEFAULTS) As Long
Private Declare Function GetPrinter Lib "winspool.drv" Alias "GetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Any, ByVal cbBuf As Long, pcbNeeded As Long) As Long
Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function SetPrinter Lib "winspool.drv" Alias "SetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Any, ByVal Command As Long) As Long

Private Declare Function GetLastError Lib "kernel32" () As Long

Private Type DEVMODE
   dmDeviceName As String * CCHDEVICENAME
   dmSpecVersion As Integer
   dmDriverVersion As Integer
   dmSize As Integer
   dmDriverExtra As Integer
   dmFields As Long
   dmOrientation As Integer
   dmPaperSize As Integer
   dmPaperLength As Integer
   dmPaperWidth As Integer
   dmScale As Integer
   dmCopies As Integer
   dmDefaultSource As Integer
   dmPrintQuality As Integer
   dmColor As Integer
   dmDuplex As Integer
   dmYResolution As Integer
   dmTTOption As Integer
   dmCollate As Integer
   dmFormName As String * CCHFORMNAME
   dmUnusedPadding As Integer
   dmBitsPerPel As Long
   dmPelsWidth As Long
   dmPelsHeight As Long
   dmDisplayFlags As Long
   dmDisplayFrequency As Long
End Type

Private Type PRINTER_DEFAULTS
   pDatatype As String
   pDevMode As DEVMODE
   DesiredAccess As Long
End Type

Public Enum enm_PRINTER_STATUS
   PS_READY = &H0
   PS_BUSY = &H200
   PS_DOOR_OPEN = &H400000
   PS_ERROR = &H2
   PS_INITIALIZING = &H8000
   PS_IO_ACTIVE = &H100
   PS_MANUAL_FEED = &H20
   PS_NO_TONER = &H40000
   PS_NOT_AVAILABLE = &H1000
   PS_OFFLINE = &H80
   PS_OUT_OF_MEMORY = &H200000
   PS_OUTPUT_BIN_FULL = &H800
   PS_PAGE_PUNT = &H80000
   PS_PAPER_JAM = &H8
   PS_PAPER_OUT = &H10
   PS_PAPER_PROBLEM = &H40
   PS_PAUSED = &H1
   PS_PENDING_DELETION = &H4
   PS_PRINTING = &H400
   PS_PROCESSING = &H4000
   PS_TONER_LOW = &H20000
   PS_USER_INTERVENTION = &H100000
   PS_WAITING = &H2000
   PS_WARMING_UP = &H10000
End Enum

Public Enum enm_JOB_STATUS
   JS_PAUSED = &H1
   JB_ERROR = &H2
   JS_DELETING = &H4
   JS_SPOOLING = &H8
   JS_PRINTING = &H10
   JS_OFFLINE = &H20
   JS_PAPEROUT = &H40
   JS_PRINTED = &H80
   JS_USER_INTERVENTION = &H10000
End Enum
Public Sub EnumerateJobs2(ByRef objpfPrinter As pfPrinter)
   ' Enumerate job with JOB_INFO_2
   Dim blnSuccess As Boolean
   Dim hPrinter As Long
   Dim udtDefault As PRINTER_DEFAULTS
   Dim lngEntries As Long
   Dim lngBufferRequired As Long
   Dim lngBufferSize As Long
   Dim lngBuffer() As Long
   Dim lngLastError As Long
   Dim lngNumJobs As Long
   Dim lngIdx  As Long
   Dim strBuffer As String
   Const pi = 26
      
   ' I need to open printer for get pinter's jobs
   blnSuccess = OpenPrinter(objpfPrinter.PrinterName, hPrinter, udtDefault)
   lngLastError = GetLastError
   
   If Not blnSuccess Then
      ' see documentation for error in lngLastError
   Else
      lngBufferSize = 3072 'This size may be suffucient
      lngNumJobs = 10
      ReDim lngBuffer(lngBufferSize \ 4 - 1) As Long 'Long is 4 bytes
      blnSuccess = EnumJobs(hPrinter, 0, lngNumJobs, 2, lngBuffer(0), lngBufferSize, lngBufferRequired, lngEntries)
      lngLastError = GetLastError
      
      If lngLastError = 122 Then ' ERROR_INSUFFICIENT_BUFFER
         lngBufferSize = lngBufferRequired
         ReDim lngBuffer(lngBufferSize \ 4 - 1) As Long 'Long is 4 bytes
         blnSuccess = EnumJobs(hPrinter, 0, lngNumJobs, 2, lngBuffer(0), lngBufferSize, lngBufferRequired, lngEntries)
         lngLastError = GetLastError ' See documentation
      End If
      
      If blnSuccess And lngEntries > 0 Then
         ReDim plstJob(lngEntries - 1)
         
         For lngIdx = 0 To lngEntries - 1
            Set plstJob(lngIdx) = New pfJob
            plstJob(lngIdx).JobId = lngBuffer(pi * lngIdx)
            
            strBuffer = Space$(StrLen(lngBuffer(lngIdx * pi + 1)))
            PtrToStr strBuffer, lngBuffer(lngIdx * pi + 1)
            plstJob(lngIdx).PrinterName = strBuffer
               
            strBuffer = Space$(StrLen(lngBuffer(lngIdx * pi + 2)))
            PtrToStr strBuffer, lngBuffer(lngIdx * pi + 2)
            plstJob(lngIdx).MachineName = strBuffer
               
            strBuffer = Space$(StrLen(lngBuffer(lngIdx * pi + 3)))
            PtrToStr strBuffer, lngBuffer(lngIdx * pi + 3)
            plstJob(lngIdx).UserName = strBuffer
         
            strBuffer = Space$(StrLen(lngBuffer(lngIdx * pi + 4)))
            PtrToStr strBuffer, lngBuffer(lngIdx * pi + 4)
            plstJob(lngIdx).Document = strBuffer
         
            strBuffer = Space$(StrLen(lngBuffer(lngIdx * pi + 5)))
            PtrToStr strBuffer, lngBuffer(lngIdx * pi + 5)
            plstJob(lngIdx).NotifyName = strBuffer
         
            strBuffer = Space$(StrLen(lngBuffer(lngIdx * pi + 6)))
            PtrToStr strBuffer, lngBuffer(lngIdx * pi + 6)
            plstJob(lngIdx).DataType = strBuffer
            
            strBuffer = Space$(StrLen(lngBuffer(lngIdx * pi + 7)))
            PtrToStr strBuffer, lngBuffer(lngIdx * pi + 7)
            plstJob(lngIdx).PrintProcessor = strBuffer
         
            strBuffer = Space$(StrLen(lngBuffer(lngIdx * pi + 8)))
            PtrToStr strBuffer, lngBuffer(lngIdx * pi + 8)
            plstJob(lngIdx).Parameters = strBuffer
         
            strBuffer = Space$(StrLen(lngBuffer(lngIdx * pi + 9)))
            PtrToStr strBuffer, lngBuffer(lngIdx * pi + 9)
            plstJob(lngIdx).DriverName = strBuffer
         
            'DEVMODE = 10
            
            strBuffer = Space$(StrLen(lngBuffer(lngIdx * pi + 11)))
            PtrToStr strBuffer, lngBuffer(lngIdx * pi + 11)
            plstJob(lngIdx).StatusDesc = strBuffer
         
            ' SECURITY_DESCRIPTOR = 12
         
            plstJob(lngIdx).Status = lngBuffer(lngIdx * pi + 13)
            plstJob(lngIdx).Priority = lngBuffer(lngIdx * pi + 14)
            plstJob(lngIdx).Position = lngBuffer(lngIdx * pi + 15)
            plstJob(lngIdx).StartTime = lngBuffer(lngIdx * pi + 16)
            plstJob(lngIdx).UntilTime = lngBuffer(lngIdx * pi + 17)
            plstJob(lngIdx).TotalPages = lngBuffer(lngIdx * pi + 18)
            plstJob(lngIdx).Size = lngBuffer(lngIdx * pi + 19)
            
            ' SYSTEMTIME = 20 and is 16 bytes. A long is 2 byte so 20 + 4 are 24
            
            plstJob(lngIdx).Time = lngBuffer(lngIdx * pi + 24)
            plstJob(lngIdx).PagesPrinted = lngBuffer(lngIdx * pi + 25)
         Next lngIdx
      End If
      
      blnSuccess = ClosePrinter(hPrinter)
   End If
End Sub


Public Sub EnumeratePrinters1()
   ' Enumerate printer with PRINTER_INFO_1
   Dim blnSuccess As Boolean
   Dim lngBufferRequired As Long
   Dim lngBufferSize As Long
   Dim lngBuffer() As Long
   Dim lngEntries As Long
   Dim lngLastError As Long
   Dim lngIdx  As Long
   Dim strBuffer As String
   Const pi = 4
   
   lngBufferSize = 3072 'This size may be suffucient
   ReDim lngBuffer(lngBufferSize \ 4 - 1) As Long 'Long is 4 bytes
   
   blnSuccess = EnumPrinters(PRINTER_ENUM_CONNECTIONS Or PRINTER_ENUM_LOCAL, vbNullString, _
               1, lngBuffer(0), lngBufferSize, lngBufferRequired, lngEntries)
   lngLastError = GetLastError
      
   If lngLastError = ERROR_INSUFFICIENT_BUFFER Then ' ERROR_INSUFFICIENT_BUFFER
      lngBufferSize = lngBufferRequired
      ReDim lngBuffer(lngBufferSize \ 4 - 1) As Long 'Long is 4 bytes
         
      blnSuccess = EnumPrinters(PRINTER_ENUM_CONNECTIONS Or PRINTER_ENUM_LOCAL, vbNullString, _
                   1, lngBuffer(0), lngBufferSize, lngBufferRequired, lngEntries)
      
      lngLastError = GetLastError ' See documentation
      If Not blnSuccess Then Exit Sub
   End If
   
   ReDim plstPrinter(lngEntries - 1)
   
   For lngIdx = 0 To lngEntries - 1
      Set plstPrinter(lngIdx) = New pfPrinter
      plstPrinter(lngIdx).flags = lngBuffer(pi * lngIdx)
      
      strBuffer = Space$(StrLen(lngBuffer(lngIdx * pi + 1)))
      PtrToStr strBuffer, lngBuffer(lngIdx * pi + 1)
      plstPrinter(lngIdx).Description = strBuffer
         
      strBuffer = Space$(StrLen(lngBuffer(lngIdx * pi + 2)))
      PtrToStr strBuffer, lngBuffer(lngIdx * pi + 2)
      plstPrinter(lngIdx).PrinterName = strBuffer
         
      strBuffer = Space$(StrLen(lngBuffer(lngIdx * pi + 3)))
      PtrToStr strBuffer, lngBuffer(lngIdx * pi + 3)
      plstPrinter(lngIdx).Comment = strBuffer
   Next lngIdx
End Sub

Public Sub EnumeratePrinters2(ByRef objpfPrinter As pfPrinter)
   ' Enumerate printer with PRINTER_INFO_2
   Dim blnSuccess As Boolean
   Dim lngBufferRequired As Long
   Dim lngBufferSize As Long
   Dim lngBuffer() As Long
   Dim lngEntries As Long
   Dim lngLastError As Long
   Dim strBuffer As String
   Dim hPrinter As Long
   Dim udtDefault As PRINTER_DEFAULTS
   Dim lngAttrib As Long
   Dim lngIdx As Long
   
   Dim lngNumJobs As Long
   Dim lngEntriesJobs As Long
   
   Const pi = 21

   lngBufferSize = 3072 'This size may be suffucient
   ReDim lngBuffer(lngBufferSize \ 4 - 1) As Long 'Long is 4 bytes
   
   ' I Can't use PRINTER_ENUM_NAME becuse it don't work with PRINTER_INFO_2
   blnSuccess = EnumPrinters(PRINTER_ENUM_CONNECTIONS Or PRINTER_ENUM_LOCAL, vbNullString, _
               2, lngBuffer(0), lngBufferSize, lngBufferRequired, lngEntries)
   lngLastError = GetLastError
      
   If Not blnSuccess Then
      If lngLastError = ERROR_INSUFFICIENT_BUFFER Then
         lngBufferSize = lngBufferRequired
         ReDim lngBuffer(lngBufferSize \ 4 - 1) As Long 'Long is 4 bytes
            
         blnSuccess = EnumPrinters(PRINTER_ENUM_CONNECTIONS Or PRINTER_ENUM_LOCAL, vbNullString, _
                      2, lngBuffer(0), lngBufferSize, lngBufferRequired, lngEntries)
            
         lngLastError = GetLastError ' See documentation
         If Not blnSuccess Then Exit Sub
      Else
         ' see documentation for error in lngLastError
         Exit Sub
      End If
   End If
   
   For lngIdx = 0 To lngEntries - 1
      ' Now I search the data of my printer
      strBuffer = Space$(StrLen(lngBuffer(pi * lngIdx + 1)))
      PtrToStr strBuffer, lngBuffer(pi * lngIdx + 1)
      If objpfPrinter.PrinterName = strBuffer Then
         With objpfPrinter
            strBuffer = Space$(StrLen(lngBuffer(pi * lngIdx)))
            PtrToStr strBuffer, lngBuffer(pi * lngIdx)
            .ServerName = strBuffer
            
            strBuffer = Space$(StrLen(lngBuffer(pi * lngIdx + 1)))
            PtrToStr strBuffer, lngBuffer(pi * lngIdx + 1)
            .PrinterName = strBuffer
               
            strBuffer = Space$(StrLen(lngBuffer(pi * lngIdx + 2)))
            PtrToStr strBuffer, lngBuffer(pi * lngIdx + 2)
            .ShareName = strBuffer
            
            strBuffer = Space$(StrLen(lngBuffer(pi * lngIdx + 3)))
            PtrToStr strBuffer, lngBuffer(pi * lngIdx + 3)
            .PortName = strBuffer
            
            strBuffer = Space$(StrLen(lngBuffer(pi * lngIdx + 4)))
            PtrToStr strBuffer, lngBuffer(pi * lngIdx + 4)
            .DriverName = strBuffer
               
            strBuffer = Space$(StrLen(lngBuffer(pi * lngIdx + 5)))
            PtrToStr strBuffer, lngBuffer(pi * lngIdx + 5)
            .Comment = strBuffer
               
            strBuffer = Space$(StrLen(lngBuffer(pi * lngIdx + 6)))
            PtrToStr strBuffer, lngBuffer(pi * lngIdx + 6)
            .Location = strBuffer
            
            'DEVMODE = 7
               
            strBuffer = Space$(StrLen(lngBuffer(pi * lngIdx + 8)))
            PtrToStr strBuffer, lngBuffer(pi * lngIdx + 8)
            .SepFile = strBuffer
            
            strBuffer = Space$(StrLen(lngBuffer(pi * lngIdx + 9)))
            PtrToStr strBuffer, lngBuffer(pi * lngIdx + 9)
            .PrintProcessor = strBuffer
            
            strBuffer = Space$(StrLen(lngBuffer(pi * lngIdx + 10)))
            PtrToStr strBuffer, lngBuffer(pi * lngIdx + 10)
            .DataType = strBuffer
                        
            strBuffer = Space$(StrLen(lngBuffer(pi * lngIdx + 11)))
            PtrToStr strBuffer, lngBuffer(pi * lngIdx + 11)
            .Parameters = strBuffer
            
            ' SECURITY_DESCRIPTOR = 12
               
            .Attributes = lngBuffer(pi * lngIdx + 13)
            .Priority = lngBuffer(pi * lngIdx + 14)
            .DefaultPriority = lngBuffer(pi * lngIdx + 15)
            .StartTime = lngBuffer(pi * lngIdx + 16)
            .UntilTime = lngBuffer(pi * lngIdx + 17)
            .Status = lngBuffer(pi * lngIdx + 18)
            .Jobs = lngBuffer(pi * lngIdx + 19)
            .AveragePPM = lngBuffer(pi * lngIdx + 20)
         End With
         Exit For
      End If
   Next lngIdx
   
   ' Now I have more information but if I want the current status of printer
   ' I need to open printer and after get PRINTER_INFO_2
   blnSuccess = OpenPrinter(objpfPrinter.PrinterName, hPrinter, udtDefault)
   lngLastError = GetLastError
   
   If Not blnSuccess Then
      objpfPrinter.Status = GetLastError
      ' see documentation for error in lngLastError
   Else
      lngBufferSize = 3072 'This size may be suffucient
      ReDim lngBuffer(lngBufferSize \ 4 - 1) As Long 'Long is 4 bytes
      blnSuccess = GetPrinter(hPrinter, 2, lngBuffer(0), lngBufferSize, lngBufferRequired)
      lngLastError = GetLastError
         
      If Not blnSuccess Then
         If lngLastError = ERROR_INSUFFICIENT_BUFFER Then
            lngBufferSize = lngBufferRequired
            ReDim lngBuffer(lngBufferSize \ 4 - 1) As Long 'Long is 4 bytes
            blnSuccess = GetPrinter(hPrinter, 2, lngBuffer(0), lngBufferSize, lngBufferRequired)
            lngLastError = GetLastError
         End If
      End If
      
      If Not blnSuccess Then
         If lngLastError = ERROR_BAD_NETPATH Or lngLastError = ERROR_UNEXP_NET_ERR Then
            ' Windows 2000 cant' access to printer, because pc is off, printer not exist or not is shared
            objpfPrinter.Status = PS_OFFLINE
         ElseIf lngLastError = ERROR_ACCESS_DENIED Then
            ' Windows 2000 cant' access to printer, because pc is off, printer not exist or not is shared or
            ' I haven't anymore the rights
            objpfPrinter.Status = PS_OFFLINE
         ElseIf lngLastError = RPC_S_SERVER_UNAVAILABLE Then
            objpfPrinter.Status = PS_OFFLINE
         ElseIf lngLastError = RPC_S_CALL_FAILED Then
            objpfPrinter.Status = PS_OFFLINE
         ElseIf lngLastError = ERROR_NETWORK_UNREACHABLE Then
            ' Are you sure to have LAN connection?
            objpfPrinter.Status = PS_OFFLINE
         Else
            objpfPrinter.Status = lngLastError
            ' see documentation for error in lngLastError
         End If
      End If
     
      If blnSuccess Then
         lngAttrib = lngBuffer(13)
         ' Generaly only for W95/98 I try to remove ATTRIBUTE_WORK_OFFLINE because if the printer
         ' is now on line, you must remove this flag manualy otherwise the printer not print
         If lngAttrib And PRINTER_ATTRIBUTE_WORK_OFFLINE Then
            lngBuffer(13) = lngBuffer(13) - PRINTER_ATTRIBUTE_WORK_OFFLINE
            SetPrinter hPrinter, 2, lngBuffer(0), 0
         End If
            
         With objpfPrinter
            strBuffer = Space$(StrLen(lngBuffer(0)))
            PtrToStr strBuffer, lngBuffer(0)
            .ServerName = strBuffer
            
            strBuffer = Space$(StrLen(lngBuffer(1)))
             PtrToStr strBuffer, lngBuffer(1)
            .PrinterName = strBuffer
               
            strBuffer = Space$(StrLen(lngBuffer(2)))
            PtrToStr strBuffer, lngBuffer(2)
            .ShareName = strBuffer
            
            strBuffer = Space$(StrLen(lngBuffer(3)))
            PtrToStr strBuffer, lngBuffer(3)
            .PortName = strBuffer
            
            strBuffer = Space$(StrLen(lngBuffer(4)))
            PtrToStr strBuffer, lngBuffer(4)
            .DriverName = strBuffer
               
            strBuffer = Space$(StrLen(lngBuffer(5)))
            PtrToStr strBuffer, lngBuffer(5)
            .Comment = strBuffer
               
            strBuffer = Space$(StrLen(lngBuffer(6)))
            PtrToStr strBuffer, lngBuffer(6)
            .Location = strBuffer
            
            'DEVMODE = 7
               
            strBuffer = Space$(StrLen(lngBuffer(8)))
            PtrToStr strBuffer, lngBuffer(8)
            .SepFile = strBuffer
            
            strBuffer = Space$(StrLen(lngBuffer(9)))
            PtrToStr strBuffer, lngBuffer(9)
            .PrintProcessor = strBuffer
            
            strBuffer = Space$(StrLen(lngBuffer(10)))
            PtrToStr strBuffer, lngBuffer(10)
            .DataType = strBuffer
                        
            strBuffer = Space$(StrLen(lngBuffer(11)))
            PtrToStr strBuffer, lngBuffer(11)
            .Parameters = strBuffer
            
            ' SECURITY_DESCRIPTOR = 12
               
            .Attributes = lngBuffer(13)
            .Priority = lngBuffer(14)
            .DefaultPriority = lngBuffer(15)
            .StartTime = lngBuffer(16)
            .UntilTime = lngBuffer(17)
            .Status = lngBuffer(18)
            .Jobs = lngBuffer(19)
            .AveragePPM = lngBuffer(20)
         End With
      
         ' Some printers return ready also when aren't ready. So I try to enum the jobs, and if
         ' the function retun an error, I am sure that the printer are not avaible
      
         lngBufferSize = 3072
         lngNumJobs = 2 'Total jobs to enumerate.
         ReDim lngBuffer(lngBufferSize \ 4 - 1) As Long 'Long è 4 bytes
         blnSuccess = EnumJobs(hPrinter, 0, lngNumJobs, 2, lngBuffer(0), lngBufferSize, lngBufferRequired, lngEntriesJobs)
         lngLastError = Err.LastDllError
      
         If Not blnSuccess Then
            If lngLastError = ERROR_INSUFFICIENT_BUFFER Then
               lngBufferSize = lngBufferRequired
               ReDim lngBuffer(lngBufferSize \ 4 - 1) As Long 'Long is 4 bytes
               blnSuccess = EnumJobs(hPrinter, 0, lngNumJobs, 2, lngBuffer(0), lngBufferSize, lngBufferRequired, lngEntriesJobs)
               lngLastError = Err.LastDllError
            End If
         
            If Not blnSuccess Then
               If lngLastError = ERROR_BAD_NETPATH Or lngLastError = ERROR_UNEXP_NET_ERR Then
                ' Windows 2000 cant' access to printer, because pc is off, printer not exist or not is shared
                  objpfPrinter.Status = PS_OFFLINE
               ElseIf lngLastError = ERROR_ACCESS_DENIED Then
                  ' Windows 2000 cant' access to printer, because pc is off, printer not exist or not is shared or
                  ' I haven't anymore the rights
                  objpfPrinter.Status = PS_OFFLINE
               ElseIf lngLastError = RPC_S_SERVER_UNAVAILABLE Then
                  objpfPrinter.Status = PS_OFFLINE
               ElseIf lngLastError = RPC_S_CALL_FAILED Then
                  objpfPrinter.Status = PS_OFFLINE
               ElseIf lngLastError = ERROR_NETWORK_UNREACHABLE Then
                  ' Are you sure to have LAN connection?
                  objpfPrinter.Status = PS_OFFLINE
               Else
                 objpfPrinter.Status = lngLastError
                 ' see documentation for error in lngLastError
               End If
            End If
         End If
      
      End If
      
      blnSuccess = ClosePrinter(hPrinter)
   End If
End Sub
