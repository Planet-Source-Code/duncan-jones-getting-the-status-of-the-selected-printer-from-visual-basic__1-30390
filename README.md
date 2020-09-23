<div align="center">

## Getting the status of the selected printer from Visual Basic


</div>

### Description

Shows how you can use the Windows API to return additional information about the printer above and beyond that which is available through the Visual Basic Printer object.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Duncan Jones](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/duncan-jones.md)
**Level**          |Advanced
**User Rating**    |3.7 (11 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/duncan-jones-getting-the-status-of-the-selected-printer-from-visual-basic__1-30390/archive/master.zip)





### Source Code

<p align=left><font face="Arial" style="BACKGROUND-COLOR: yellow">&nbsp;1. What the <i>Printer</i> object missed</font></p>
   <p align=justify>&nbsp;<font face="Arial" >Printing has long been a very problematic part of developing complete
   and professional applications in Visual Basic. This was redressed to a large degree with the new <i>Printer</i>
   object introduced in Visual Basic 4.<br>
   &nbsp;However, there are shortcomings with this object. The biggest shortcoming is that you cannot find out whether
   the printer is ready, busy, out of paper etc. from your application.<br>
   However, there is an API call, <i>GetPrinter</i> which returns a great deal more information about a printer.
   </font></p>
   <!-- API Declaration -->
   <p align=left class="sourcecode">
   <font class="keyword">
   Private Declare Function</font> GetPrinterApi <font class="keyword">Lib</font> "winspool.drv" <font class="keyword">Alias</font> _ <br>
   &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"GetPrinterA" <font class="keyword">(ByVal</font> hPrinter <font class="keyword">As Long,</font> _ <br>
   &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font class="keyword">ByVal</font> Level <font class="keyword">As Long,</font> _ <br>
   &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;buffer <font class="keyword">As Long,</font> _ <br>
   &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font class="keyword">ByVal</font> pbSize <font class="keyword">As Long,</font> _ <br>
   &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;pbSizeNeeded <font class="keyword">As Long) As Long
   </font>
   </p>
   <p align=justify>&nbsp;<font face="Arial" >This takes the handle to a printer in <i>hPrinter</i> and fills
   the buffer provided to it with information from the printer driver. To get the handle from the Printer object,
   you need to use the <i>OpenPrinter</i> API call. <br>
   &nbsp;This handle must be released using the <i>ClosePrinter</i> API call as soon as you are finished with it.
   </font></p>
   <!-- API Declaration -->
   <p align=left class="sourcecode">
   <font class="keyword">
   Private Type</font> PRINTER_DEFAULTS <br>
   &nbsp;&nbsp;pDatatype <font class="keyword">As String</font> <br>
   &nbsp;&nbsp;pDevMode <font class="keyword">As DEVMODE</font> <br>
   &nbsp;&nbsp;DesiredAccess <font class="keyword">As Long</font> <br>
   <font class="keyword">End Type</font> <br>
   <br>
   <font class="keyword">Private Declare Function</font> OpenPrinter <font class="keyword">Lib</font> "winspool.drv" _ <br>
   &nbsp;&nbsp;&nbsp;&nbsp;<font class="keyword">Alias</font> "OpenPrinterA" <font class="keyword">(ByVal</font> pPrinterName <font class="keyword">As String,</font> _ <br>
   &nbsp;&nbsp;&nbsp;&nbsp;phPrinter <font class="keyword">As Long</font>, pDefault <font class="keyword">As</font> PRINTER_DEFAULTS) As Long <br>
   <br>
   <font class="keyword">Private Declare Function</font> ClosePrinter <font class="keyword">Lib</font> "winspool.drv" _ <br>
   &nbsp;&nbsp;&nbsp;&nbsp;<font class="keyword">(ByVal</font> hPrinter <font class="keyword">As Long) As Long</font> <br>
   </p>
   <p align=justify>&nbsp;<font face="Arial" >You pass the Printer.DeviceName to this to get the handle.
   </font></p>
   <!-- Use -->
   <p align=left class="sourcecode">
   <font class="keyword">
    Dim</font> lret <font class="keyword">As Long</font> <br>
    <font class="keyword">Dim</font> pDef <font class="keyword">As</font> PRINTER_DEFAULTS <br>
    <br>
    lret = OpenPrinter(<font class="keyword">Printer.DeviceName</font>, mhPrinter, pDef)
   </font>
   </p>
   <p align=left><font face="Arial" style="BACKGROUND-COLOR: yellow">&nbsp;2. The different statuses</font></p>
   <p align=justify>&nbsp;<font face="Arial" >There are a number of standard statuses that can be returned by the
   printer driver.
   </font></p>
   <!-- Enumerated type -->
   <p align=left class="sourcecode">
   <font class="keyword">
   Public Enum</font> <a href="http://www.merrioncomputing.com/EventVB/Printer_Status.html">Printer_Status</a> <br>
   &nbsp;&nbsp;&nbsp;PRINTER_STATUS_READY = &amp;H0 <br>
   &nbsp;&nbsp;&nbsp;PRINTER_STATUS_PAUSED = &amp;H1 <br>
   &nbsp;&nbsp;&nbsp;PRINTER_STATUS_ERROR = &amp;H2 <br>&nbsp;&nbsp;&nbsp;PRINTER_STATUS_PENDING_DELETION =
    &amp;H4 <br>
   &nbsp;&nbsp;&nbsp;PRINTER_STATUS_PAPER_JAM = &amp;H8 <br>&nbsp;&nbsp;&nbsp;PRINTER_STATUS_PAPER_OUT =
    &amp;H10 <br>&nbsp;&nbsp;&nbsp;PRINTER_STATUS_MANUAL_FEED =
    &amp;H20 <br>&nbsp;&nbsp;&nbsp;PRINTER_STATUS_PAPER_PROBLEM =
    &amp;H40 <br>
   &nbsp;&nbsp;&nbsp;PRINTER_STATUS_OFFLINE = &amp;H80 <br>&nbsp;&nbsp;&nbsp;PRINTER_STATUS_IO_ACTIVE =
    &amp;H100 <br>&nbsp;&nbsp;&nbsp;PRINTER_STATUS_BUSY =
    &amp;H200 <br>
   &nbsp;&nbsp;&nbsp;PRINTER_STATUS_PRINTING = &amp;H400 <br>
   &nbsp;&nbsp;&nbsp;PRINTER_STATUS_OUTPUT_BIN_FULL = &amp;H800 <br>&nbsp;&nbsp;&nbsp;PRINTER_STATUS_NOT_AVAILABLE =
    &amp;H1000 <br>
   &nbsp;&nbsp;&nbsp;PRINTER_STATUS_WAITING = &amp;H2000 <br>&nbsp;&nbsp;&nbsp;PRINTER_STATUS_PROCESSING =
    &amp;H4000 <br>&nbsp;&nbsp;&nbsp;PRINTER_STATUS_INITIALIZING =
    &amp;H8000 <br>&nbsp;&nbsp;&nbsp;PRINTER_STATUS_WARMING_UP =
    &amp;H10000 <br>&nbsp;&nbsp;&nbsp;PRINTER_STATUS_TONER_LOW =
    &amp;H20000 <br>&nbsp;&nbsp;&nbsp;PRINTER_STATUS_NO_TONER =
    &amp;H40000 <br>&nbsp;&nbsp;&nbsp;PRINTER_STATUS_PAGE_PUNT =
    &amp;H80000 <br>&nbsp;&nbsp;&nbsp;PRINTER_STATUS_USER_INTERVENTION =
    &amp;H100000 <br>&nbsp;&nbsp;&nbsp;PRINTER_STATUS_OUT_OF_MEMORY =
    &amp;H200000 <br>&nbsp;&nbsp;&nbsp;PRINTER_STATUS_DOOR_OPEN =
    &amp;H400000 <br>&nbsp;&nbsp;&nbsp;PRINTER_STATUS_SERVER_UNKNOWN =
    &amp;H800000 <br>&nbsp;&nbsp;&nbsp;PRINTER_STATUS_POWER_SAVE =
    &amp;H1000000 <br>
   <font class="keyword">End Enum
   </font>
   </p>
  <p align=left><font face="Arial" style="BACKGROUND-COLOR: yellow">&nbsp;3. The data structures</font></p>
  <p align=justify>&nbsp;
  <font face=Arial >
   As each printer driver is responsible for returning
   this data there has to be a standard to which this returned data conforms
   in order for one application to be able to query a number of different
   types of printers. As it happens, there are nine different standard data
   types that can be returned by the <EM>GetPrinter</EM>
   API call in Windows 2000 (only the first two are universal to all current versions of Windows). <br>
  Of these, the second is the most interesting - named PRINTER_INFO_2
  </font>
  <!-- Data structure -->
  <p align=left class="sourcecode">
  <font class="keyword">
  Private Type</font> PRINTER_INFO_2 <br>
  &nbsp;&nbsp;&nbsp;pServerName <font class="keyword">As String</font> <br>
  &nbsp;&nbsp;&nbsp;pPrinterName <font class="keyword">As String</font> <br>
  &nbsp;&nbsp;&nbsp;pShareName <font class="keyword">As String</font> <br>
  &nbsp;&nbsp;&nbsp;pPortName <font class="keyword">As String</font> <br>
  &nbsp;&nbsp;&nbsp;pDriverName <font class="keyword">As String</font> <br>
  &nbsp;&nbsp;&nbsp;pComment <font class="keyword">As String</font> <br>
  &nbsp;&nbsp;&nbsp;pLocation <font class="keyword">As String</font> <br>
  &nbsp;&nbsp;&nbsp;pDevMode <font class="keyword">As Long</font> <br>
  &nbsp;&nbsp;&nbsp;pSepFile <font class="keyword">As String</font> <br>
  &nbsp;&nbsp;&nbsp;pPrintProcessor <font class="keyword">As String</font> <br>
  &nbsp;&nbsp;&nbsp;pDatatype <font class="keyword">As String</font> <br>
  &nbsp;&nbsp;&nbsp;pParameters <font class="keyword">As String</font> <br>
  &nbsp;&nbsp;&nbsp;pSecurityDescriptor <font class="keyword">As Long</font> <br>
  &nbsp;&nbsp;&nbsp;Attributes <font class="keyword">As Long</font> <br>
  &nbsp;&nbsp;&nbsp;Priority <font class="keyword">As Long</font> <br>
  &nbsp;&nbsp;&nbsp;DefaultPriority <font class="keyword">As Long</font> <br>
  &nbsp;&nbsp;&nbsp;StartTime <font class="keyword">As Long</font> <br>
  &nbsp;&nbsp;&nbsp;UntilTime <font class="keyword">As Long</font> <br>
  &nbsp;&nbsp;&nbsp;Status <font class="keyword">As Long</font> <br>
  &nbsp;&nbsp;&nbsp;JobsCount <font class="keyword">As Long</font> <br>
  &nbsp;&nbsp;&nbsp;AveragePPM <font class="keyword">As Long</font> <br>
  <font class="keyword">End Type
  </font>
  </p>
  <p align=justify>&nbsp;
  <font face=Arial >
  However, it is not as simple as just passing this structure to the <i>GetPrinter</i> API call as a printer can
  return more information than is in that structure and if you do not allocate sufficent buffer space for it to
  do so your application will crash. <br>
  Fortunately the API call caters for this - if you pass zero in the <i>pbSize</i> parameter then the API call will
  tell you how big a buffer you will require in the <i>pbSizeNeeded</i>. <br>
  This means that filling the information from the printer driver becomes a two step process:
  </font>
  </p>
  <!-- Using the GetPrinter API call -->
  <p align=left class="sourcecode">  <font class="keyword">
  &nbsp;&nbsp;Dim</font> lret <font class="keyword">As Long</font> <br>
  &nbsp;&nbsp;<font class="keyword">Dim</font> SizeNeeded <font class="keyword">As Long</font><br>
  <br>
  &nbsp;&nbsp;<font class="keyword">Dim</font> buffer() <font class="keyword">As Long</font><br>
  <br>
  &nbsp;&nbsp;<font class="keyword">ReDim Preserve</font> buffer(0 To 1) <font class="keyword">As Long</font> <br>
  &nbsp;&nbsp;lret = GetPrinterApi(mhPrinter, Index, buffer(0), UBound(buffer), SizeNeeded) <br>
  &nbsp;&nbsp;<font class="keyword">ReDim Preserve</font> buffer(0 To (SizeNeeded / 4) + 3) <font class="keyword">As Long</font> <br>
  &nbsp;&nbsp;lret = GetPrinterApi(mhPrinter, Index, buffer(0), UBound(buffer) * 4, SizeNeeded) <br>
  </p>
  <!-- Retrieving the string part -->
  <p align=justify>&nbsp;
  <font face=Arial >
  However the buffer is just an array of <i>Long</i> data types. Some of the data within the PRINTER_INFO_2
  data structure is String data. This must be collected from the addresses which are stored in the appropriate
  buffer position.
  </font>
  </p><p align=justify>&nbsp;
  <font face=Arial >
  To get a string from a pointer the <i>CopyMemory</i> API call is used and there is also an API call,
  <i>IsBadStringPtr</i>, which can be used to verify that the address pointed to does actually contain a valid
  string.
  </font>
  </p>
  <!-- Declarations -->
  <p align=left class="sourcecode">
  <font class="comment">
  '\\ Memory manipulation routines </font><br>
  <font class="keyword">Private Declare Sub</font> CopyMemory <font class="keyword">Lib</font> "kernel32" <font class="keyword">Alias</font> "RtlMoveMemory" (Destination <font class="keyword">As Any,</font> Source <font class="keyword">As Any, ByVal</font> Length <font class="keyword">As Long)</font> <br>
  <font class="comment">'\\ Pointer validation in StringFromPointer </font><br>
  <font class="keyword">Private Declare Function</font> IsBadStringPtrByLong <font class="keyword">Lib</font> "kernel32" <font class="keyword">Alias</font> "IsBadStringPtrA" <font class="keyword">(ByVal</font> lpsz <font class="keyword">As Long, ByVal</font> ucchMax <font class="keyword">As Long) As Long</font> <br>
  </p>
  <p align=justify >
  <font face=Arial >
  Retrieving the string from a pointer is a common thing to have to do so it is worth having this utility function
  in your arsenal.
  </font>
  </p>
  <p align=left class="sourcecode">
  <font class="keyword">
  Public Function</font> StringFromPointer(lpString <font class="keyword">As Long</font>, lMaxLength <font class="keyword">As Long) As String</font><br>
  <br>
  &nbsp;&nbsp;<font class="keyword">Dim</font> sRet <font class="keyword">As String</font><br>
  &nbsp;&nbsp;<font class="keyword">Dim</font> lret <font class="keyword">As Long</font><br>
  <br>
  &nbsp;&nbsp;<font class="keyword">If</font> lpString = 0 <font class="keyword">Then</font><br>
  &nbsp;&nbsp;&nbsp;&nbsp;StringFromPointer = ""<br>
  &nbsp;&nbsp;&nbsp;&nbsp;<font class="keyword">Exit Function</font><br>
  &nbsp;&nbsp;<font class="keyword">End If</font><br>
  <br>
  &nbsp;&nbsp;<font class="keyword">If</font> IsBadStringPtrByLong(lpString, lMaxLength) <font class="keyword">Then</font><br>
  &nbsp;&nbsp;&nbsp;&nbsp;<font class="comment">'\\ An error has occured - do not attempt to use this pointer</font><br>
  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;StringFromPointer = ""<br>
  &nbsp;&nbsp;&nbsp;&nbsp;<font class="keyword">Exit Function</font><br>
  &nbsp;&nbsp;<font class="keyword">End If</font><br>
  <br>
  &nbsp;&nbsp;<font class="comment">'\\ Pre-initialise the return string...</font><br>
  &nbsp;&nbsp;sRet = <font class="keyword">Space$</font>(lMaxLength)<br>
  &nbsp;&nbsp;CopyMemory <font class="keyword">ByVal</font> sRet, <font class="keyword">ByVal</font> lpString, <font class="keyword">ByVal Len(</font>sRet)<br>
  &nbsp;&nbsp;<font class="keyword">If Err.LastDllError</font> = 0 <font class="keyword">Then</font><br>
  &nbsp;&nbsp;&nbsp;&nbsp;<font class="keyword">If InStr(</font>sRet, Chr$(0)) &gt; 0 <font class="keyword">Then</font> <br>
  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;sRet = <font class="keyword">Left$(</font>sRet, InStr(sRet, Chr$(0)) - 1)<br>
  &nbsp;&nbsp;&nbsp;&nbsp;<font class="keyword">End If</font><br>
  &nbsp;&nbsp;<font class="keyword">End If</font><br>
  <br>
  &nbsp;&nbsp;StringFromPointer = sRet<br>
  <br>
  <font class="keyword">End Function</font><br>
  </p>
  <p align=justify>&nbsp;So to use this to populate your PRINTER_INFO_2 variable:
  </p>
  <p class="sourcecode">
  <font class="keyword">With</font> mPRINTER_INFO_2 <font class="comment">'\\ This variable is of type PRINTER_INFO_2</font><br>
  &nbsp;&nbsp;&nbsp;.pServerName = StringFromPointer(buffer(0), 1024)<br>
  &nbsp;&nbsp;&nbsp;.pPrinterName = StringFromPointer(buffer(1), 1024)<br>
  &nbsp;&nbsp;&nbsp;.pShareName = StringFromPointer(buffer(2), 1024)<br>
  &nbsp;&nbsp;&nbsp;.pPortName = StringFromPointer(buffer(3), 1024)<br>
  &nbsp;&nbsp;&nbsp;.pDriverName = StringFromPointer(buffer(4), 1024)<br>
  &nbsp;&nbsp;&nbsp;.pComment = StringFromPointer(buffer(5), 1024)<br>
  &nbsp;&nbsp;&nbsp;.pLocation = StringFromPointer(buffer(6), 1024)<br>
  &nbsp;&nbsp;&nbsp;.pDevMode = buffer(7)<br>
  &nbsp;&nbsp;&nbsp;.pSepFile = StringFromPointer(buffer(8), 1024)<br>
  &nbsp;&nbsp;&nbsp;.pPrintProcessor = StringFromPointer(buffer(9), 1024)<br>
  &nbsp;&nbsp;&nbsp;.pDatatype = StringFromPointer(buffer(10), 1024)<br>
  &nbsp;&nbsp;&nbsp;.pParameters = StringFromPointer(buffer(11), 1024)<br>
  &nbsp;&nbsp;&nbsp;.pSecurityDescriptor = buffer(12)<br>
  &nbsp;&nbsp;&nbsp;.Attributes = buffer(13)<br>
  &nbsp;&nbsp;&nbsp;.Priority = buffer(14)<br>
  &nbsp;&nbsp;&nbsp;.DefaultPriority = buffer(15)<br>
  &nbsp;&nbsp;&nbsp;.StartTime = buffer(16)<br>
  &nbsp;&nbsp;&nbsp;.UntilTime = buffer(17)<br>
  &nbsp;&nbsp;&nbsp;.Status = buffer(18)<br>
  &nbsp;&nbsp;&nbsp;.JobsCount = buffer(19)<br>
  &nbsp;&nbsp;&nbsp;.AveragePPM = buffer(20)<br>
  <font class="keyword">End With</font>
  </p>
    <p align=center><IMG alt="Source code for this article to download" src="../Images/source.gif" align=middle ></p>
   <p align=justify>
   <font face=Arial>The complete source code for these examples is available for download <a href="http://groups.yahoo.com/group/MerrionComputing/files/PrintWatchClient.zip
">here</a><br>
   You will be asked to register with <b>Yahoo!Groups</b> in order to access it.
   </font>
   </p>

