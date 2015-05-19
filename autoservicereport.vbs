'' Info.vbs
'' VBScript for auditing servers
'' Author Kristian B
'' Version 1.0 - April 2010
'' --------------------------------------------------------------' 
'' System Info: select Caption, CSDVersion from Win32_OperatingSystem
'' Processor Info: select Name, MaxClockSpeed from Win32_Processor
'' Ram Info: select select TotalPhysicalMemory from Win32_LogicalMemoryConfiguration
'' Pagefile Info: select FileSize, EightDotThreeFileName, MaximumSize, InitialSize from Win32_PageFile (1.5 gange ram er anbefalet pagefile)
'' Partition Info: select DeviceID, FreeSpace, Size from Win32_LogicalDisk Where DriveType = 3
'' Pagefile Calculated: select CurrentUsage, PeakUsage from Win32_PageFileUsage
'' Processor Calculated: select LoadPercentage from Win32_Processor
'' Ram Calculated: select AvailableBytes from Win32_PerfFormattedData_PerfOS_Memory

Option Explicit

Dim i, j, q, f, arrTemp, intAmount, intServers, intOLEid, intAscii, intTemp
Dim strReportDocument, strCurrentDirectory
Dim arrComputer(), arrUser(), arrPassword(), arrPARTInfo(), arrPAGEInfo()
Dim arrOSInfoName(), arrOSInfoType(), arrOSInfoPack(), intarrOSLength
Dim arrCPUInfoType(), arrCPUInfoMaxspeed(), intarrCPULength
Dim arrPAGEInfoSize(), arrPAGEInfoFilename(), arrPAGEInfoMaxsize(), arrPAGEInfoInitialsize(), intarrPAGELength
Dim arrPARTInfoFilesystem(), arrPARTInfoVolumename(), arrPARTInfoDevice(), arrPARTInfoFreespace(), arrPARTInfoSize(), intarrPARTLength
Dim arrRAMInfo(), intRAMInfoMB, intarrRAMLength
Dim arrCPUCalc(), arrRAMCalc(), intCPUMax, intCPUAvg, intRAMMax, intRAMAvg, intRAMMaxc, intRAMAvgc
Dim objWMIService, objWord, objIShape, objOLE, objLocator, objItem, objItems, xmlDoc, objNodeName, objNodeUser, objNodePassword, objTemp
Dim wbemImpersonationLevelImpersonate, wbemAuthenticationLevelPktPrivacy

get_options()
strCurrentDirectory = left(WScript.ScriptFullName,(Len(WScript.ScriptFullName))-(len(WScript.ScriptName)))
Set objWord=CreateObject("Word.Application")
objWord.Application.Documents.Open(strCurrentDirectory & strReportDocument)
objWord.Visible=True
Set objIShape = objWord.ActiveDocument.InlineShapes


function document_input(intOLENo, strRange, strData, blnRW)
    '' Activates the the inline shape by number(intOLENo) and defines it as the OLE object
    objIShape(intOLENo).OLEFormat.Activate
    Set objOLE = objIShape(intOLENo).OLEFormat.Object
    
    '' Detects the ClassType of the inline shape and uses a class specific counter to count which datafields have data
    Dim strClass, i, p, intSheetno
    intSheetno = 1
    strClass = objIShape(intOLENo).OLEFormat.ClassType
    i = 0
    if Left(strClass, 8) = "MSGraph." then
        if (blnRW) then
            objOLE.Application.DataSheet.Range(strRange) = strData
        elseif (strData = "Scan") then
            for each p in objOLE.Application.DataSheet.Range(strRange)
                if RTrim(p) <> "" then
                    i = i+1
                end if
            next
        elseif (strData = "Read") then
            i = objOLE.Application.DataSheet.Range(strRange)
        end if
    elseif Left(strClass, 6) = "Excel." then
        if (blnRW) then
            objOLE.Worksheets(intSheetno).Range(strRange) = strData
        elseif (strData = "Scan") then
            for each p In objOLE.Worksheets(intSheetno).Range(strRange)
                if RTrim(p) <> "" then
                    i = i+1
                end if
            next
        elseif (strData = "Read") then
            i = objOLE.Worksheets(intSheetno).Range(strRange)
        end if
    end if
    document_input = i
end function


function array_sort(arrSort)
    Dim i, j, temp
    for i = UBound(arrSort) - 1 To 0 Step -1
        for j= 0 to i
            if arrSort(j)>arrSort(j+1) then
                temp=arrSort(j+1)
                arrSort(j+1)=arrSort(j)
                arrSort(j)=temp
            end if
        next
    next
    array_sort = arrSort
end function


function get_options()
    '' Generates three arrays with name, username and password for the servers, from an xml document and fetches any other options (report)
    Set xmlDoc = CreateObject("Msxml2.DOMDocument") 
    xmlDoc.load("config.xml") 
    Set objNodeName = xmlDoc.getElementsByTagName("name")
    Set objNodeUser = xmlDoc.getElementsByTagName("user")
    Set objNodePassword = xmlDoc.getElementsByTagName("password")

    strReportDocument = xmlDoc.getElementsByTagName("report").item(0).Text
    intServers = objNodeName.length
    intAmount =  (intServers * 3)

    Redim arrComputer(intAmount), arrUser(intAmount), arrPassword(intAmount)

    i = 0
    while intServers > i
        arrComputer(i) = objNodeName.item(i).Text
        arrUser(i) = objNodeUser.item(i).Text
        arrPassword(i) = objNodePassword.item(i).Text
        i = i+1
    wend
end function


'' While loop for running through all of the servers counte from config.xml
j = 1
while intServers >= j
    wbemImpersonationLevelImpersonate = 3
    wbemAuthenticationLevelPktPrivacy = 6

    Set objLocator = CreateObject("WbemScripting.SWbemLocator")
    q = j-1
    Set objWMIService = objLocator.ConnectServer (arrComputer(q), "root\cimv2", arrUser(q), arrPassword(q))
    objWMIService.Security_.ImpersonationLevel = wbemImpersonationLevelImpersonate
    objWMIService.Security_.AuthenticationLevel = wbemAuthenticationLevelPktPrivacy

    '' System Info: select Caption, CSDVersion from Win32_OperatingSystem
    ReDim arrOSInfoName(intAmount)
    ReDim arrOSInfoType(intAmount)
    ReDim arrOSInfoPack(intAmount)
    intarrOSLength = 0
    Set objItems = objWMIService.ExecQuery("select Caption, CSName, CSDVersion from Win32_OperatingSystem")
    i = 0
    For Each objItem in objItems
        arrOSInfoName(i) = objItem.CSName
        arrOSInfoType(i) = objItem.Caption
        arrOSInfoPack(i) = objItem.CSDVersion
        i = i+1
        intarrOSLength = i
    Next
    intOLEid = 1
    Call document_input(intOLEid, "A" & j+1, arrOSInfoName(0), true)
    Call document_input(intOLEid, "B" & j+1, arrOSInfoType(0), true)
    Call document_input(intOLEid, "C" & j+1, arrOSInfoPack(0), true)

    '' Processor Info: select Name, MaxClockSpeed from Win32_Processor
    ReDim arrCPUInfoType(intAmount)
    ReDim arrCPUInfoMaxspeed(intAmount)
    intarrCPULength = 0
    Set objItems = objWMIService.ExecQuery("select Name, MaxClockSpeed from Win32_Processor")
    i = 0
    For Each objItem in objItems
        arrCPUInfoType(i) = objItem.Name
        arrCPUInfoMaxspeed(i) = objItem.MaxClockSpeed
        i = i+1
        intarrCPULength = i
    Next
    intOLEid = 2
    Call document_input(intOLEid, "A" & j+1, arrOSInfoName(0), true)
    Call document_input(intOLEid, "B" & j+1, arrCPUInfoType(0), true)
    Call document_input(intOLEid, "C" & j+1, (arrCPUInfoMaxspeed(0) / 1000) & " GHz", true)

    '' Ram Info: select TotalPhysicalMemory from Win32_LogicalMemoryConfiguration
    ReDim arrRAMInfo(intAmount)
    intarrRAMLength = 0
    Set objItems = objWMIService.ExecQuery("select TotalPhysicalMemory from Win32_LogicalMemoryConfiguration")
    i = 0
    For Each objItem in objItems
        arrRAMInfo(i) = objItem.TotalPhysicalMemory
        i = i+1
        intarrRAMLength = i
    Next
    intOLEid = 4
    intRAMInfoMB = Round((arrRAMInfo(0) / 1024))
    Call document_input(intOLEid, "A" & j+1, arrOSInfoName(0), true)
    Call document_input(intOLEid, "B" & j+1, intRAMInfoMB, true)
    intAscii = j+64
    Call document_input(intOLEid+1, Chr(intAscii) & 1, intRAMInfoMB, true)

    '' Pagefile Info: select FileSize, EightDotThreeFileName, MaximumSize, InitialSize from Win32_PageFile (1.5 gange ram er anbefalet pagefile)
    ReDim arrPAGEInfoSize(intAmount)
    ReDim arrPAGEInfoFilename(intAmount)
    ReDim arrPAGEInfoMaxsize(intAmount)
    ReDim arrPAGEInfoInitialsize(intAmount)
    intarrPAGELength = 0
    Set objTemp = objWMIService.ExecQuery("select Name, AllocatedBaseSize, CurrentUsage from Win32_PageFileUsage")
    i = 0
    For Each objItem in objTemp
        arrPAGEInfoMaxsize(i) = (intRAMInfoMB * 1.5)
        arrPAGEInfoInitialsize(i) = objItem.AllocatedBaseSize
        arrPAGEInfoFilename(i) = objItem.Name
        arrPAGEInfoSize(i) = objItem.CurrentUsage
        i = i+1
    Next
    intOLEid = 6
    Call document_input(intOLEid, "A" & j+1, arrOSInfoName(0), true)
    Call document_input(intOLEid, "B" & j+1, arrPAGEInfoMaxsize(0) & " MB", true)
    Call document_input(intOLEid, "C" & j+1, arrPAGEInfoInitialsize(0) & " MB", true)
    Call document_input(intOLEid, "D" & j+1, arrPAGEInfoSize(0) & " MB", true)
    Call document_input(intOLEid, "E" & j+1, arrPAGEInfoFilename(0), true)

    '' Partition Info: select DeviceID, FreeSpace, Size from Win32_LogicalDisk Where DriveType = 3
    ReDim arrPARTInfoDevice(intAmount)
    ReDim arrPARTInfoVolumename(intAmount)
    ReDim arrPARTInfoFilesystem(intAmount)
    ReDim arrPARTInfoSize(intAmount)
    ReDim arrPARTInfoFreespace(intAmount)
    intarrPARTLength = 0
    Set objItems = objWMIService.ExecQuery("select VolumeName, FileSystem, DeviceID, FreeSpace, Size from Win32_LogicalDisk Where DriveType = 3")
    i = 0
    For Each objItem in objItems
        arrPARTInfoDevice(i) = objItem.DeviceID
        arrPARTInfoVolumename(i) = objItem.VolumeName
        arrPARTInfoFilesystem(i) = objItem.FileSystem
        arrPARTInfoSize(i) = objItem.Size
        arrPARTInfoFreespace(i) = objItem.FreeSpace
        i = i+1
        intarrPARTLength = i
    Next
    intOLEid = 6 + (2 * j)
    intAscii = document_input(intOLEid+1, "A1:Z1", "Scan", false) + 65
    i = 0
    while intarrPARTLength > i
        Call document_input(intOLEid, "A" & i+2, arrPARTInfoDevice(i) & " (" & arrPARTInfoVolumename(i) & ")", true)
        Call document_input(intOLEid, "B" & i+2, arrPARTInfoFilesystem(i), true)
        Call document_input(intOLEid, "C" & i+2, Round(((arrPARTInfoSize(i)/1024)/1024)/1024), true)
        Call document_input(intOLEid, "D" & i+2, Round(((arrPARTInfoFreespace(i)/1024)/1024)/1024), true)
        intTemp = document_input(intOLEid, "E" & i+2, "Read", false)
        Call document_input(intOLEid+1, Chr(intAscii) & i+1, intTemp, true)
        i = i+1
    wend
    
    '' Processor Calculated: select LoadPercentage from Win32_Processor
    '' Ram Calculated: select AvailableBytes from Win32_PerfFormattedData_PerfOS_Memory
    f = 8
    i = 0
    ReDim arrCPUCalc(f)
    ReDim arrRAMCalc(f)
    while i <= f
        Set objItems = objWMIService.ExecQuery("select LoadPercentage from Win32_Processor")
        For Each objItem in objItems
            arrCPUCalc(i) = objItem.LoadPercentage
        Next
        Set objItems = objWMIService.ExecQuery("select AvailableBytes from Win32_PerfFormattedData_PerfOS_Memory")
        For Each objItem in objItems
            arrRAMCalc(i) = objItem.AvailableBytes
        Next
        if i/3 = 1 OR i/3 = 2 then
            wscript.sleep 5000
        else
            wscript.sleep 1000
        end if
        i = i+1
    wend
    
    intTemp = 0
    arrTemp = array_sort(arrCPUCalc)
    For Each objItem in arrTemp
        intTemp = intTemp + objItem
    Next
    intCPUAvg = Round((intTemp / (f+1)))
    intCPUMax = arrTemp(f)
    
    intTemp = 0
    arrTemp = array_sort(arrRAMCalc)
    For Each objItem in arrTemp
        intTemp = intTemp + objItem
    Next
    intRAMAvg = (intTemp / (f+1))
    intRAMMax = arrTemp(UBound(arrTemp) - 1)
    
    intOLEid = 2
    intAscii = j + 64
    Call document_input(intOLEid, "D" & j+1, intCPUAvg & "%", true)
    Call document_input(intOLEid, "E" & j+1, intCPUMax & "%", true)
    Call document_input(intOLEid+1, Chr(intAscii) & 0, arrOSInfoName(0), true)
    Call document_input(intOLEid+1, Chr(intAscii) & 1, intCPUAvg & "%", true)
    Call document_input(intOLEid+1, Chr(intAscii) & 2, intCPUMax & "%", true)
    
    intRAMAvgc = Round((intRAMAvg/1024)/1024)
    intRAMMaxc = Round((intRAMMax/1024)/1024)
    Call document_input(intOLEid+2, "C" & j+1, intRAMAvgc, true)
    Call document_input(intOLEid+2, "F" & j+1, intRAMMaxc, true)
    intTemp = document_input(intOLEid+2, "G" & j+1, "Read", false)
    Call document_input(intOLEid+3, Chr(intAscii) & 0, arrOSInfoName(0), true)
    Call document_input(intOLEid+3, Chr(intAscii) & 2, intRAMAvgc, true)
    Call document_input(intOLEid+3, Chr(intAscii) & 3, intTemp, true)
    Call document_input(intOLEid+5, Chr(intAscii) & 0, arrOSInfoName(0), true)
    Call document_input(intOLEid+5, Chr(intAscii) & 1, arrPAGEInfoMaxsize(0), true)
    Call document_input(intOLEid+5, Chr(intAscii) & 2, arrPAGEInfoSize(0), true)
    
    j = j+1
wend


''objWord.Application.Documents.Save
''objWord.Application.Documents.Close
''objWord.Application.Quit
wscript.echo "The program ended successfully!"
WSCript.Quit(0)