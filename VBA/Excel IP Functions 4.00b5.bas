Attribute VB_Name = "IPFunctions"
' IP Functions v4.00
' © 2013-2025 Nicolas Grison

'
' Main Module for IP Functions
'


' Version History
'
' 4.00b4 Added ipClassless function
'        Most functions are now Range compatible
' 4.00 Rewritten from the ground up to move to a Class IP object and optimisation
'      Functions speeded up
'      Excel errors now returned when needed instead of a text string
'      OutputFormat moved to a custom type that includes all parameters
'      ipIncluded Array function removed, replaced by an hybrid standard/array ipRoute function that determines
'          the best match, not the first
'      VBA Immediate window displays detailed error messages including the calling worksheet and cell
' 3.06 Added ipRoute function
' 3.05 Corrected bug with subnet masks provided in quad octets form (Mask Length was not correctly filled)
'      Added option not to merge subnets when summarising
'      Added ipSort array function
'      Corrected a bug in IPBinMask2Length that returned 31 for 32 mask length
'      Added ipIncluded array function
' 3.04 Added ipReformat and ipSortand Reformat Macros
'      Added ipSumExtract array function
'      Corrected bugs with ipHosts and ipSubnets functions
' 3.03 if both input parameters of ipSumCheck function are ranges, if the Input is a cell included in the Summary range
'      it will be ignored and compared only to the other cells in the range. This is useful to quickly check whether a range of networks
'      overlap with each others.
' 3.02 ipSumCheck function now accepts string or range of cells as summary input
' 3.01 Added ipSumCheck function
' 3.00 Added ipSumSubX and ipSumSubY functions. Maximum supported decimal offset value
'      increased to 79,228,162,514,264,337,593,543,950,335.
'      Added ipHostX and ipHostY functions
'      Added ipHostCount and ipSumSubCount (Those will return approximations for huge numbers in IPv6)
'      Updated and renamed First and Last Summary Subnet functions, breaks compatibility with previous version.
'      Functions are now called ipSumFirstSub and ipSumLastSub
'      Corrected reserved anycast address bug for IPv6 addresses
'      Renamed all host functions ipHostXXX, breaks compatibility with previous version.
'      Internal code optimisation and cleanup
'      Changed offset behaviour: for host functions, offset returns an error if the result is outside the subnet
'      For subnet functions the offset can only be an IP address
'      Instruction spreadsheet updated
'      Cue card removed due to different arguments
'      CDec function replaced with CLng for Office Mac compatibility
' 2.38 Implied prefix length changed from /32 (IPv4) or /128 (IPv6) to /0
' 2.37 Corrected a bug with quad-dotted addresses, offset and short IPv6 notation
' 2.36 ipSubMask now returns a Byte value instead of a string to allow for number comparisons
' 2.35 Added ipHosts function, changed OutputFormat format, renamed ipSubnetting as ipSubnets
' 2.34 Offset bug correction (checks output is still part of the subnet unless the offset is IP, always check against the summary)
' 2.33 Removed global "IP" function
' 2.32 Added Array subnetting function
' 2.24-31 Mostly rewritten for internal optimisation and clean up
' 2.23 Added subnetting (incl. function 13 & 14), binary output - edit shortcuts
' 2.22 Added trim function to remove extra spaces
' 2.21 Modified output formats. Padded option now a boolean. Corrected shortcut errors.
' 2.20 Added name shortcuts for FunctionSelector and macro for functionselector shortcut
' 2.19 Added IPFunction version number lookup (function 99)
' 2.18 Bug corrections
' 2.17 Added 'include subnet mask' option
' 2.16 Changed PaddedOption to OutputFormat to include 'significant Bytes' feature


' The following Excel errors can be returned by the functions:
'
'#VALUE! Error in the input data (i.e. unexpected characters in the ip address, incorrect format string,...)
'#N/A!   Asked a function something that does not exist (i.e. ipBroadcastAddress on an IPv6 address)
'#NUM!   The result of the operation is out of bound (i.e. ipFirstHostAddress with an offset of -1)
'
' For errors to return correctly, VBA Editor Error Trapping must be set to "Break on Unhandled Errors"



Option Explicit





'##############################################
'#
'# Declarations
'#
'##############################################

Private Const gIPfversion As Single = 4
Private Const EnableDebug As Boolean = True

Private Type tIPFormat
    Format As Long
    Padded As Boolean
    QuadDotted As Boolean
End Type





'##############################################
'#
'# Methods
'#
'##############################################

''''''''''''''''''''''''''''''
' Error Handler
''''''''''''''''''''''''''''''

Private Function errorHandling() As Variant
    
    Dim caller As String
    
    Select Case Err.Number
        Case vbObjectError + 1010
            ' Input error (wrong address,...)
            errorHandling = CVErr(xlErrValue)
        
        Case vbObjectError + 1020
            ' Non existent value (i.e. ipv6 broadcast address)
            errorHandling = CVErr(xlErrNA)
        
        Case vbObjectError + 1030
            ' Out Of Bound output value
            errorHandling = CVErr(xlErrNum)
    End Select
    
    Select Case TypeName(Application.caller)
        Case "Range"
            caller = ChrW(39) & Application.caller.Worksheet.Name & ChrW(39) & "!" & Application.caller.Address
        Case "String"
            caller = Application.caller
        Case "Error"
            caller = "Error"
        Case Else
            caller = "unknown"
    End Select
    
    Debug.Print "Caller: " & caller & "   Error: " & Err.Number - vbObjectError & " (" & Err.Description & ") " & Err.Source
    
End Function





'##############################################
'#
'# Public IP Functions
'#
'##############################################



''''''''''''''''''''''''''''''
' Host Functions
''''''''''''''''''''''''''''''

Public Function ipHostFirst(inputVal As Variant, Optional DisplayFormat As Long = 1, _
                          Optional Offset As String = vbNullString) As Variant

    ' Returns the first host address of the subnet of an IP/Mask string

    Dim BinaryipAddress As clsIP
    Dim ProcessedIP As clsIP
    Dim i As Long, j As Long
    Dim outputArray() As Variant

    If TypeName(inputVal) = "Range" Then
        
        On Error GoTo ErrorHandlerRange
        
        ReDim outputArray(1 To inputVal.Rows.Count, 1 To inputVal.Columns.Count)
        
        For i = 1 To inputVal.Rows.Count
            For j = 1 To inputVal.Columns.Count
                
                'Convert the input into a Binary pair
                Set BinaryipAddress = ParseIP(CStr(inputVal.Cells(i, j).Value))
            
                'Binary Function
                Set ProcessedIP = BinaryipAddress.FirstHost.Offset(parseOffset(Offset))
            
                'Post Process the output
                outputArray(i, j) = FormatOutput(ProcessedIP, DisplayFormat)

nextCell:
            Next j
        Next i
    
        ipHostFirst = outputArray
        
    ElseIf TypeName(inputVal) = "String" Then
        
        On Error GoTo ErrorHandlerString
        
        'Convert the input into a Binary pair
        Set BinaryipAddress = ParseIP(CStr(inputVal))
    
        'Binary Function
        Set ProcessedIP = BinaryipAddress.FirstHost.Offset(parseOffset(Offset))
    
        'Post Process the output
        ipHostFirst = FormatOutput(ProcessedIP, DisplayFormat)

    Else
        ' If the input is not a range or a string, return an error message
        ipHostFirst = "Invalid input: must be a range or a string"
    End If
    
    Exit Function

ErrorHandlerRange:
                outputArray(i, j) = "#VALUE!"
                Resume nextCell 'Clear the error and continues the executionat the label

ErrorHandlerString:
        ipHostFirst = CVErr(xlErrValue)

End Function



Public Function ipHostPrev(inputVal As Variant, Optional DisplayFormat As Long = 1, _
                          Optional Offset As String = vbNullString) As Variant

    ' Returns the previous host address of an IP/Mask string

    Dim BinaryipAddress As clsIP
    Dim ProcessedIP As clsIP
    Dim i As Long, j As Long
    Dim outputArray() As Variant

    If TypeName(inputVal) = "Range" Then
        
        On Error GoTo ErrorHandlerRange
        
        ReDim outputArray(1 To inputVal.Rows.Count, 1 To inputVal.Columns.Count)
        
        For i = 1 To inputVal.Rows.Count
            For j = 1 To inputVal.Columns.Count
                
                'Convert the input into a Binary pair
                Set BinaryipAddress = ParseIP(CStr(inputVal.Cells(i, j).Value))
            
                'Binary Function
                Set ProcessedIP = BinaryipAddress.PreviousHost.Offset(parseOffset(Offset))
            
                'Post Process the output
                outputArray(i, j) = FormatOutput(ProcessedIP, DisplayFormat)

nextCell:
            Next j
        Next i
    
        ipHostPrev = outputArray
        
    ElseIf TypeName(inputVal) = "String" Then
        
        On Error GoTo ErrorHandlerString
        
        'Convert the input into a Binary pair
        Set BinaryipAddress = ParseIP(CStr(inputVal))
    
        'Binary Function
        Set ProcessedIP = BinaryipAddress.PreviousHost.Offset(parseOffset(Offset))
    
        'Post Process the output
        ipHostPrev = FormatOutput(ProcessedIP, DisplayFormat)

    Else
        ' If the input is not a range or a string, return an error message
        ipHostPrev = "Invalid input: must be a range or a string"
    End If
    
    Exit Function

ErrorHandlerRange:
                outputArray(i, j) = "#VALUE!"
                Resume nextCell 'Clear the error and continues the executionat the label

ErrorHandlerString:
        ipHostPrev = CVErr(xlErrValue)

End Function



Public Function ipAddress(inputVal As Variant, Optional DisplayFormat As Long = 1, _
                          Optional Offset As String = vbNullString) As Variant

    ' Returns the IP address of an IP/Mask string

    Dim BinaryipAddress As clsIP
    Dim ProcessedIP As clsIP
    Dim i As Long, j As Long
    Dim outputArray() As Variant

    If TypeName(inputVal) = "Range" Then
        
        On Error GoTo ErrorHandlerRange
        
        ReDim outputArray(1 To inputVal.Rows.Count, 1 To inputVal.Columns.Count)
        
        For i = 1 To inputVal.Rows.Count
            For j = 1 To inputVal.Columns.Count
                
                'Convert the input into a Binary pair
                Set BinaryipAddress = ParseIP(CStr(inputVal.Cells(i, j).Value))
            
                'Binary Function
                Set ProcessedIP = BinaryipAddress.Offset(parseOffset(Offset))
            
                'Post Process the output
                outputArray(i, j) = FormatOutput(ProcessedIP, DisplayFormat)

nextCell:
            Next j
        Next i
    
        ipAddress = outputArray
        
    ElseIf TypeName(inputVal) = "String" Then
        
        On Error GoTo ErrorHandlerString
        
        'Convert the input into a Binary pair
        Set BinaryipAddress = ParseIP(CStr(inputVal))
    
        'Binary Function
        Set ProcessedIP = BinaryipAddress.Offset(parseOffset(Offset))
    
        'Post Process the output
        ipAddress = FormatOutput(ProcessedIP, DisplayFormat)

    Else
        ' If the input is not a range or a string, return an error message
        ipAddress = "Invalid input: must be a range or a string"
    End If
    
    Exit Function

ErrorHandlerRange:
                outputArray(i, j) = "#VALUE!"
                Resume nextCell 'Clear the error and continues the executionat the label

ErrorHandlerString:
        ipAddress = CVErr(xlErrValue)

End Function



Public Function ipHostNext(inputVal As Variant, Optional DisplayFormat As Long = 1, _
                          Optional Offset As String = vbNullString) As Variant

    ' Returns the next host address of an IP/Mask string

    Dim BinaryipAddress As clsIP
    Dim ProcessedIP As clsIP
    Dim i As Long, j As Long
    Dim outputArray() As Variant

    If TypeName(inputVal) = "Range" Then
        
        On Error GoTo ErrorHandlerRange
        
        ReDim outputArray(1 To inputVal.Rows.Count, 1 To inputVal.Columns.Count)
        
        For i = 1 To inputVal.Rows.Count
            For j = 1 To inputVal.Columns.Count
                
                'Convert the input into a Binary pair
                Set BinaryipAddress = ParseIP(CStr(inputVal.Cells(i, j).Value))
            
                'Binary Function
                Set ProcessedIP = BinaryipAddress.NextHost.Offset(parseOffset(Offset))
            
                'Post Process the output
                outputArray(i, j) = FormatOutput(ProcessedIP, DisplayFormat)

nextCell:
            Next j
        Next i
    
        ipHostNext = outputArray
        
    ElseIf TypeName(inputVal) = "String" Then
        
        On Error GoTo ErrorHandlerString
        
        'Convert the input into a Binary pair
        Set BinaryipAddress = ParseIP(CStr(inputVal))
    
        'Binary Function
        Set ProcessedIP = BinaryipAddress.NextHost.Offset(parseOffset(Offset))
    
        'Post Process the output
        ipHostNext = FormatOutput(ProcessedIP, DisplayFormat)

    Else
        ' If the input is not a range or a string, return an error message
        ipHostNext = "Invalid input: must be a range or a string"
    End If
    
    Exit Function

ErrorHandlerRange:
                outputArray(i, j) = "#VALUE!"
                Resume nextCell 'Clear the error and continues the executionat the label

ErrorHandlerString:
        ipHostNext = CVErr(xlErrValue)

End Function



Public Function ipHostLast(inputVal As Variant, Optional DisplayFormat As Long = 1, _
                          Optional Offset As String = vbNullString) As Variant

    ' Returns the last host address of the subnet of an IP/Mask string

    Dim BinaryipAddress As clsIP
    Dim ProcessedIP As clsIP
    Dim i As Long, j As Long
    Dim outputArray() As Variant

    If TypeName(inputVal) = "Range" Then
        
        On Error GoTo ErrorHandlerRange
        
        ReDim outputArray(1 To inputVal.Rows.Count, 1 To inputVal.Columns.Count)
        
        For i = 1 To inputVal.Rows.Count
            For j = 1 To inputVal.Columns.Count
                
                'Convert the input into a Binary pair
                Set BinaryipAddress = ParseIP(CStr(inputVal.Cells(i, j).Value))
            
                'Binary Function
                Set ProcessedIP = BinaryipAddress.LastHost.Offset(parseOffset(Offset))
            
                'Post Process the output
                outputArray(i, j) = FormatOutput(ProcessedIP, DisplayFormat)

nextCell:
            Next j
        Next i
    
        ipHostLast = outputArray
        
    ElseIf TypeName(inputVal) = "String" Then
        
        On Error GoTo ErrorHandlerString
        
        'Convert the input into a Binary pair
        Set BinaryipAddress = ParseIP(CStr(inputVal))
    
        'Binary Function
        Set ProcessedIP = BinaryipAddress.LastHost.Offset(parseOffset(Offset))
    
        'Post Process the output
        ipHostLast = FormatOutput(ProcessedIP, DisplayFormat)

    Else
        ' If the input is not a range or a string, return an error message
        ipHostLast = "Invalid input: must be a range or a string"
    End If
    
    Exit Function

ErrorHandlerRange:
                outputArray(i, j) = "#VALUE!"
                Resume nextCell 'Clear the error and continues the executionat the label

ErrorHandlerString:
        ipHostLast = CVErr(xlErrValue)

End Function



Public Function ipHostX(inputVal As Variant, HostNumber As Variant, Optional DisplayFormat As Long = 1, _
                          Optional Offset As String = vbNullString) As Variant

    ' Returns the Xth IP address in the subnet

    Dim BinaryipAddress As clsIP
    Dim ProcessedIP As clsIP
    Dim BinHostNumber As String
    Dim i As Long, j As Long
    Dim outputArray() As Variant

    If TypeName(inputVal) = "Range" Then
        
        On Error GoTo ErrorHandlerRange
        
        ReDim outputArray(1 To inputVal.Rows.Count, 1 To inputVal.Columns.Count)
        
        For i = 1 To inputVal.Rows.Count
            For j = 1 To inputVal.Columns.Count
                
                'Convert the input into a Binary pair
                Set BinaryipAddress = ParseIP(CStr(inputVal.Cells(i, j).Value))
                
                BinHostNumber = parseOffset(HostNumber)
                
                'Binary Function
                Set ProcessedIP = BinaryipAddress.HostX(BinHostNumber).Offset(parseOffset(Offset))
            
                'Post Process the output
                outputArray(i, j) = FormatOutput(ProcessedIP, DisplayFormat)

nextCell:
            Next j
        Next i
    
        ipHostX = outputArray
        
    ElseIf TypeName(inputVal) = "String" Then
        
        On Error GoTo ErrorHandlerString
        
        'Convert the input into a Binary pair
        Set BinaryipAddress = ParseIP(CStr(inputVal))
        
        BinHostNumber = parseOffset(HostNumber)
    
        'Binary Function
        Set ProcessedIP = BinaryipAddress.HostX(BinHostNumber).Offset(parseOffset(Offset))
    
        'Post Process the output
        ipHostX = FormatOutput(ProcessedIP, DisplayFormat)

    Else
        ' If the input is not a range or a string, return an error message
        ipHostX = "Invalid input: must be a range or a string"
    End If
    
    Exit Function

ErrorHandlerRange:
                outputArray(i, j) = "#VALUE!"
                Resume nextCell 'Clear the error and continues the executionat the label

ErrorHandlerString:
        ipHostX = CVErr(xlErrValue)

End Function



Public Function ipHostY(inputVal As Variant, HostNumber As Variant, Optional DisplayFormat As Long = 1, _
                          Optional Offset As String = vbNullString) As Variant

    ' Returns the Yth IP address in the subnet from the end

    Dim BinaryipAddress As clsIP
    Dim ProcessedIP As clsIP
    Dim BinHostNumber As String
    Dim i As Long, j As Long
    Dim outputArray() As Variant

    If TypeName(inputVal) = "Range" Then
        
        On Error GoTo ErrorHandlerRange
        
        ReDim outputArray(1 To inputVal.Rows.Count, 1 To inputVal.Columns.Count)
        
        For i = 1 To inputVal.Rows.Count
            For j = 1 To inputVal.Columns.Count
                
                'Convert the input into a Binary pair
                Set BinaryipAddress = ParseIP(CStr(inputVal.Cells(i, j).Value))
                
                BinHostNumber = parseOffset(HostNumber)
                
                'Binary Function
                Set ProcessedIP = BinaryipAddress.HostY(BinHostNumber).Offset(parseOffset(Offset))
            
                'Post Process the output
                outputArray(i, j) = FormatOutput(ProcessedIP, DisplayFormat)

nextCell:
            Next j
        Next i
    
        ipHostY = outputArray
        
    ElseIf TypeName(inputVal) = "String" Then
        
        On Error GoTo ErrorHandlerString
        
        'Convert the input into a Binary pair
        Set BinaryipAddress = ParseIP(CStr(inputVal))
        
        BinHostNumber = parseOffset(HostNumber)
    
        'Binary Function
        Set ProcessedIP = BinaryipAddress.HostY(BinHostNumber).Offset(parseOffset(Offset))
    
        'Post Process the output
        ipHostY = FormatOutput(ProcessedIP, DisplayFormat)

    Else
        ' If the input is not a range or a string, return an error message
        ipHostY = "Invalid input: must be a range or a string"
    End If
    
    Exit Function

ErrorHandlerRange:
                outputArray(i, j) = "#VALUE!"
                Resume nextCell 'Clear the error and continues the executionat the label

ErrorHandlerString:
        ipHostY = CVErr(xlErrValue)

End Function


Public Function ipHostCount(inputVal As Variant, Optional DisplayFormat As Long = 1, _
                          Optional Offset As String = vbNullString) As Variant

    ' Returns the last host address of the subnet of an IP/Mask string

    Dim BinaryipAddress As clsIP
    Dim ProcessedIP As clsIP
    Dim i As Long, j As Long
    Dim outputArray() As Variant

    If TypeName(inputVal) = "Range" Then
        
        On Error GoTo ErrorHandlerRange
        
        ReDim outputArray(1 To inputVal.Rows.Count, 1 To inputVal.Columns.Count)
        
        For i = 1 To inputVal.Rows.Count
            For j = 1 To inputVal.Columns.Count
                
                'Convert the input into a Binary pair
                Set BinaryipAddress = ParseIP(CStr(inputVal.Cells(i, j).Value))
            
                'Binary Function
                outputArray(i, j) = cvBinToDec(BinaryipAddress.HostCount())

nextCell:
            Next j
        Next i
    
        ipHostCount = outputArray
        
    ElseIf TypeName(inputVal) = "String" Then
        
        On Error GoTo ErrorHandlerString
        
        'Convert the input into a Binary pair
        Set BinaryipAddress = ParseIP(CStr(inputVal))
    
        'Post Process the output
        ipHostCount = cvBinToDec(BinaryipAddress.HostCount())

    Else
        ' If the input is not a range or a string, return an error message
        ipHostCount = "Invalid input: must be a range or a string"
    End If
    
    Exit Function

ErrorHandlerRange:
                outputArray(i, j) = "#VALUE!"
                Resume nextCell 'Clear the error and continues the executionat the label

ErrorHandlerString:
        ipHostCount = CVErr(xlErrValue)

End Function





''''''''''''''''''''''''''''''
' Subnet Functions
''''''''''''''''''''''''''''''

Public Function ipSubAddress(inputVal As Variant, Optional DisplayFormat As Long = 1, _
                          Optional Offset As String = vbNullString, Optional Summary As String = vbNullString) As Variant

    ' Returns the subnet IP address of an IP/Mask string

    Dim BinaryipAddress As clsIP
    Dim ProcessedIP As clsIP
    Dim i As Long, j As Long
    Dim outputArray() As Variant

    If TypeName(inputVal) = "Range" Then
        
        On Error GoTo ErrorHandlerRange
        
        ReDim outputArray(1 To inputVal.Rows.Count, 1 To inputVal.Columns.Count)
        
        For i = 1 To inputVal.Rows.Count
            For j = 1 To inputVal.Columns.Count
                
                'Convert the input into a Binary pair
                Set BinaryipAddress = ParseIP(CStr(inputVal.Cells(i, j).Value))
            
                'Binary Function
                Set ProcessedIP = BinaryipAddress.SubAddress().Offset(parseOffset(Offset))
            
                'Post Process the output
                outputArray(i, j) = FormatOutput(ProcessedIP, DisplayFormat)

nextCell:
            Next j
        Next i
    
        ipSubAddress = outputArray
        
    ElseIf TypeName(inputVal) = "String" Then
        
        On Error GoTo ErrorHandlerString
        
        'Convert the input into a Binary pair
        Set BinaryipAddress = ParseIP(CStr(inputVal))
    
        'Binary Function
        Set ProcessedIP = BinaryipAddress.SubAddress().Offset(parseOffset(Offset))
    
        'Post Process the output
        ipSubAddress = FormatOutput(ProcessedIP, DisplayFormat)

    Else
        ' If the input is not a range or a string, return an error message
        ipSubAddress = "Invalid input: must be a range or a string"
    End If
    
    Exit Function

ErrorHandlerRange:
                outputArray(i, j) = "#VALUE!"
                Resume nextCell 'Clear the error and continues the executionat the label

ErrorHandlerString:
        ipSubAddress = CVErr(xlErrValue)

End Function


Public Function ipSubMask(inputVal As Variant, Optional DisplayFormat As Long = 2, _
                          Optional Offset As String, Optional Summary As String) As Variant
    
    ' Returns the number of subnets of PrefixLength size within the summary
    
    Dim BinaryipAddress As clsIP
    Dim OutputFormat As tIPFormat
    Dim i As Long, j As Long
    Dim outputArray() As Variant
    
    OutputFormat = parseDisplayFormat(DisplayFormat)
    
    If TypeName(inputVal) = "Range" Then
        
        On Error GoTo ErrorHandlerRange
        
        ReDim outputArray(1 To inputVal.Rows.Count, 1 To inputVal.Columns.Count)
        
        For i = 1 To inputVal.Rows.Count
            For j = 1 To inputVal.Columns.Count
                
                'Convert the input into a Binary pair
                Set BinaryipAddress = ParseIP(CStr(inputVal.Cells(i, j).Value))
                
                Select Case OutputFormat.Format
            
                Case 2
                
                    ' Returns the mask length of an IP/Mask string
            
                    ipSubMask = BinaryipAddress.PrefixLength
            
                    If OutputFormat.Padded = True Then
                        If BinaryipAddress.IPv6 = True Then
                            ipSubMask = Format(ipSubMask, "000")
                        Else
                            ipSubMask = Format(ipSubMask, "00")
                        End If
                    End If
            
                    outputArray(i, j) = ipSubMask
                
                Case 3
                
                    ' Returns the subnet mask of an IP/Mask string
            
                    If BinaryipAddress.IPv6 Then
                        Err.Raise vbObjectError + 1020, "ipMask", "#SubnetMask not supported with IPv6 Addresses!"
                    Else
                        outputArray(i, j) = cvIP4Bin2Dec(BinaryipAddress.Mask, OutputFormat.Padded, False)
                    End If
                    
                
                Case 4
            
                    ' Returns the wildcard mask of an IP/Mask string
                        
                    If BinaryipAddress.IPv6 Then
                        Err.Raise vbObjectError + 1020, "ipMask", "#SubnetMask not supported with IPv6 Addresses!"
                    Else
                        outputArray(i, j) = cvIP4Bin2Dec(BinaryipAddress.WildcardMask, OutputFormat.Padded, False)
                    End If
            
                
                Case 5
                
                    ' Returns the subnet mask in binary
            
                    outputArray(i, j) = BinaryipAddress.Mask
            
            
                Case Else
                
                    Err.Raise vbObjectError + 1010, "ipSubMask", "#Unsupported format!"
                
                End Select

nextCell:
            Next j
        Next i
    
        ipSubMask = outputArray
        
    ElseIf TypeName(inputVal) = "String" Then
        
        On Error GoTo ErrorHandlerString
        
        'Convert the input into a Binary pair
        Set BinaryipAddress = ParseIP(CStr(inputVal))
        
        Select Case OutputFormat.Format
    
        Case 2
        
            ' Returns the mask length of an IP/Mask string
    
            ipSubMask = BinaryipAddress.PrefixLength
    
            If OutputFormat.Padded = True Then
                If BinaryipAddress.IPv6 = True Then
                    ipSubMask = Format(ipSubMask, "000")
                Else
                    ipSubMask = Format(ipSubMask, "00")
                End If
            End If
    
            Exit Function
        
        Case 3
        
            ' Returns the subnet mask of an IP/Mask string
    
            If BinaryipAddress.IPv6 Then
                Err.Raise vbObjectError + 1020, "ipMask", "#SubnetMask not supported with IPv6 Addresses!"
            Else
                ipSubMask = cvIP4Bin2Dec(BinaryipAddress.Mask, OutputFormat.Padded, False)
            End If
            
            Exit Function
        
        Case 4
    
            ' Returns the wildcard mask of an IP/Mask string
                
            If BinaryipAddress.IPv6 Then
                Err.Raise vbObjectError + 1020, "ipMask", "#SubnetMask not supported with IPv6 Addresses!"
            Else
                ipSubMask = cvIP4Bin2Dec(BinaryipAddress.WildcardMask, OutputFormat.Padded, False)
            End If
    
            Exit Function
        
        Case 5
        
            ' Returns the subnet mask in binary
    
            ipSubMask = BinaryipAddress.Mask
    
            Exit Function
    
        Case Else
        
            Err.Raise vbObjectError + 1010, "ipSubMask", "#Unsupported format!"
        
        End Select
    Else
        ' If the input is not a range or a string, return an error message
        ipSubMask = "Invalid input: must be a range or a string"
    End If
    
    Exit Function

ErrorHandlerRange:
                outputArray(i, j) = "#VALUE!"
                Resume nextCell 'Clear the error and continues the executionat the label

ErrorHandlerString:
        ipSubMask = CVErr(xlErrValue)

End Function


Public Function ipSubBroadcast(inputVal As Variant, Optional DisplayFormat As Long = 1, _
                          Optional Offset As String = vbNullString) As Variant

    ' Returns the broadcast address of the subnet of an IP/Mask string

    Dim BinaryipAddress As clsIP
    Dim ProcessedIP As clsIP
    Dim i As Long, j As Long
    Dim outputArray() As Variant

    If TypeName(inputVal) = "Range" Then
        
        On Error GoTo ErrorHandlerRange
        
        ReDim outputArray(1 To inputVal.Rows.Count, 1 To inputVal.Columns.Count)
        
        For i = 1 To inputVal.Rows.Count
            For j = 1 To inputVal.Columns.Count
                
                'Convert the input into a Binary pair
                Set BinaryipAddress = ParseIP(CStr(inputVal.Cells(i, j).Value))
            
                'Binary Function
                Set ProcessedIP = BinaryipAddress.BroadcastAddress().Offset(parseOffset(Offset))
            
                'Post Process the output
                outputArray(i, j) = FormatOutput(ProcessedIP, DisplayFormat)

nextCell:
            Next j
        Next i
    
        ipSubBroadcast = outputArray
        
    ElseIf TypeName(inputVal) = "String" Then
        
        On Error GoTo ErrorHandlerString
        
        'Convert the input into a Binary pair
        Set BinaryipAddress = ParseIP(CStr(inputVal))
    
        'Binary Function
        Set ProcessedIP = BinaryipAddress.BroadcastAddress().Offset(parseOffset(Offset))
    
        'Post Process the output
        ipSubBroadcast = FormatOutput(ProcessedIP, DisplayFormat)

    Else
        ' If the input is not a range or a string, return an error message
        ipSubBroadcast = "Invalid input: must be a range or a string"
    End If
    
    Exit Function

ErrorHandlerRange:
                outputArray(i, j) = "#VALUE!"
                Resume nextCell 'Clear the error and continues the executionat the label

ErrorHandlerString:
        ipSubBroadcast = CVErr(xlErrValue)

End Function


Public Function ipSubPrev(inputVal As Variant, Optional DisplayFormat As Long = 1, _
                          Optional Offset As String = vbNullString, Optional Summary As String = vbNullString) As Variant

    ' Returns the previous subnet with the same prefix length

    Dim BinaryipAddress As clsIP
    Dim ProcessedIP As clsIP
    Dim SummaryIPAddress As clsIP
    Dim i As Long, j As Long
    Dim outputArray() As Variant

    If TypeName(inputVal) = "Range" Then
        
        On Error GoTo ErrorHandlerRange
        
        ReDim outputArray(1 To inputVal.Rows.Count, 1 To inputVal.Columns.Count)
        
        For i = 1 To inputVal.Rows.Count
            For j = 1 To inputVal.Columns.Count
                
                'Convert the input into a Binary pair
                Set BinaryipAddress = ParseIP(CStr(inputVal.Cells(i, j).Value))
            
                'Binary Function
                If Summary <> vbNullString Then
                    Set SummaryIPAddress = ParseIP(Summary).SubAddress()
                    Set ProcessedIP = BinaryipAddress.SubPrev().Offset(parseOffset(Offset), False).Summary(SummaryIPAddress)
                Else
                    Set ProcessedIP = BinaryipAddress.SubPrev().Offset(parseOffset(Offset), False)
                End If
            
                'Post Process the output
                outputArray(i, j) = FormatOutput(ProcessedIP, DisplayFormat)

nextCell:
            Next j
        Next i
    
        ipSubPrev = outputArray
        
    ElseIf TypeName(inputVal) = "String" Then
        
        On Error GoTo ErrorHandlerString
        
        'Convert the input into a Binary pair
        Set BinaryipAddress = ParseIP(CStr(inputVal))
    
        'Binary Function
        If Summary <> vbNullString Then
            Set SummaryIPAddress = ParseIP(Summary).SubAddress()
            Set ProcessedIP = BinaryipAddress.SubPrev().Offset(parseOffset(Offset), False).Summary(SummaryIPAddress)
        Else
            Set ProcessedIP = BinaryipAddress.SubPrev().Offset(parseOffset(Offset), False)
        End If
    
        'Post Process the output
        ipSubPrev = FormatOutput(ProcessedIP, DisplayFormat)

    Else
        ' If the input is not a range or a string, return an error message
        ipSubPrev = "Invalid input: must be a range or a string"
    End If
    
    Exit Function

ErrorHandlerRange:
                outputArray(i, j) = "#VALUE!"
                Resume nextCell 'Clear the error and continues the executionat the label

ErrorHandlerString:
        ipSubPrev = CVErr(xlErrValue)

End Function

Public Function ipSubNext(inputVal As Variant, Optional DisplayFormat As Long = 1, _
                          Optional Offset As String = vbNullString, Optional Summary As String = vbNullString) As Variant

    ' Returns the next subnet with the same prefix length

    Dim BinaryipAddress As clsIP
    Dim ProcessedIP As clsIP
    Dim SummaryIPAddress As clsIP
    Dim i As Long, j As Long
    Dim outputArray() As Variant

    If TypeName(inputVal) = "Range" Then
        
        On Error GoTo ErrorHandlerRange
        
        ReDim outputArray(1 To inputVal.Rows.Count, 1 To inputVal.Columns.Count)
        
        For i = 1 To inputVal.Rows.Count
            For j = 1 To inputVal.Columns.Count
                
                'Convert the input into a Binary pair
                Set BinaryipAddress = ParseIP(CStr(inputVal.Cells(i, j).Value))
            
                'Binary Function
                If Summary <> vbNullString Then
                    Set SummaryIPAddress = ParseIP(Summary).SubAddress()
                    Set ProcessedIP = BinaryipAddress.SubNext().Offset(parseOffset(Offset), False).Summary(SummaryIPAddress)
                Else
                    Set ProcessedIP = BinaryipAddress.SubNext().Offset(parseOffset(Offset), False)
                End If
            
                'Post Process the output
                outputArray(i, j) = FormatOutput(ProcessedIP, DisplayFormat)

nextCell:
            Next j
        Next i
    
        ipSubNext = outputArray
        
    ElseIf TypeName(inputVal) = "String" Then
        
        On Error GoTo ErrorHandlerString
        
        'Convert the input into a Binary pair
        Set BinaryipAddress = ParseIP(CStr(inputVal))
    
        'Binary Function
        If Summary <> vbNullString Then
            Set SummaryIPAddress = ParseIP(Summary).SubAddress()
            Set ProcessedIP = BinaryipAddress.SubNext().Offset(parseOffset(Offset), False).Summary(SummaryIPAddress)
        Else
            Set ProcessedIP = BinaryipAddress.SubNext().Offset(parseOffset(Offset), False)
        End If
    
        'Post Process the output
        ipSubNext = FormatOutput(ProcessedIP, DisplayFormat)

    Else
        ' If the input is not a range or a string, return an error message
        ipSubNext = "Invalid input: must be a range or a string"
    End If
    
    Exit Function

ErrorHandlerRange:
                outputArray(i, j) = "#VALUE!"
                Resume nextCell 'Clear the error and continues the executionat the label

ErrorHandlerString:
        ipSubNext = CVErr(xlErrValue)

End Function





''''''''''''''''''''''''''''''
' Summary Functions
''''''''''''''''''''''''''''''

Public Function ipSumFirstSub(inputVal As Variant, PrefixLength As Long, Optional DisplayFormat As Long = 1, _
                          Optional Offset As String = vbNullString) As Variant

    Dim BinaryipAddress As clsIP
    Dim ProcessedIP As clsIP
    Dim i As Long, j As Long
    Dim outputArray() As Variant

    If TypeName(inputVal) = "Range" Then
        
        On Error GoTo ErrorHandlerRange
        
        ReDim outputArray(1 To inputVal.Rows.Count, 1 To inputVal.Columns.Count)
        
        For i = 1 To inputVal.Rows.Count
            For j = 1 To inputVal.Columns.Count
                
                'Convert the input into a Binary pair
                Set BinaryipAddress = ParseIP(CStr(inputVal.Cells(i, j).Value))
            
                'Binary Function
                Set ProcessedIP = BinaryipAddress.SumFirstSub(PrefixLength).Offset(parseOffset(Offset), False)
            
                'Post Process the output
                outputArray(i, j) = FormatOutput(ProcessedIP, DisplayFormat)

nextCell:
            Next j
        Next i
    
        ipSumFirstSub = outputArray
        
    ElseIf TypeName(inputVal) = "String" Then
        
        On Error GoTo ErrorHandlerString
        
        'Convert the input into a Binary pair
        Set BinaryipAddress = ParseIP(CStr(inputVal))
    
        'Binary Function
        Set ProcessedIP = BinaryipAddress.SumFirstSub(PrefixLength).Offset(parseOffset(Offset), False)
    
        'Post Process the output
        ipSumFirstSub = FormatOutput(ProcessedIP, DisplayFormat)

    Else
        ' If the input is not a range or a string, return an error message
        ipSumFirstSub = "Invalid input: must be a range or a string"
    End If
    
    Exit Function

ErrorHandlerRange:
                outputArray(i, j) = "#VALUE!"
                Resume nextCell 'Clear the error and continues the executionat the label

ErrorHandlerString:
        ipSumFirstSub = CVErr(xlErrValue)

End Function


Public Function ipSumSubX(inputVal As Variant, PrefixLength As Long, SubNumber As String, Optional DisplayFormat As Long = 1, _
                          Optional Offset As String = vbNullString) As Variant
    
    ' Return the X'th subnet of a larger summary
    
    Dim BinaryipAddress As clsIP
    Dim ProcessedIP As clsIP
    Dim BinSubNumber As String
    Dim i As Long, j As Long
    Dim outputArray() As Variant

    If TypeName(inputVal) = "Range" Then
        
        On Error GoTo ErrorHandlerRange
        
        ReDim outputArray(1 To inputVal.Rows.Count, 1 To inputVal.Columns.Count)
        
        For i = 1 To inputVal.Rows.Count
            For j = 1 To inputVal.Columns.Count
                
                'Convert the input into a Binary pair
                Set BinaryipAddress = ParseIP(CStr(inputVal.Cells(i, j).Value))
            
                'Binary Function
                BinSubNumber = parseOffset(SubNumber)
                Set ProcessedIP = BinaryipAddress.SumSubX(PrefixLength, BinSubNumber).Offset(parseOffset(Offset), False)
                
                'Post Process the output
                outputArray(i, j) = FormatOutput(ProcessedIP, DisplayFormat)

nextCell:
            Next j
        Next i
    
        ipSumSubX = outputArray
        
    ElseIf TypeName(inputVal) = "String" Then
        
        On Error GoTo ErrorHandlerString
        
        'Convert the input into a Binary pair
        Set BinaryipAddress = ParseIP(CStr(inputVal))
    
        'Binary Function
        BinSubNumber = parseOffset(SubNumber)
        Set ProcessedIP = BinaryipAddress.SumSubX(PrefixLength, BinSubNumber).Offset(parseOffset(Offset), False)
        
        'Post Process the output
        ipSumSubX = FormatOutput(ProcessedIP, DisplayFormat)

    Else
        ' If the input is not a range or a string, return an error message
        ipSumSubX = "Invalid input: must be a range or a string"
    End If
    
    Exit Function

ErrorHandlerRange:
                outputArray(i, j) = "#VALUE!"
                Resume nextCell 'Clear the error and continues the executionat the label

ErrorHandlerString:
        ipSumSubX = CVErr(xlErrValue)

End Function


Public Function ipSumSubY(inputVal As Variant, PrefixLength As Long, SubNumber As String, Optional DisplayFormat As Long = 1, _
                          Optional Offset As String = vbNullString) As Variant
    
    ' Return the Y'th subnet of a larger summary from the end
    
    Dim BinaryipAddress As clsIP
    Dim ProcessedIP As clsIP
    Dim BinSubNumber As String
    Dim i As Long, j As Long
    Dim outputArray() As Variant

    If TypeName(inputVal) = "Range" Then
        
        On Error GoTo ErrorHandlerRange
        
        ReDim outputArray(1 To inputVal.Rows.Count, 1 To inputVal.Columns.Count)
        
        For i = 1 To inputVal.Rows.Count
            For j = 1 To inputVal.Columns.Count
                
                'Convert the input into a Binary pair
                Set BinaryipAddress = ParseIP(CStr(inputVal.Cells(i, j).Value))
            
                'Binary Function
                BinSubNumber = parseOffset(SubNumber)
                Set ProcessedIP = BinaryipAddress.SumSubY(PrefixLength, BinSubNumber).Offset(parseOffset(Offset), False)
                
                'Post Process the output
                outputArray(i, j) = FormatOutput(ProcessedIP, DisplayFormat)

nextCell:
            Next j
        Next i
    
        ipSumSubY = outputArray
        
    ElseIf TypeName(inputVal) = "String" Then
        
        On Error GoTo ErrorHandlerString
        
        'Convert the input into a Binary pair
        Set BinaryipAddress = ParseIP(CStr(inputVal))
    
        'Binary Function
        BinSubNumber = parseOffset(SubNumber)
        Set ProcessedIP = BinaryipAddress.SumSubY(PrefixLength, BinSubNumber).Offset(parseOffset(Offset), False)
        
        'Post Process the output
        ipSumSubY = FormatOutput(ProcessedIP, DisplayFormat)

    Else
        ' If the input is not a range or a string, return an error message
        ipSumSubY = "Invalid input: must be a range or a string"
    End If
    
    Exit Function

ErrorHandlerRange:
                outputArray(i, j) = "#VALUE!"
                Resume nextCell 'Clear the error and continues the executionat the label

ErrorHandlerString:
        ipSumSubY = CVErr(xlErrValue)

End Function


Public Function ipSumLastSub(inputVal As Variant, PrefixLength As Long, Optional DisplayFormat As Long = 1, _
                          Optional Offset As String = vbNullString) As Variant
    
    ' Return the Y'th subnet of a larger summary from the end
    
    Dim BinaryipAddress As clsIP
    Dim ProcessedIP As clsIP
    Dim i As Long, j As Long
    Dim outputArray() As Variant

    If TypeName(inputVal) = "Range" Then
        
        On Error GoTo ErrorHandlerRange
        
        ReDim outputArray(1 To inputVal.Rows.Count, 1 To inputVal.Columns.Count)
        
        For i = 1 To inputVal.Rows.Count
            For j = 1 To inputVal.Columns.Count
                
                'Convert the input into a Binary pair
                Set BinaryipAddress = ParseIP(CStr(inputVal.Cells(i, j).Value))
            
                'Binary Function
                Set ProcessedIP = BinaryipAddress.SumLastSub(PrefixLength).Offset(parseOffset(Offset), False)
                
                'Post Process the output
                outputArray(i, j) = FormatOutput(ProcessedIP, DisplayFormat)

nextCell:
            Next j
        Next i
    
        ipSumLastSub = outputArray
        
    ElseIf TypeName(inputVal) = "String" Then
        
        On Error GoTo ErrorHandlerString
        
        'Convert the input into a Binary pair
        Set BinaryipAddress = ParseIP(CStr(inputVal))
    
        'Binary Function
        Set ProcessedIP = BinaryipAddress.SumLastSub(PrefixLength).Offset(parseOffset(Offset), False)
        
        'Post Process the output
        ipSumLastSub = FormatOutput(ProcessedIP, DisplayFormat)

    Else
        ' If the input is not a range or a string, return an error message
        ipSumLastSub = "Invalid input: must be a range or a string"
    End If
    
    Exit Function

ErrorHandlerRange:
                outputArray(i, j) = "#VALUE!"
                Resume nextCell 'Clear the error and continues the executionat the label

ErrorHandlerString:
        ipSumLastSub = CVErr(xlErrValue)

End Function


Public Function ipSumCheck(IPInputV As Variant, SummaryInput As Variant) As Variant

' Checks whether IPInput is included within SummaryInput
' SummaryInput can be a single summary (string) to check or can be a range of cells
' containing multiple summaries
' If IPInput is a cell, if the cell is included in the SummaryInput range it will be
' ignored and compared with all the other cells in the range.

    Dim BinaryipAddress As clsIP
    Dim BinarySummaryAddress As clsIP
    Dim c As Range
    Dim InputRange As Range
    Dim IPInput As String

    On Error GoTo ErrorHandler
    
    If TypeName(IPInputV) = "Range" Then
        ' If IPInput is a range then we check it is a single cell and exctract its value
        If IPInputV.Count > 1 Then Err.Raise vbObjectError + 1010, "ipSumCheck", "IPInput can be only one cell"
        IPInput = IPInputV.Value2
    Else
        ' If it is not a range it must be a string
        IPInput = IPInputV
    End If
    
    'Convert the input into a Binary IPAddress pair
    Set BinaryipAddress = ParseIP(IPInput)

    ' If no subnet mask is given for IPInput, we assume /32
    ' Inverse behaviour of ParseIP function so we manually change the mask
    If BinaryipAddress.PrefixLength = 0 Then BinaryipAddress.PrefixLength = BinaryipAddress.IPLength
    
    If TypeName(SummaryInput) = "Range" Then
    
        ipSumCheck = False
        
        For Each c In SummaryInput.Cells
            If Not IsEmpty(c) Then
                Set BinarySummaryAddress = ParseIP(c.Value)
                If BinaryipAddress.SumCheck(BinarySummaryAddress) = True Then
                    ipSumCheck = True
                    Exit For
                End If
            End If
        Next c
    
    Else
    
        Set BinarySummaryAddress = ParseIP(SummaryInput)
        ipSumCheck = BinarySummaryAddress.SumCheck(BinarySummaryAddress)
    
    End If
    
    Exit Function

ErrorHandler:
    ipSumCheck = errorHandling()

End Function


Public Function ipSumSubCount(IPInput As String, PrefixLength As Integer) As Variant

    ' Returns the number of subnets of PrefixLength size within the summary

    Dim BinaryipAddress As clsIP

    On Error GoTo ErrorHandler

   ' Convert the input into a Binary IPAddress pair
    Set BinaryipAddress = ParseIP(IPInput)
    
    If (PrefixLength < BinaryipAddress.PrefixLength) Then Err.Raise 1020, "ipSumSubCount", "#PrefixLength bigger than summary!"
    
    If BinaryipAddress.IPv6 = True Then
        If (PrefixLength > 128) Then Err.Raise vbObjectError + 1010, "ipSumSubCount", "#PrefixLength > 128!"
    Else
        If (PrefixLength > 32) Then Err.Raise vbObjectError + 1010, "ipSumSubCount", "#PrefixLength > 32!"
    End If
    
    ipSumSubCount = 2 ^ (PrefixLength - BinaryipAddress.PrefixLength)
    
    Exit Function

ErrorHandler:
    ipSumSubCount = errorHandling()

End Function





''''''''''''''''''''''''''''''
' Subnet Functions
''''''''''''''''''''''''''''''

Public Function ipOffset(inputVal As Variant, Offset As String, Optional DisplayFormat As Long = 1 _
                          ) As Variant
    
    ' Returns the number of subnets of PrefixLength size within the summary
    
    Dim BinaryipAddress As clsIP
    Dim ProcessedIP As clsIP
    Dim i As Long, j As Long
    Dim outputArray() As Variant

    If TypeName(inputVal) = "Range" Then
        
        On Error GoTo ErrorHandlerRange
        
        ReDim outputArray(1 To inputVal.Rows.Count, 1 To inputVal.Columns.Count)
        
        For i = 1 To inputVal.Rows.Count
            For j = 1 To inputVal.Columns.Count
                
                'Convert the input into a Binary pair
                Set BinaryipAddress = ParseIP(CStr(inputVal.Cells(i, j).Value))
            
                'Binary Function
                Set ProcessedIP = BinaryipAddress.Offset(parseOffset(Offset), False)
                
                'Post Process the output
                outputArray(i, j) = FormatOutput(ProcessedIP, DisplayFormat)

nextCell:
            Next j
        Next i
    
        ipOffset = outputArray
        
    ElseIf TypeName(inputVal) = "String" Then
        
        On Error GoTo ErrorHandlerString
        
        'Convert the input into a Binary pair
        Set BinaryipAddress = ParseIP(CStr(inputVal))
    
        'Binary Function
        Set ProcessedIP = BinaryipAddress.Offset(parseOffset(Offset), False)
        
        'Post Process the output
        ipOffset = FormatOutput(ProcessedIP, DisplayFormat)

    Else
        ' If the input is not a range or a string, return an error message
        ipOffset = "Invalid input: must be a range or a string"
    End If
    
    Exit Function

ErrorHandlerRange:
                outputArray(i, j) = "#VALUE!"
                Resume nextCell 'Clear the error and continues the executionat the label

ErrorHandlerString:
        ipOffset = CVErr(xlErrValue)

End Function


Public Function ipVersion() As String

    ipVersion = "IP Functions v" & gIPfversion & " ©2013-17 N. Grison http://tiny.cc/xlsipf"

End Function





''''''''''''''''''''''''''''''
' Array Functions
''''''''''''''''''''''''''''''

Public Function ipSubnets(Summary As String, PrefixLength As Long, Optional DisplayFormat As Long = 2) As Variant

    Dim BinarySummary As clsIP
    Dim ProcessedIP As clsIP
    Dim LastAddress As String
    Dim Subnets() As Variant
    Dim ButtonChosen As Integer
    Dim numberOfSubnets As Long

    On Error GoTo ErrorHandler
    
    Set BinarySummary = ParseIP(Summary)
    
    If BinarySummary.PrefixLength > PrefixLength Then
        Err.Raise 1010, "ipSubnets", "#Summary too small!"
    Else
    
        If (PrefixLength - BinarySummary.PrefixLength) > 10 Then
        ' 2 ^ 10 = 1024 Subnets
            ButtonChosen = MsgBox("More than 1024 subnets, this may take a while." & vbLf & "Do you want to continue?", vbQuestion + vbYesNo + vbDefaultButton1, "IP Subnets")
    
            If ButtonChosen = vbNo Then Exit Function

        End If
    
        LastAddress = BinarySummary.LastAddress.Address
        
        ' Get the first subnet
        Set ProcessedIP = BinarySummary.SumFirstSub(PrefixLength)
        
        ReDim Preserve Subnets(1 To 1)
        Subnets(1) = FormatOutput(ProcessedIP, DisplayFormat)
        
        Set ProcessedIP = ProcessedIP.SubNext
    
        ' ProcessedIP.Address = lastAddress if we are subnetting into /32
        Do While ProcessedIP.Address <= LastAddress
            ReDim Preserve Subnets(1 To UBound(Subnets) + 1)
            Subnets(UBound(Subnets)) = FormatOutput(ProcessedIP, DisplayFormat)
            Set ProcessedIP = ProcessedIP.SubNext
        Loop
    End If
    
    If UBound(Subnets) = 1 Then
        ' If we return a single value, Excel assigns it to all destination cells
        ' We instead return a 2nd cell with an error so that only the first cell is populated
        ReDim Preserve Subnets(1 To 2)
        Subnets(2) = CVErr(xlErrNA)
    End If
    
    If IsObject(Application.caller) And Application.caller.Rows.Count > Application.caller.Columns.Count Then
        ipSubnets = Application.WorksheetFunction.Transpose(Subnets)
    Else
        ipSubnets = Subnets
    End If
    
    Exit Function

ErrorHandler:
    ipSubnets = errorHandling()

End Function


Public Function ipHosts(Net As String, Optional DisplayFormat As Long = 2) As Variant

    ' Array function that returns all the host addresses in a subnet

    Dim BinaryIP As clsIP
    Dim ProcessedIP As clsIP
    Dim LastAddress As String
    Dim Hosts() As Variant
    Dim i As Long
    Dim addressLength As Long
    Dim ButtonChosen As Integer

    On Error GoTo ErrorHandler
    
    'Function
    Set BinaryIP = ParseIP(Net)
    
    If BinaryIP.IPLength = BinaryIP.PrefixLength Then
        ' We are dealing with a host subnet
        ReDim Preserve Hosts(1 To 2)
        Hosts(1) = BinaryIP.Address
        Hosts(1) = FormatOutput(BinaryIP, DisplayFormat)
        Hosts(2) = CVErr(xlErrNA)
        ipHosts = Hosts
    
    ElseIf BinaryIP.PrefixLength = BinaryIP.IPLength - 1 Then
    
        ReDim Preserve Hosts(1 To 2)
        Hosts(1) = "No host address"
        Hosts(2) = CVErr(xlErrNA)
        ipHosts = Hosts

    Else

        If (BinaryIP.IPLength - BinaryIP.PrefixLength) >= 10 Then
            ' 2 ^ 10 = 1024 Hosts
            ButtonChosen = MsgBox("More than 1024 hosts, this may take a while." & vbLf & "Do you want to continue?", vbQuestion + vbYesNo + vbDefaultButton1, "IP Hosts")

            If ButtonChosen = vbNo Then Exit Function

        End If
        
        ' Get the last host address
        LastAddress = BinaryIP.LastHost.Address
    
        Set ProcessedIP = BinaryIP.FirstHost
        
        ReDim Preserve Hosts(1 To 1)

        Hosts(1) = FormatOutput(ProcessedIP, DisplayFormat)
    
        Do While ProcessedIP.Address < LastAddress
            ReDim Preserve Hosts(1 To UBound(Hosts) + 1)
            Set ProcessedIP = ProcessedIP.NextHost
            Hosts(UBound(Hosts)) = FormatOutput(ProcessedIP, DisplayFormat)
        Loop
    End If
    
    If IsObject(Application.caller) And Application.caller.Rows.Count > Application.caller.Columns.Count Then
        ipHosts = Application.WorksheetFunction.Transpose(Hosts)
    Else
        ipHosts = Hosts
    End If
    
    Exit Function

ErrorHandler:
    ipHosts = errorHandling()

End Function


Public Function ipSummarise(SourceRange As Range, Optional DisplayFormat As Long = 2, Optional Merge As Boolean = True) As Variant
    
    ' Return Array without overlapped subnets
    
    On Error Resume Next
        
    Dim cel As Range
    Dim selectedRange As Range
    Dim Subnets() As clsIP
    Dim OutputArr() As Variant
    Dim i As Long
    Dim j As Long
    Dim sumSubIP As clsIP
    Dim sumLastIP As clsIP
    Dim ChangeRecorded As Boolean
    
    ReDim Subnets(1 To SourceRange.Cells.Count)
    ReDim OutputArr(1 To SourceRange.Cells.Count)
    
    ' create a clean reference array Subnets with all subnets IPs
    For i = LBound(Subnets) To UBound(Subnets)
        Set Subnets(i) = ParseIP(SourceRange.Cells(i).Value2).SubAddress
    Next i
    
    ' sort the array
    Call fnIPSort(Subnets)

    ' First remove the subnets that are already included in larger ones
    
    For i = LBound(Subnets) To UBound(Subnets)
        ' compare each subnet with all in Subnets() array except itself to find a larger one if it exists
        ' and remove duplicates
        
        For j = LBound(Subnets) To UBound(Subnets)
            ' if we have reached the current value itself or empty subnet address we skip to next
            ' As the array is sorted, when we get to an address that's higher we know we can't match anything further
            If (j <> i And Subnets(j).Address <> "" And Subnets(j).Address <= Subnets(i).Address) Then
            
                If Subnets(i).SumCheck(Subnets(j)) = True Then
                    ' we have an overlap so we remove the current subnet from Subnets() array
                    Subnets(i).Address = ""
                
                    'no need to compare with the rest, we move to the next subnet
                    j = UBound(Subnets)
                End If
            End If
        Next j
        
    Next i
    
    ' Merge subnets that can be merged
    
    If Merge = True Then
    
        'repeat until we do not have any more changes
    
        If UBound(Subnets) - LBound(Subnets) > 1 Then
            Do
                ChangeRecorded = False
            
                For i = LBound(Subnets) To UBound(Subnets) - 1
            
                    If Subnets(i).Address <> "" Then
                        ' find the next non empty item
                        j = i + 1
                        Do Until Subnets(j).Address <> "" Or j = UBound(Subnets)
                            j = j + 1
                        Loop
                    
                        Set sumSubIP = New clsIP
                        sumSubIP.Address = Subnets(i).Address
                        sumSubIP.PrefixLength = Subnets(i).PrefixLength - 1
                    
                        ' compare subnet address and last address
                        ' as we compare n and n + 1, no need to apply ipSubAddress
                        If sumSubIP.Address = Subnets(i).Address And sumSubIP.LastAddress.Address = Subnets(j).LastAddress.Address Then
                                                    
                            ' we have found two summarisable subnets
                            Set Subnets(i) = sumSubIP
                        
                            Subnets(j).Address = ""
                            Subnets(j).Mask = ""
                        
                            ChangeRecorded = True
                    
                        End If
                    End If
                Next i
            Loop While ChangeRecorded = True
        End If
    End If
    
    
    ' Generate the output array
    
    j = 1
    
    ' We put all the values at the top
    
    For i = LBound(Subnets) To UBound(Subnets)
        If Subnets(i).Address <> "" Then
            OutputArr(j) = FormatOutput(Subnets(i), DisplayFormat)
            j = j + 1
        End If
    Next i

    ' And empty the remaining values
    
    For i = j To UBound(OutputArr)
        OutputArr(i) = CVErr(xlErrNA)
    Next i
    
    If IsObject(Application.caller) Then
        If Application.caller.Rows.Count > Application.caller.Columns.Count Then
            ipSummarise = Application.WorksheetFunction.Transpose(OutputArr)
        Else
            ipSummarise = OutputArr
        End If
    Else
        ipSummarise = OutputArr
    End If

End Function


Public Function ipSort(SourceRange As Range, Optional DisplayFormat As Long = 2) As Variant
    
    ' Return a sorted Array
    
    Dim Subnets() As clsIP
    Dim OutputArr() As Variant
    Dim i As Long
    Dim j As Long
    
    On Error Resume Next
  
    ReDim Subnets(1 To SourceRange.Cells.Count)
    ReDim OutputArr(1 To SourceRange.Cells.Count)
    
    ' create a clean reference array Subnets with all subnets IPs as IPBins

    For i = LBound(Subnets) To UBound(Subnets)
        Set Subnets(i) = ParseIP(SourceRange.Cells(i).Value2)
    Next i

    ' sort the array
    Call fnIPSort(Subnets)

    ' Generate the output array
    
    j = 1
    
    ' We put all the values at the top
    For i = LBound(Subnets) To UBound(Subnets)
        If Subnets(i).Address <> "" Then
            OutputArr(j) = FormatOutput(Subnets(i), DisplayFormat)
            j = j + 1
        End If
    Next i

    ' And empty the remaining values
    
    For i = j To UBound(OutputArr)
        OutputArr(i) = CVErr(xlErrNA)
    Next i
    
    If IsObject(Application.caller) Then
        If Application.caller.Rows.Count > Application.caller.Columns.Count Then
            ipSort = Application.WorksheetFunction.Transpose(OutputArr)
        Else
            ipSort = OutputArr
        End If
    Else
        ipSort = OutputArr
    End If

End Function


Public Function ipRoute(SourceRange As Range, RouteRange As Range, Optional DisplayFormat As Long = 2) As Variant
    
    ' For each source address, returns the best route in the RouteRange
        
    Dim SourceSubnets() As clsIP
    Dim ListSubnets() As clsIP
    Dim OutputSubnets() As clsIP
    Dim OutputArr() As Variant
    Dim i As Long
    Dim j As Long
    Dim ProvisoBestRoute As clsIP
    
    ReDim SourceSubnets(1 To SourceRange.Cells.Count)
    ReDim ListSubnets(1 To RouteRange.Cells.Count)
    ReDim OutputSubnets(1 To SourceRange.Cells.Count)
    ReDim OutputArr(1 To SourceRange.Cells.Count)
    
    On Error Resume Next
    
    ' create clean reference arrays Subnets with all subnets IPs as IPBins

    For i = LBound(SourceSubnets) To UBound(SourceSubnets)
        Set SourceSubnets(i) = ParseIP(SourceRange.Cells(i).Value2).SubAddress
    Next i
    
    For i = LBound(ListSubnets) To UBound(ListSubnets)
        Set ListSubnets(i) = ParseIP(RouteRange.Cells(i).Value2).SubAddress
    Next i
    
    ' Compare each subnet in SourceSubnets with the subnets in ListSubnets
    
    For i = LBound(SourceSubnets) To UBound(SourceSubnets)
        If Not SourceSubnets(i) Is Nothing Then
            ' compare each subnet with all in Subnets() array except itself to find a larger one if it exists
            Set ProvisoBestRoute = New clsIP
            
            For j = LBound(ListSubnets) To UBound(ListSubnets)
    
                If SourceSubnets(i).SumCheck(ListSubnets(j)) = True Then
                    ' we have a match
                    Set ProvisoBestRoute = ListSubnets(j).SubAddress
                    
                    If OutputSubnets(i) Is Nothing Or ProvisoBestRoute.Mask > OutputSubnets(i).Mask Then
                        Set OutputSubnets(i) = ProvisoBestRoute
                    End If
    
                End If
                            
            Next j
        End If
        
    Next i
    
    ' Generate the output array
    
    For i = LBound(OutputSubnets) To UBound(OutputSubnets)
        If OutputSubnets(i) Is Nothing Or OutputSubnets(i).Address = "" Then
            OutputArr(i) = CVErr(xlErrNA)
        Else
            OutputArr(i) = FormatOutput(OutputSubnets(i), DisplayFormat)
        End If
    Next i
    
    If IsObject(Application.caller) Then
        If Application.caller.Rows.Count > Application.caller.Columns.Count Then
            ipRoute = Application.WorksheetFunction.Transpose(OutputArr)
        Else
            ipRoute = OutputArr
        End If
    Else
        ipRoute = OutputArr
    End If

End Function


''''''''''''''''''''''''''''''
' Other Functions
''''''''''''''''''''''''''''''


Private Function appendMask(inputString As String) As String
    ' Add the prefix length to a classfull subnet
    ' i.e.10.0.0.0 > 10.0.0.0/8
    ' 198.1.0.0 > 198.1.0.0/24

    Dim parts() As String
    
    On Error GoTo ErrorHandler
    
    If InStr(inputString, "/") > 0 Then
        appendMask = inputString
        Exit Function
    End If
    
    parts = Split(inputString, ".")

    ' Check if the first part exists and is numeric
    If UBound(parts) = 3 And IsNumeric(parts(0)) Then
        
        If CInt(parts(0)) <= 126 Then
            appendMask = inputString & "/8"
        ElseIf CInt(parts(0)) <= 191 Then
            appendMask = inputString & "/16"
        ElseIf CInt(parts(0)) <= 223 Then
            appendMask = inputString & "/24"
       Else
            appendMask = inputString
        End If

    Else
        appendMask = inputString
    End If
    
    Exit Function
    
ErrorHandler:
    
    appendMask = errorHandling()

End Function

Public Function ipClassless(inputVal As Variant) As Variant

    ' Add the prefix length to a classfull subnet address
    ' i.e.10.0.0.0 > 10.0.0.0/8
    ' 198.1.0.0 > 198.1.0.0/24

    Dim i As Long, j As Long
    Dim outputArray() As Variant
    
    If TypeName(inputVal) = "Range" Then

        On Error GoTo ErrorHandlerRange

        ReDim outputArray(1 To inputVal.Rows.Count, 1 To inputVal.Columns.Count)
    
        For i = 1 To inputVal.Rows.Count
            For j = 1 To inputVal.Columns.Count
                outputArray(i, j) = appendMask(CStr(inputVal.Cells(i, j).Value))
nextCell:
            Next j
        Next i
    
        ipClassless = outputArray
        
    ElseIf TypeName(inputVal) = "String" Then
        
        On Error GoTo ErrorHandlerString
        
        ipClassless = appendMask(CStr(inputVal))
    Else
        ' If the input is not a range or a string, return an error message
        ipClassless = "Invalid input: must be a range or a string"
    End If
    
    Exit Function
    
ErrorHandlerRange:
                outputArray(i, j) = "#VALUE!"
                Resume nextCell 'Clear the error and continues the executionat the label

ErrorHandlerString:
        ipClassless = CVErr(xlErrValue)

End Function



''''''''''''''''''''''''''''''
' Macros
''''''''''''''''''''''''''''''

Public Sub IP_Reformat()
    
    ' Change the IP format of the selected cells
    ' Enter the format as three digits as in the main ip functions
    
    Dim cel As Range
    Dim selectedRange As Range
    Dim DisplayFormat As Long
    
    Application.ScreenUpdating = False
    
    On Error Resume Next
    
    Set selectedRange = Application.Selection
    
    DisplayFormat = InputBox("Desired IP format (1-3 digits):")
        
    For Each cel In selectedRange.Cells
        cel.Value2 = FormatOutput(ParseIP(cel.Value2), DisplayFormat)
    Next cel
    
    Application.ScreenUpdating = True
    
End Sub


Public Sub IP_SortAndReformat()
    
    ' Change the IP format of the selected cells
    ' Enter the format as three digits as in the main ip functions
    
    Dim Subnets() As clsIP
    Dim OutputArr() As Variant
    Dim cel As Range
    Dim selectedRange As Range
    Dim DisplayFormat As Long
    Dim i As Long
    
    Application.ScreenUpdating = False
    
    Set selectedRange = Application.Selection
    
    DisplayFormat = InputBox("Desired IP format (1-3 digits):")
        
    ReDim Subnets(1 To selectedRange.Cells.Count)
    ReDim OutputArr(1 To selectedRange.Cells.Count)

    i = 1
    
    ' create a reference array Subnets with all IPBins
    
    For Each cel In selectedRange.Cells
        Set Subnets(i) = ParseIP(cel.Value2)
        i = i + 1
    Next cel
    
    ' sort the array in alphabetical order
    Call fnIPSort(Subnets)

    For i = LBound(Subnets) To UBound(Subnets)
        selectedRange.Cells(i) = FormatOutput(Subnets(i), DisplayFormat)
    Next i
    
    Application.ScreenUpdating = True
    
End Sub


Public Sub IP_SortAndReformat2()
    
    ' Change the IP format of the selected cells
    ' Enter the format as three digits as in the main ip functions
    
    Dim Subnets() As clsIP
    Dim InputArr As Variant
    Dim OutputArr() As Variant
    Dim cel As Range
    Dim selectedRange As Range
    Dim DisplayFormat As Long
    Dim i As Long
    
    Application.ScreenUpdating = False
    
    Set selectedRange = Application.Selection
    
    ' Transfer Range to array
    InputArr = selectedRange.Value
        
    ' Convert first column to clsIP objects
    For i = 1 To UBound(InputArr, 1)
        Set InputArr(i, 1) = ParseIP(InputArr(i, 1))
        i = i + 1
    Next i
    
    Stop
'    DisplayFormat = InputBox("Desired IP format (1-3 digits):")
        
'    ReDim Subnets(1 To selectedRange.Cells.Count)
'    ReDim OutputArr(1 To selectedRange.Cells.Count)

'    i = 1
    
    ' create a reference array Subnets with all IPBins
    
'    For Each cel In selectedRange.Cells
'        Set Subnets(i) = ParseIP(cel.Value2)
'        i = i + 1
'    Next cel
    
    ' sort the array in alphabetical order
'    Call fnIPSort(Subnets)

'    For i = LBound(Subnets) To UBound(Subnets)
'        selectedRange.Cells(i) = FormatOutput(Subnets(i), DisplayFormat)
'    Next i
    
    Application.ScreenUpdating = True
    
End Sub


Public Sub IP_Summarise()
    
    ' Removes overlapped subnets from selected cells
    
    Dim DisplayFormat As Long
    Dim OutputArr() As Variant
    Dim selectedRange As Range
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    Application.ScreenUpdating = False
    
    DisplayFormat = InputBox("Desired IP format (three digits):")
    
    Set selectedRange = Application.Selection

    OutputArr = ipSummarise(selectedRange, DisplayFormat)
    
    ' Finally we paste the array into the range
    
    For i = LBound(OutputArr) To UBound(OutputArr)
        selectedRange.Cells(i).Value = OutputArr(i)
    Next i
    
    Application.ScreenUpdating = True

End Sub





''''''''''''''''''''''''''''''
' IP Internal Functions
''''''''''''''''''''''''''''''

Private Function ParseIP(ByVal strInput As String) As clsIP

    'Takes an IP address and optional mask as input and returns an IPBin object
    
    Dim varAddress As Variant
    Dim parts() As String
    
    Set ParseIP = New clsIP

    On Error GoTo ErrorHandler

    ' Clean the input
    strInput = Trim(strInput)

    If InStr(1, strInput, ":", vbTextCompare) = 0 Then
    
    
        ' We have an IPv4 Address
        
        varAddress = Split(strInput, "/", , vbTextCompare)
        
        If UBound(varAddress) = 1 Then
            ' x.x.x.x/x
            
            ParseIP.Address = cvIP4toBIN(varAddress(0))
            ParseIP.PrefixLength = CLng(varAddress(1))
            
        Else
            varAddress = Split(strInput, " ", , vbTextCompare)
            
            If UBound(varAddress) = 0 Then
                ' x.x.x.x
                
                ParseIP.Address = cvIP4toBIN(varAddress(0))
                
                'ParseIP.PrefixLength = Len(ParseIP.Address)
                
                parts = Split(varAddress(0), ".")

                If CInt(parts(0)) <= 126 Then
                    ParseIP.PrefixLength = 8
                ElseIf CInt(parts(0)) <= 191 Then
                    ParseIP.PrefixLength = 16
                ElseIf CInt(parts(0)) <= 223 Then
                    ParseIP.PrefixLength = 24
                End If

            Else
                ' x.x.x.x x.x.x.x
                
                ParseIP.Address = cvIP4toBIN(varAddress(0))
                ParseIP.Mask = cvIP4toBIN(varAddress(1))
                
            End If
            
        End If
        
    Else
    
        ' We have an IPv6 address
        
        varAddress = Split(strInput, "/", , vbTextCompare)
        
        If UBound(varAddress) = 0 Then
            ' ::
            
            ParseIP.Address = cvIP6toBIN(varAddress(0))
            ParseIP.PrefixLength = Len(varAddress(0))
            
        Else
            ' ::/x
            
            ParseIP.Address = cvIP6toBIN(varAddress(0))
            ParseIP.PrefixLength = varAddress(1)
            
        End If
    
    End If

    Exit Function

ErrorHandler:
    Err.Raise vbObjectError + 1010, "ParseIP", "#Parsing Error"

End Function


Private Function parseDisplayFormat(ByVal DisplayFormat As Long) As tIPFormat

    Dim OutputFormat As tIPFormat
    
    If DisplayFormat < 10 Then
        OutputFormat.QuadDotted = False
        OutputFormat.Padded = False
        OutputFormat.Format = DisplayFormat
    ElseIf DisplayFormat < 100 Then
        OutputFormat.QuadDotted = False
        OutputFormat.Padded = DisplayFormat Mod 10
        OutputFormat.Format = DisplayFormat \ 10
    ElseIf DisplayFormat < 1000 Then
        OutputFormat.QuadDotted = DisplayFormat Mod 10
        DisplayFormat = DisplayFormat \ 10
        OutputFormat.Padded = DisplayFormat Mod 10
        OutputFormat.Format = DisplayFormat \ 10
    Else
        Err.Raise vbObjectError + 1020, "parseDisplayFormat", "#Parameter error!"
    End If
    
    parseDisplayFormat = OutputFormat
    
    Exit Function

End Function


Private Function parseOffset(Offset As Variant) As String

    Dim BinaryOffset As Variant
    
  '  On Error GoTo ErrorHandler
    
    ' Determine offset type and convert to binary accordingly
    
    If Offset = vbNullString Then
        parseOffset = vbNullString
        Exit Function
    End If
        
    If Left(Offset, 2) = "0b" Then

        ' Offset is binary, we keep as is

        parseOffset = Right(Offset, Len(Offset) - 2)

    ElseIf Left(Offset, 2) = "0x" Then
        
        ' Offset is hexadecimal
        
        Offset = Right(Offset, Len(Offset) - 2)
        
        If Left(Offset, 1) = "-" Then
            parseOffset = "-"
            Offset = Right(Offset, Len(Offset) - 1)
        End If

        parseOffset = parseOffset & cvHexToBin(CStr(Offset))

    ElseIf Left(Offset, 2) = "0i" Then

        ' Offset is an IP address
        
        Offset = Right(Offset, Len(Offset) - 2)
        
        If Left(Offset, 1) = "-" Then
            parseOffset = "-"
            Offset = Right(Offset, Len(Offset) - 1)
        End If
        
        parseOffset = parseOffset & ParseIP(Offset).Address

    Else
        Offset = CLng(Offset)
        ' Offset is decimal
        If Offset < 0 Then
            parseOffset = "-"
            Offset = -1 * Offset
        End If
        
        parseOffset = parseOffset & cvDecToBin(CLng(Offset))

    End If
     
    Exit Function

'ErrorHandler:
'    If EnableDebug Then Debug.Print "Error " & Err.Number & " (" & Err.Description & ") in procedure parseOffset of Module IPFunctions"
'    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function


Private Function FormatOutput(BinaryIP As clsIP, DisplayFormat As Long) As Variant

    ' Convert result binary address into destination format

    Dim OutputFormat As tIPFormat
    
    OutputFormat = parseDisplayFormat(DisplayFormat)

'1 returns the IP address only, e.g. 192.168.0.1
'2 returns the IP address and decimal subnet/prefix length, e.g. 192.168.0.1/24
'3 returns the IP address and subnet mask, e.g. 192.168.0.1 255.255.255.0 (ipv4 only)
'4 returns the IP address and wildcard mask, e.g. 192.168.0.1 0.0.0.255 (ipv4 only)
'5 returns the significant octet(s) of the IP address only, e.g. .x if subnet mask <= 24, .x.y if subnet mask <= 16 (IPv4 only)
'6 returns the IP address in binary, e.g. 00001010100001110000001111111100

    If OutputFormat.Format = 6 Then

        FormatOutput = BinaryIP.Address
        Exit Function
    End If

    If (OutputFormat.Format = 3 Or OutputFormat.Format = 4 Or OutputFormat.Format = 5) And BinaryIP.IPv6 Then
        Err.Raise vbObjectError + 1030, "FormatOutput", "#Format not supported with IPv6 Addresses!"
    End If

    If BinaryIP.IPv6 Then
         FormatOutput = cvIP6Bin2Hex(BinaryIP.Address, OutputFormat.Padded, OutputFormat.QuadDotted)
    ElseIf OutputFormat.Format = 5 Then
'        FormatOutput = IP4SignificantOctets(cvIP4Bin2Dec(BinaryIP.Address, Padded, True), BinaryIP.MaskLength)
    Else
        FormatOutput = cvIP4Bin2Dec(BinaryIP.Address, OutputFormat.Padded, False)
    End If
    
    If OutputFormat.Format = 2 Then
        If OutputFormat.Padded = True And BinaryIP.IPv6 Then
            FormatOutput = FormatOutput & "/" & Format(BinaryIP.PrefixLength, "000")
        ElseIf OutputFormat.Padded = True Then
            FormatOutput = FormatOutput & "/" & Format(BinaryIP.PrefixLength, "00")
        Else
            FormatOutput = FormatOutput & "/" & BinaryIP.PrefixLength
        End If
    End If
    
    If OutputFormat.Format = 3 Then
        FormatOutput = FormatOutput & " " & cvIP4Bin2Dec(BinaryIP.Mask, OutputFormat.Padded)
    End If
    
    If OutputFormat.Format = 4 Then
        FormatOutput = FormatOutput & " " & cvIP4Bin2Dec(BinaryIP.WildcardMask, OutputFormat.Padded)
    End If
    
    Exit Function

End Function


Private Function IP6ShortenAddress(IP6 As Variant) As String

' New version RFC 5952 compliant (http://tools.ietf.org/html/rfc5952)

' Leading zeros in each 16-bit field are suppressed. For example, 2001:0db8::0001 is rendered
' as 2001:db8::1, though any all-zero field that is explicitly presented is rendered as 0.

' "::" is not used to shorten just a single 0 field. For example, 2001:db8:0:0:0:0:2:1 is
' shortened to 2001:db8::2:1, but 2001:db8:0000:1:1:1:1:1 is rendered as 2001:db8:0:1:1:1:1:1.

' Representations are shortened as much as possible. The longest sequence of consecutive all-zero
' fields is replaced by double-colon. If there are multiple longest runs of all-zero fields, then
' it is the leftmost that is compressed. E.g., 2001:db8:0:0:1:0:0:1 is rendered as 2001:db8::1:0:0:1
' rather than as 2001:db8:0:0:1::1.

' Hexadecimal digits are expressed as lower-case letters. For example, 2001:db8::1 is preferred over
' 2001:DB8::1.

    Dim cnvarr As Variant
    Dim i As Long, n As Long, M As Long

    On Error GoTo ErrorHandler

    cnvarr = Array(":0:0:0:0:0:0:", ":0:0:0:0:0:", ":0:0:0:0:", ":0:0:0:", ":0:0:")

    'Remove leading "0"s in each cell of the input array. Leave one "0" minimum.
    For n = LBound(IP6) To UBound(IP6)
        If InStr(1, IP6(n), ".") = 0 Then
            For M = 4 To 2 Step -1
                If Left(IP6(n), 1) = "0" Then
                    IP6(n) = Mid(IP6(n), 2, M)
                End If
            Next M
        End If
    Next n

    'Create string
    IP6ShortenAddress = Join(IP6, ":")

    'Find biggest group of ":0:"s and replace by ::
    For i = LBound(cnvarr) To UBound(cnvarr)
        If InStr(IP6ShortenAddress, cnvarr(i)) <> 0 Then
            IP6ShortenAddress = Replace(IP6ShortenAddress, cnvarr(i), "::", 1, 1, vbTextCompare)
            i = UBound(cnvarr)
        End If
    Next i


    'Remove first and last "0"s if followed/preced by double colons (i.e. 0::ef/64 -> ::ef/64, ef::0/64 -> ef::/64)
    If Left(IP6ShortenAddress, 3) = "0::" Then IP6ShortenAddress = Mid(IP6ShortenAddress, 2, Len(IP6ShortenAddress))
    If Right(IP6ShortenAddress, 3) = "::0" Then IP6ShortenAddress = Left(IP6ShortenAddress, Len(IP6ShortenAddress) - 1)

    Exit Function

ErrorHandler:
    If EnableDebug Then Debug.Print "Error " & Err.Number & " (" & Err.Description & ") in procedure IP6ShortenAddress of Module IPFunctions"

End Function


Private Function fnIPSort(InputArray() As clsIP)

    Dim i As Long
    Dim OutputSubnets() As String
    
    ' to quickly sort, we create an array that contains all subnets in the form (6 or 4) + ip address in binary + invert subnet mask in binary
    ' using this they can be sorted as strings
    
    ReDim OutputSubnets(LBound(InputArray) To UBound(InputArray))
    
    ' create an arrray with bin address + bin invert mask
    ' We use invert mask so the list can be sorted in one go
    ' we will invert subnet mask again before returning
    
    For i = LBound(InputArray) To UBound(InputArray)
        OutputSubnets(i) = InputArray(i).SortKey
    Next i
    
    ' sort the array in alphabetical order
    Call QuickSort(OutputSubnets, LBound(OutputSubnets), UBound(OutputSubnets))
    
    ' re-create the sorted array
    
    For i = LBound(InputArray) To UBound(InputArray)
        InputArray(i).SortKey = OutputSubnets(i)
    Next i
    
    Erase OutputSubnets
    
End Function





''''''''''''''''''''''''''''''
' Conversion Functions
''''''''''''''''''''''''''''''

Private Function cvIP4toBIN(ByVal strIPAddress As String) As String

    ' Checks that this is an IPv4 Address and converts to BIN

    On Error GoTo ErrorHandler

    Dim varAddress As Variant, n As Long, lCount As Long, BINSTRING As String
    
    varAddress = Split(strIPAddress, ".", , vbTextCompare)
    
    If isArray(varAddress) Then
    
        For n = LBound(varAddress) To UBound(varAddress)
            lCount = lCount + 1
            varAddress(n) = CByte(varAddress(n))
            BINSTRING = BINSTRING & WorksheetFunction.Dec2Bin(varAddress(n), 8)
        Next
        
        If (lCount <> 4) Then Err.Raise vbObjectError + 1010, "cvIP4toBIN", "#Invalid IP Address!"
        
    End If
    
    cvIP4toBIN = BINSTRING

    Exit Function

ErrorHandler:
    If EnableDebug Then Debug.Print "Error " & Err.Number & " (" & Err.Description & ") in procedure cvIP4toBIN of Module IPFunctions"

End Function


Private Function cvIP6toBIN(ByVal strIPAddress As String) As String

    ' Takes a short or full IPv6 address string input and returns an expanded one. Also Checks the address syntax.

    Dim varAddress As Variant
    Dim varAddressLeft As Variant
    Dim varAddressRight As Variant
    Dim arrIP6() As String
    Dim lCount As Long

    On Error GoTo ErrorHandler

    'If we find a dot in the address then this is quad-dotted and we will only have 7 groups
    
    If InStr(1, strIPAddress, ".", vbTextCompare) > 0 Then
        ReDim arrIP6(0 To 6)
    Else
        ReDim arrIP6(0 To 7)
    End If

    varAddress = Split(strIPAddress, "::", , vbTextCompare)

    If UBound(varAddress) = 0 Then
        'Fill the arrIP6 array with all the hex groups
        
        For lCount = 0 To UBound(varAddressLeft)
            arrIP6(lCount) = varAddress(lCount)
        Next lCount
        
    ElseIf UBound(varAddress) = 1 Then
    
        varAddressLeft = Split(varAddress(0), ":", , vbTextCompare)
        varAddressRight = Split(varAddress(1), ":", , vbTextCompare)
        
        If UBound(varAddressLeft) + UBound(varAddressRight) > UBound(arrIP6) Then
            ' Raise error
            Exit Function
        End If
        
        'Fill the arrIP6 array from the left with all the hex groups left of the double colons
        For lCount = 0 To UBound(varAddressLeft)
            arrIP6(lCount) = varAddressLeft(lCount)
        Next lCount
        
        'Fill the arrIP6 array from the right with all the hex groups right of the double colons
        For lCount = 0 To UBound(varAddressRight)
            arrIP6(UBound(arrIP6) - UBound(varAddressRight) + lCount) = varAddressRight(lCount)
        Next lCount
    Else
        ' We have multiple :: > error
        ' Raise error
        Exit Function
    End If

    'If some groups are empty or short (less than 4 characters) then pad with 0s
    For lCount = LBound(arrIP6) To UBound(arrIP6)
        If Len(arrIP6(lCount)) < 4 Then
            arrIP6(lCount) = String(4 - Len(arrIP6(lCount)), "0") & arrIP6(lCount)
        End If
    Next lCount

    'Convert to BIN. Last hex group can contain an IPv4 Address so ignored for now
    For lCount = LBound(arrIP6) To UBound(arrIP6) - 1
        cvIP6toBIN = cvIP6toBIN & cvHexToBin(arrIP6(lCount), 16)
    Next lCount

    If UBound(arrIP6) = 6 Then
        ' Last block is an IPv4 address
        cvIP6toBIN = cvIP6toBIN & cvIP4toBIN(arrIP6(UBound(arrIP6)))
    Else
        ' Last block is hex, we convert it
        cvIP6toBIN = cvIP6toBIN & cvHexToBin(arrIP6(UBound(arrIP6)), 16)
    End If

    Exit Function

ErrorHandler:
    If EnableDebug Then Debug.Print "Error " & Err.Number & " (" & Err.Description & ") in procedure cvIP6toBIN of Module IPFunctions"

End Function


Private Function cvIP4Bin2Dec(IPInput As String, Optional Padded As Boolean = False, Optional SignificantBytes As Boolean = False) As String

    Dim DecimalIPAddress As String
    Dim Dot As String
    Dim i As Long, j As Long

    On Error GoTo ErrorHandler

    DecimalIPAddress = ""
    Dot = ""

    For i = 0 To 3
        j = cvBinToDec(Mid(IPInput, i * 8 + 1, 8))

        If Padded = True Then
            DecimalIPAddress = DecimalIPAddress & Dot & Format(j, "000")
        Else
            DecimalIPAddress = DecimalIPAddress & Dot & j
        End If
        Dot = "."
    Next i

    cvIP4Bin2Dec = DecimalIPAddress

    Exit Function

ErrorHandler:
    If EnableDebug Then Debug.Print "Error " & Err.Number & " (" & Err.Description & ") in procedure cvIP4Bin2Dec of Module IPFunctions"

End Function


Private Function cvIP6Bin2Hex(Address As String, Optional Padded As Boolean = False, Optional QuadDotted As Boolean = False) As String

    Dim IParrRes() As String
    Dim i As Long

    If QuadDotted Then
        ReDim IParrRes(1 To 7)
    Else
        ReDim IParrRes(1 To 8)
    End If


    If QuadDotted Then
        For i = 0 To 5
            IParrRes(i + 1) = LCase(cvBinToHex(Mid(Address, (i * 16 + 1), 16)))
        Next i
        IParrRes(7) = cvIP4Bin2Dec(Mid(Address, (i * 16 + 1), 32), Padded)
    Else
        For i = 0 To 7
            IParrRes(i + 1) = LCase(cvBinToHex(Mid(Address, (i * 16 + 1), 16)))
        Next i
    End If

    If Padded = True Then
        cvIP6Bin2Hex = Join(IParrRes, ":")
    Else
        cvIP6Bin2Hex = IP6ShortenAddress(IParrRes)
    End If

    Exit Function

End Function


Private Function cvHexToBin(HexNum As String, Optional bits As Long) As String

' From http://www.ozgrid.com/forum/showthread.php?t=55298
' Modified to support longer numbers and padding

    Dim BinNum As String
    Dim lHexNum As Double
    Dim i As Double

    On Error GoTo ErrorHandler

    i = 0
    lHexNum = CDbl("&h" & HexNum)
    Do
        If lHexNum And 2 ^ i Then
            BinNum = "1" & BinNum
        Else
            BinNum = "0" & BinNum
        End If
        i = i + 1
    Loop Until (2 ^ i > lHexNum) And i >= bits

    cvHexToBin = BinNum

    Exit Function

ErrorHandler:
    If EnableDebug Then Debug.Print "Error " & Err.Number & " (" & Err.Description & ") in procedure cvHexToBin of Module IPFunctions"

End Function


Function cvBinToHex(bstr)
    
    'convert binary string to hex string

    Dim cnvarr As Variant
    Dim a As Long
    Dim hstr As String
    Dim dgt As String
    Dim ndgt As String
    Dim i As Long
    Dim ix As Long
    Dim k As Long

    cnvarr = Array("0000", "0001", "0010", "0011", _
                   "0100", "0101", "0110", "0111", "1000", _
                   "1001", "1010", "1011", "1100", "1101", _
                   "1110", "1111")
    
    'find number of HEX digits
    a = Len(bstr)
    ndgt = a / 4
    If (a Mod 4 > 0) Then
        MsgBox ("must be Long multiple of 4Bits")
        Exit Function
    End If
    hstr = ""
    For i = 1 To ndgt
        dgt = Mid(bstr, (i * 4) - 3, 4)
        For k = 0 To 15
            If (dgt = cnvarr(k)) Then
                ix = k
            End If
        Next
        hstr = hstr & Hex(ix)
    Next
    cvBinToHex = hstr

    Exit Function

End Function


Function cvBinToDec(D As String) As String

    ' From http://www.ozgrid.com/forum/showthread.php?t=15064

    Dim n As Double
    Dim Res As Double

    On Error GoTo ErrorHandler

    For n = Len(D) To 1 Step -1
        Res = Res + ((2 ^ (Len(D) - n)) * CLng(Mid(D, n, 1)))
    Next n
    cvBinToDec = Str(Res)

    Exit Function

ErrorHandler:
    If EnableDebug Then Debug.Print "Error " & Err.Number & " (" & Err.Description & ") in procedure cvBinToDec of Module IPFunctions"

End Function


Function cvDecToBin(ByVal SubNumber As Long) As String
    
    ' Convert a decimal number in a string into a binary number
    ' Limited to Decimal data type (79,228,162,514,264,337,593,543,950,335)
    
    Dim Negative As Boolean
    
    cvDecToBin = ""
  
    Do While SubNumber <> 0
        cvDecToBin = Trim$(Str$(SubNumber - 2 * Int(SubNumber / 2))) & cvDecToBin
        SubNumber = Int(SubNumber / 2)
    Loop
    
    Exit Function

End Function





''''''''''''''''''''''''''''''
' Conversion Functions
''''''''''''''''''''''''''''''

Public Sub QuickSort(vArray As Variant, inLow As Long, inHi As Long)
    ' From Robert Nunemaker at:
    ' http://en.allexperts.com/q/Visual-Basic-1048/string-manipulation.htm
    
  Dim pivot   As Variant
  Dim tmpSwap As Variant
  Dim tmpLow  As Long
  Dim tmpHi   As Long
   
  tmpLow = inLow
  tmpHi = inHi
   
  pivot = vArray((inLow + inHi) \ 2)
 
  While (tmpLow <= tmpHi)
 
     While (vArray(tmpLow) < pivot And tmpLow < inHi)
        tmpLow = tmpLow + 1
     Wend
     
     While (pivot < vArray(tmpHi) And tmpHi > inLow)
        tmpHi = tmpHi - 1
     Wend

     If (tmpLow <= tmpHi) Then
        tmpSwap = vArray(tmpLow)
        vArray(tmpLow) = vArray(tmpHi)
        vArray(tmpHi) = tmpSwap
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
     End If
  
  Wend
 
  If (inLow < tmpHi) Then QuickSort vArray, inLow, tmpHi
  If (tmpLow < inHi) Then QuickSort vArray, tmpLow, inHi
 
End Sub

