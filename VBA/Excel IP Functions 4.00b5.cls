VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' IP Functions v4.00
' © 2013-2017 Nicolas Grison

'
' IP Class Module for IP Functions
'





'##############################################
'#
'# Member Variables
'#
'##############################################



Private dAddress As String
Private dMask As String





'##############################################
'#
'# Properties
'#
'##############################################



''''''''''''''''''''''''''''''
' Address
''''''''''''''''''''''''''''''

Property Get Address() As String
    
    Address = dAddress

End Property

Property Let Address(dValue As String)

    dAddress = dValue
    dIPLength = Len(dValue)

End Property


''''''''''''''''''''''''''''''
' Mask
''''''''''''''''''''''''''''''

Property Get Mask() As String

    Mask = dMask

End Property

Property Let Mask(dValue As String)
    
    dMask = dValue

End Property



''''''''''''''''''''''''''''''
' Prefix Length
''''''''''''''''''''''''''''''

' This property doesn't exist as such, it is just another way to read/write Mask

Property Get PrefixLength() As Long
    
    Dim i As Long
    Dim MaskLength As Long
    
    i = 1
    MaskLength = Len(dMask)
    
    Do While Mid(Mask, i, 1) = "1" And i <= MaskLength
        i = i + 1
    Loop
    
    PrefixLength = i - 1
    
End Property

Property Let PrefixLength(dValue As Long)
    
    If dValue >= 0 And dValue <= IPLength Then
        dMask = String(dValue, "1") & String(Me.IPLength - dValue, "0")
    Else
        Err.Raise vbObjectError + 1010, "clsIP.PrefixLength", "#Incorrect Prefix Length"
    End If

End Property


''''''''''''''''''''''''''''''
' Wildcard Mask
''''''''''''''''''''''''''''''

Property Get WildcardMask() As String

    WildcardMask = bInvert(dMask)

End Property


''''''''''''''''''''''''''''''
' IP Length
''''''''''''''''''''''''''''''

Property Get IPLength() As Long
    IPLength = Len(dAddress)
End Property


''''''''''''''''''''''''''''''
' IPv6
''''''''''''''''''''''''''''''

Property Get IPv6() As Boolean

    If Me.IPLength = 128 Then
        IPv6 = True
    Else
        IPv6 = False
    End If
    
End Property


''''''''''''''''''''''''''''''
' Sorting Key
''''''''''''''''''''''''''''''

Property Get SortKey() As String
    
    SortKey = SortKey & Me.Address & Me.WildcardMask
    
End Property

Property Let SortKey(dValue As String)
    
    dAddress = Left(dValue, Len(dValue) / 2)
    dMask = bInvert(Right(dValue, Len(dValue) / 2))
    
End Property


''''''''''''''''''''''''''''''
' Subnet Address
''''''''''''''''''''''''''''''

Property Get SubAddress() As clsIP

' Takes a binary IP address and subnet mask and returns the subnet address in binary
' (Bitwise AND Operator on the Binary Network IPAddress and the CIDR)
        
    Dim i As Long
    Dim builtAddress As String

    builtAddress = ""

    For i = 1 To Len(dAddress)

        If Mid(dAddress, i, 1) And Mid(Me.Mask, i, 1) Then
            builtAddress = builtAddress & "1"
        Else
            builtAddress = builtAddress & "0"
        End If

    Next i
    
    Set SubAddress = New clsIP
    SubAddress.Address = builtAddress
    SubAddress.Mask = Me.Mask

End Property


''''''''''''''''''''''''''''''
' First Host Address
''''''''''''''''''''''''''''''

Property Get FirstHost() As clsIP

    ' Returns First Host IPAddress. Subnet address always ends in 0 so we just switch the last bit
    
    Dim builtAddress As String
    
    If Me.PrefixLength >= Me.IPLength - 1 Then
        ' return error
        Exit Property
    Else
        builtAddress = bAdd(Me.SubAddress().Address, "1")
    End If
    
    Set FirstHost = New clsIP
    FirstHost.Address = builtAddress
    FirstHost.Mask = Me.Mask

End Property


''''''''''''''''''''''''''''''
' Previous Host Address
''''''''''''''''''''''''''''''

Property Get PreviousHost() As clsIP

    'Return Previous Host Address
    
    Dim builtAddress As String
    
    If Me.Address <= Me.FirstHost().Address Then
        ' return error
        Err.Raise vbObjectError + 1030, "clsIP.PreviousHost", "Out of address space"
        Exit Property
    End If
    
    builtAddress = bSub(dAddress, "1")
    
    Set PreviousHost = New clsIP
    PreviousHost.Address = builtAddress
    PreviousHost.Mask = Me.Mask
    
End Property


''''''''''''''''''''''''''''''
' Next Host Address
''''''''''''''''''''''''''''''

Property Get NextHost() As clsIP

'Return Next Host Address

' If no subnet mask then check the address isn't all 1s to begin with
    
    Dim builtAddress As String
        
    If Me.Address >= Me.LastHost().Address Then
        ' return error
        Exit Property
    End If
    
    builtAddress = bAdd(dAddress, "1")

    Set NextHost = New clsIP
    NextHost.Address = builtAddress
    NextHost.Mask = Me.Mask
    
End Property


''''''''''''''''''''''''''''''
' Last Host Address
''''''''''''''''''''''''''''''

Property Get LastHost() As clsIP

' Return Last Host IPAddress. In IPv6, if subnet > 127 hosts then last 127 addresses reserved.
    
    Dim builtAddress As String

    If Me.PrefixLength >= Me.IPLength - 1 Then
        Exit Property
        
    ElseIf Not Me.IPv6 Then
        ' We have an IPv4 address, last host address is broadcast - 1
        builtAddress = bSub(Me.BroadcastAddress().Address, "1")
        
    ElseIf Me.PrefixLength > 120 Then
        ' We have a small IPv6 subnet, last address is last host
        builtAddress = Me.LastAddress().Address
        
    Else
        ' We have an IPv6 subnet > 120, last 127 addresses are reserved anycast
        builtAddress = Left(Me.LastAddress().Address, Me.IPLength - 8) & "01111111"
        
    End If

    Set LastHost = New clsIP
    LastHost.Address = builtAddress
    LastHost.Mask = Me.Mask

End Property


''''''''''''''''''''''''''''''
' Broadcast Address
''''''''''''''''''''''''''''''

Property Get BroadcastAddress() As clsIP

' Return the broadcast address with an IPv4 address. Returns an error with IPv6
' (Bitwise OR Operator on the Binary Network IPAddress and the inverted CIDR)
    
    Set BroadcastAddress = New clsIP
        
    If Me.IPv6 Then
        Err.Raise vbObjectError + 1020, "clsIP.BroadcastAddress", "#No broadcast address in IPv6!"
    ElseIf Me.PrefixLength >= 31 Then
        Err.Raise vbObjectError + 1020, "clsIP.BroadcastAddress", "#No broadcast address if mask >= 31!"
    End If
        
    BroadcastAddress.Address = Me.LastAddress.Address
    BroadcastAddress.Mask = Me.Mask

End Property


''''''''''''''''''''''''''''''
' Last Address
''''''''''''''''''''''''''''''

Property Get LastAddress() As clsIP

' Return the very last address in the subnet.
' For IPv4 this is the broadcast address, for IPv6 the last host address
' (Bitwise OR Operator on the Binary Network IPAddress and the inverted CIDR)

    Dim builtAddress As String
    Dim InvertedMask As String
    Dim i As Long, j As Long
        
    ' Start by inverting the CIDR
    InvertedMask = bInvert(Me.Mask)

    ' Do an OR between Address and inverted CIDR

    builtAddress = ""

    For i = 1 To Me.IPLength
        If Mid(dAddress, i, 1) Or Mid(InvertedMask, i, 1) Then
            builtAddress = builtAddress & "1"
        Else
            builtAddress = builtAddress & "0"
        End If
    Next i
    
    Set LastAddress = New clsIP
    LastAddress.Address = builtAddress
    LastAddress.Mask = Me.Mask
    
End Property


''''''''''''''''''''''''''''''
' Previous Subnet Address
''''''''''''''''''''''''''''''

Property Get SubPrev() As clsIP

'  Return Previous Subnet
    Dim builtAddress As String
    
    builtAddress = Me.SubAddress().Address

    ' Get the current subnet address
    ' Start by checking that this isn't already the first subnet available
    If fnStringIsAllX(builtAddress, "0") Then Err.Raise vbObjectError + 1030, "clsIP.PreviousSubnet", "#Subnet out of range!"

    ' Then calculate the previous subnet by subtracting the inverted mask from the subnet address and substract 1
    builtAddress = bSub(builtAddress, bInvert(Me.Mask))
    builtAddress = bSub(builtAddress, "1")

    Set SubPrev = New clsIP
    SubPrev.Address = builtAddress
    SubPrev.Mask = Me.Mask
    
End Property


''''''''''''''''''''''''''''''
' Next Subnet Address
''''''''''''''''''''''''''''''

Property Get SubNext() As clsIP

' Return Next Subnet. Get the broadcast address and add 1.
    
    Dim builtAddress As String

    ' Get the last address in the subnet
    builtAddress = Me.LastAddress().Address

    ' Start by checking that this isn't already the last subnet available
    If fnStringIsAllX(builtAddress, "1") Then Err.Raise vbObjectError + 1030, "clsIP.NextSubnet", "#Subnet out of range!"

    ' Then calculate the next subnet by adding 1 to the broadcast address
    builtAddress = bAdd(builtAddress, "1")
    
    Set SubNext = New clsIP
    SubNext.Address = builtAddress
    SubNext.Mask = Me.Mask
    
End Property


''''''''''''''''''''''''''''''
' Subnet Host Count
''''''''''''''''''''''''''''''

Property Get HostCount() As String

    ' Returns the number of host addresses in the subnet

    HostCount = bSub(Me.LastHost().Address, Me.FirstHost().Address)
    HostCount = bAdd(HostCount, "1")
    
    HostCount = HostCount

End Property




'##############################################
'#
'# Events
'#
'##############################################

Private Sub Class_Initialize()
    ' Class Constructor
    
End Sub


Private Sub Class_Terminate()
   ' Class Destructor
   
End Sub





'##############################################
'#
'# IP Methods
'#
'##############################################



''''''''''''''''''''''''''''''
' Host X
''''''''''''''''''''''''''''''

Public Function HostX(astr As String) As clsIP

    ' Returns host address X from the beginning of the subnet
    Dim builtAddress As String
    
    ' Check that Host Number is positive
    If Left(astr, 1) = "-" Then Err.Raise vbObjectError + 1010, "clsIP.fnHostX", "#Negative Host Number!"
    
    ' Check that it isn't zero
    If fnStringIsAllX(astr, "0") Then Err.Raise vbObjectError + 1010, "clsIP.fnHostX", "#Host Number = 0!"
    
    ' Remove one (address 1 is Me.LastHost.Address)
    astr = bSub(astr, "1")
    
    'Calculate
'    builtAddress = bAdd(Me.SubAddress.Address, astr)
    builtAddress = bAdd(Me.FirstHost.Address, astr)
    
    ' Check that we haven't exceeded the last host
    If builtAddress > Me.LastHost().Address Then Err.Raise vbObjectError + 1030, "clsIP.fnHostX", "#Host Number larger than subnet!"
    
    Set HostX = New clsIP
    HostX.Address = builtAddress
    HostX.Mask = Me.Mask
    
End Function


''''''''''''''''''''''''''''''
' Host Y
''''''''''''''''''''''''''''''

Public Function HostY(astr As String) As clsIP

    Dim builtAddress As String
    
    ' Check that Host Number is positive
    If Left(HostNumber, 1) = "-" Then Err.Raise vbObjectError + 1010, "clsIP.fnHostX", "#Negative Host Number!"
    
    ' Check that it isn't zero
    If fnStringIsAllX(astr, "0") Then Err.Raise vbObjectError + 1010, "clsIP.fnHostX", "#Host Number = 0!"
    
    ' Remove one (address 1 is Me.LastHost.Address)
    astr = bSub(astr, "1")
    
    'Calculate
    builtAddress = bSub(Me.LastHost.Address, astr)
    
    ' Check that we haven't exceeded the first host
    If builtAddress < Me.FirstHost().Address Then Err.Raise vbObjectError + 1030, "clsIP.HostY", "#Host Number larger than subnet!"
    
    Set HostY = New clsIP
    HostY.Address = builtAddress
    HostY.Mask = Me.Mask
    
End Function


''''''''''''''''''''''''''''''
' Offset
''''''''''''''''''''''''''''''

Public Function Offset(BinaryOffset As String, Optional checkSubnet As Boolean = True) As clsIP

    ' Binary IP Offset
    ' We expect a binary string preceded by a - if negative

    Dim Negative As Boolean
    Dim builtAddress As String
    Dim test1 As String
    Dim test2 As String
    
    If BinaryOffset = vbNullString Then
        Set Offset = Me
        Exit Function
    End If
    
    Negative = False

    If Left(BinaryOffset, 1) = "-" Then
        Negative = True
        BinaryOffset = Right(BinaryOffset, Len(BinaryOffset) - 1)
    End If

    If Len(BinaryOffset) > Me.IPLength Then Err.Raise vbObjectError + 1010, "clsIP.Offset", "#Offset bigger than address space!"

    'Ensure that the offset is same length as the address
    While Len(BinaryOffset) < Me.IPLength
        BinaryOffset = "0" & BinaryOffset
    Wend
    
    If Negative = True Then
        builtAddress = bSub(Me.Address, BinaryOffset)
    Else
        builtAddress = bAdd(Me.Address, BinaryOffset)
    End If
    
    Set Offset = New clsIP
    Offset.Address = builtAddress
    Offset.Mask = Me.Mask
    
    If checkSubnet Then
        test1 = Me.FirstHost().Address
        test2 = Me.LastHost().Address
        If builtAddress < test1 Or builtAddress > test2 Then
            Err.Raise vbObjectError + 1030, "clsIP.IPOffset", "Out of address space"
        End If
    Else
        If Len(builtAddress) > Me.IPLength Or Left(builtAddress, 1) = "-" Then
            Err.Raise vbObjectError + 1030, "clsIP.IPOffset", "Out of address space"
        End If
    End If
    
    Exit Function
    
End Function


''''''''''''''''''''''''''''''
' Included in Summary
''''''''''''''''''''''''''''''

Public Function Summary(SummaryInput As clsIP) As clsIP

    ' Check that Me is included in the provided summary
    ' Used in conjunction with main subnet functions
    ' return the same object if included, an error if not
    
    Dim builtAddress As String

    If SummaryInput.Address > Me.SubAddress().Address Or SummaryInput.LastAddress().Address < Me.LastAddress().Address Then
        ' We are out of summary, return error
        Err.Raise vbObjectError + 1030, "clsIP.Summary", "Out of summary"
    End If
    
    Set Summary = Me
    
End Function


''''''''''''''''''''''''''''''
' First Summary Subnet
''''''''''''''''''''''''''''''

Public Function SumFirstSub(PrefixLength As Long) As clsIP

    Dim builtAddress As String
    Dim ProcessedIP As clsIP
    
    'Pre-checks
    If PrefixLength <= 0 Or PrefixLength > Me.IPLength Then Err.Raise vbObjectError + 1020, "clsIP.SumFirstSub", "#Incorrect Prefix Length!"
    
    If Me.PrefixLength > PrefixLength Then Err.Raise vbObjectError + 1030, "clsIP.SumFirstSub", "#Summary too small!"
    
    ' We build a subnet object with the address of Me and the prefix length provided
    Set ProcessedIP = New clsIP
    ProcessedIP.Address = Me.Address
    ProcessedIP.PrefixLength = PrefixLength
    
    Set SumFirstSub = New clsIP
    SumFirstSub.Address = ProcessedIP.SubAddress.Address
    SumFirstSub.PrefixLength = ProcessedIP.PrefixLength

End Function


''''''''''''''''''''''''''''''
' Summary Subnet X
''''''''''''''''''''''''''''''

Public Function SumSubX(PrefixLength As Long, astr As String) As clsIP

' Takes a summary, a smaller prefix and
' returns the subnet number SubNumber based on the smaller prefix in the summary

    Dim bstr As String
    Dim binOffset As String
    Dim ProcessedIP As clsIP
    
    ' Check that Subnet Number is positive
    If Left(astr, 1) = "-" Then Err.Raise vbObjectError + 1010, "clsIP.SumSubX", "#Negative Subnet Number!"
    
    ' Check that it isn't zero
    If fnStringIsAllX(astr, "0") Then Err.Raise vbObjectError + 1010, "clsIP.SumSubX", "#Subnet Number = 0!"
    
    ' Remove one (address 1 is Me.LastHost.Address)
    astr = bSub(astr, "1")

    ' We get the last subnet of the summary
    Set ProcessedIP = Me.SumFirstSub(PrefixLength)
    
    ' Number of the subnet we are looking for in binary
    SubNumber = bSub(astr, "1")
    
    bstr = bAdd(bInvert(ProcessedIP.Mask), "1")
    
    ' For each subnet we add the subnet mask + 1
    binOffset = bMult(astr, bstr)
    
    ProcessedIP.Address = bAdd(ProcessedIP.Address, binOffset)
    
    If ProcessedIP.LastAddress.Address > Me.LastAddress.Address Then Err.Raise vbObjectError + 1030, "clsIP.SumSubX", "#Out of summary!"
    
    Set SumSubX = ProcessedIP
    
End Function


''''''''''''''''''''''''''''''
' Summary Subnet Y
''''''''''''''''''''''''''''''

Public Function SumSubY(PrefixLength As Long, astr As String) As clsIP

' Takes a summary, a smaller prefix and
' returns the subnet number SubNumber based on the smaller prefix in the summary

    Dim bstr As String
    Dim binOffset As String
    Dim ProcessedIP As clsIP
    
    ' Check that Subnet Number is positive
    If Left(astr, 1) = "-" Then Err.Raise vbObjectError + 1010, "clsIP.SumSubY", "#Negative Subnet Number!"
    
    ' Check that it isn't zero
    If fnStringIsAllX(astr, "0") Then Err.Raise vbObjectError + 1010, "clsIP.SumSubY", "#Subnet Number = 0!"
    
    ' We get the last subnet of the summary
    Set ProcessedIP = Me.SumLastSub(PrefixLength)
        
    ' Remove one from subnet number (subnet 1 is Me.LastHost.Address)
    astr = bSub(astr, "1")
    
    bstr = bAdd(bInvert(ProcessedIP.Mask), "1")
    
    ' For each subnet we add the subnet mask + 1
    binOffset = bMult(astr, bstr)
    
    ProcessedIP.Address = bSub(ProcessedIP.Address, binOffset)
    
    If ProcessedIP.SubAddress.Address < Me.SubAddress.Address Then Err.Raise vbObjectError + 1030, "clsIP.SumSubY", "#Out of summary!"

    Set SumSubY = ProcessedIP
    
End Function


''''''''''''''''''''''''''''''
' Last Summary Subnet
''''''''''''''''''''''''''''''

Public Function SumLastSub(PrefixLength As Long) As clsIP

    Dim builtAddress As String
    Dim ProcessedIP As clsIP
    
    'Pre-checks
    If PrefixLength <= 0 Or PrefixLength > Me.IPLength Then Err.Raise vbObjectError + 1020, "clsIP.SumFirstSub", "#Incorrect Prefix Length!"
    
    If Me.PrefixLength > PrefixLength Then Err.Raise vbObjectError + 1030, "clsIP.SumFirstSub", "#Summary too small!"
    
    ' We build a subnet object with the address of Me and the prefix length provided
    Set ProcessedIP = New clsIP
    ProcessedIP.Address = Me.LastAddress.Address
    ProcessedIP.PrefixLength = PrefixLength
    
    Set SumLastSub = New clsIP
    SumLastSub.Address = ProcessedIP.SubAddress.Address
    SumLastSub.PrefixLength = ProcessedIP.PrefixLength

End Function


''''''''''''''''''''''''''''''
' Boolean Summary Check
''''''''''''''''''''''''''''''

Public Function SumCheck(BinarySummaryAddress As clsIP) As Boolean

' Checks whether IPInput is included within SummaryInput

    Dim IPInputFirstAddress As String
    Dim IPInputLastAddress As String
    Dim SummaryInputFirstAddress As String
    Dim SummaryInputLastAddress As String

    ' Compare and output
    If BinarySummaryAddress.SubAddress.Address <= Me.SubAddress.Address And BinarySummaryAddress.LastAddress.Address >= Me.LastAddress.Address Then
        SumCheck = True
    Else
        SumCheck = False
    End If

End Function





'##############################################
'#
'# Other Methods
'#
'##############################################



''''''''''''''''''''''''''''''
' Binary Operations
''''''''''''''''''''''''''''''

Private Function bAdd(bstr As String, addstr As String) As String

' Add two binary strings together
' The strings must be of the same length
' Returns an error if addition is bigger than string length

    Dim CarryOver As Long
    Dim n As Long
    Dim Sum As Long
    
    'Ensure that both strings are of the same length

    If Len(bstr) > Len(addstr) Then
        Do
            addstr = "0" & addstr
        Loop Until Len(addstr) = Len(bstr)
    ElseIf Len(bstr) < Len(addstr) Then
        Do
            bstr = "0" & bstr
        Loop Until Len(addstr) = Len(bstr)
    End If
    
    bAdd = ""
    CarryOver = 0

    For n = Len(bstr) To 1 Step -1
        Sum = CInt(Mid(bstr, n, 1)) + CInt(Mid(addstr, n, 1)) + CarryOver
        If Sum = 3 Then
            bAdd = "1" & bAdd
            CarryOver = 1
        ElseIf Sum = 2 Then
            bAdd = "0" & bAdd
            CarryOver = 1
        ElseIf Sum = 1 Then
            bAdd = "1" & bAdd
            CarryOver = 0
        Else
            bAdd = "0" & bAdd
            CarryOver = 0
        End If
    Next n
    
    If CarryOver = 1 Then bAdd = "1" & bAdd

    Exit Function

End Function



Private Function bSub(bstr As String, substr As String) As String

'Substract a binary strings from another

    Dim CarryOver As Long
    Dim n As Long
    Dim Sum As Long

    'Ensure that both strings are of the same length

    If Len(bstr) > Len(substr) Then
        Do
            substr = "0" & substr
        Loop Until Len(substr) = Len(bstr)
    ElseIf Len(bstr) < Len(substr) Then
        Do
            bstr = "0" & bstr
        Loop Until Len(substr) = Len(bstr)
    End If

    bSub = ""
    CarryOver = 0

    For n = Len(bstr) To 1 Step -1
        Sum = CInt(Mid(bstr, n, 1)) - CInt(Mid(substr, n, 1)) - CarryOver
        If Sum = 1 Then
            bSub = "1" & bSub
            CarryOver = 0
        ElseIf Sum = 0 Then
            bSub = "0" & bSub
            CarryOver = 0
        ElseIf Sum = -1 Then
            bSub = "1" & bSub
            CarryOver = 1
        ElseIf Sum = -2 Then
            bSub = "0" & bSub
            CarryOver = 1
        End If
    Next n

    If CarryOver = 1 Then bSub = "-" & bSub(substr, bstr)

    Exit Function

End Function


Private Function bMult(astr As String, bstr As String) As String

'Multiply two binary strings of any length together

    Dim CarryOver As Long
    Dim n As Long
    
    bMult = "0"
    
    For n = Len(bstr) To 1 Step -1
        
        If CInt(Mid(bstr, n, 1)) = "1" Then
            
            bMult = bAdd(bMult, astr)
        End If
        
        astr = astr & "0"
        
    Next n
        
    Exit Function

End Function


Private Function bInvert(Mask As String) As String

    Dim InvertedMask As String
    Dim i As Long

    InvertedMask = ""

    For i = 1 To Len(Mask)
        If Mid(Mask, i, 1) Then
            InvertedMask = InvertedMask & "0"
        Else
            InvertedMask = InvertedMask & "1"
        End If
    Next i

    bInvert = InvertedMask

End Function





''''''''''''''''''''''''''''''
' Others
''''''''''''''''''''''''''''''

Private Function IPBinMask2Length(Mask As String) As Long
    
    ' Determines PrefixLength from Mask
    
    Dim i As Long
    Dim MaskLength As Long
    
    i = 1
    MaskLength = Len(Mask)
    
    Do While Mid(Mask, i, 1) = "1" And i <= MaskLength
        i = i + 1
    Loop
    
    IPBinMask2Length = i - 1

End Function


Private Function fnStringIsAllX(IPInput As String, CharToCheck As String) As Boolean

' Takes a string as input and checks that it only contains the CharToCheck number

    If Len(Trim(Replace(IPInput, CharToCheck, vbNullString))) = 0 Then
        fnStringIsAllX = True
    Else
        fnStringIsAllX = False
    End If

End Function
