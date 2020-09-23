Attribute VB_Name = "modNumberToWords"
Option Explicit

Function NumToWords(strNumberString As String) As String

    Dim intGroupX           As Integer
    Dim intUnitDigit        As Integer
    Dim intTensDigit        As Integer
    Dim intHundDigit        As Integer
    Dim intGroupStartPos    As Integer
    Dim intNbrLeadingZeros  As Integer
    Dim intX                As Integer
    Dim strWorkNumber       As String
    Dim strGroupVerbage     As String
    Dim arrUnitName         As Variant
    Dim arrCompoundName     As Variant
    Dim arrTensName         As Variant
    Dim arrGroupName        As Variant
    
    arrUnitName = Array("", "One", "Two", "Three", "Four", "Five", _
                        "Six", "Seven", "Eight", "Nine")
    arrCompoundName = Array("", "Eleven", "Twelve", "Thirteen", _
                            "Fourteen", "Fifteen", "Sixteen", _
                            "Seventeen", "Eighteen", "Nineteen")
    arrTensName = Array("", "Ten", "Twenty", "Thirty", "Forty", _
                        "Fifty", "Sixty", "Seventy", "Eighty", _
                        "Ninety")
    arrGroupName = Array("", "Thousand", "Million", "Billion", _
                         "Trillion", "Quadrillion", "Quintillion", _
                         "Sextillion", "Septillion", "Octillion", _
                         "Nintillion", "Decillion")
                         
    If Not IsAllDigits(strNumberString) Then
        MsgBox "Numeric argument required.", vbInformation, _
               "Invalid Argument"
        NumToWords = ""
        Exit Function
    End If
    
    If Len(strNumberString) > 36 Then
        MsgBox "Argument must not exceed 36 digits.", vbInformation, _
               "Argument Too Big"
        NumToWords = ""
        Exit Function
    End If
    
    intNbrLeadingZeros = 36 - Len(strNumberString)
    strWorkNumber = strNumberString
    For intX = 1 To intNbrLeadingZeros
        strWorkNumber = "0" & strWorkNumber
    Next
    
    intGroupStartPos = 34
    For intGroupX = 0 To 11
        intUnitDigit = Val(Mid$(strWorkNumber, intGroupStartPos + 2, 1))
        intTensDigit = Val(Mid$(strWorkNumber, intGroupStartPos + 1, 1))
        intHundDigit = Val(Mid$(strWorkNumber, intGroupStartPos, 1))
        If intUnitDigit = 0 And intTensDigit = 0 Then
            strGroupVerbage = ""
        ElseIf intUnitDigit = 0 Then
            strGroupVerbage = arrTensName(intTensDigit)
        ElseIf intTensDigit = 1 Then
            strGroupVerbage = arrCompoundName(intUnitDigit)
        ElseIf intTensDigit = 0 Then
            strGroupVerbage = arrUnitName(intUnitDigit)
        Else
            strGroupVerbage = arrTensName(intTensDigit) & "-" _
                            & arrUnitName(intUnitDigit)
        End If
        If intHundDigit <> 0 Then
            strGroupVerbage = arrUnitName(intHundDigit) & " Hundred " _
                            & strGroupVerbage
        End If
        If strGroupVerbage <> "" Then
            strGroupVerbage = strGroupVerbage & " " _
                            & arrGroupName(intGroupX)
        End If
        NumToWords = strGroupVerbage & " " & NumToWords
        intGroupStartPos = intGroupStartPos - 3
    Next
    
    NumToWords = Trim$(NumToWords)
    If NumToWords = "" Then NumToWords = "ZERO"
    
End Function

Function IsAllDigits(strTestString As String) As Boolean

    Dim intX    As Integer
    
    For intX = 1 To Len(strTestString)
        If InStr("0123456789", Mid$(strTestString, intX, 1)) > 0 Then
            ' continue
        Else
            IsAllDigits = False
            Exit Function
        End If
    Next
    
    IsAllDigits = True
    
End Function


