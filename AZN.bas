Attribute VB_Name = "AZN"

Option Explicit
'Main Function
Function SpellNumberAzn(ByVal MyNumber)
    Dim Manat, Qepik, Temp
    Dim DecimalPlace, Count
    ReDim Place(9) As String
    Place(2) = " Min "
    Place(3) = " Milyon "
    Place(4) = " Milyard "
    Place(5) = " Trilyon "
 
    MyNumber = Trim(Str(MyNumber))
    DecimalPlace = InStr(MyNumber, ".")
    If DecimalPlace > 0 Then
        Qepik = GetTens(Left(Mid(MyNumber, DecimalPlace + 1) & _
                  "00", 2))
        MyNumber = Trim(Left(MyNumber, DecimalPlace - 1))
    End If
    Count = 1
    Do While MyNumber <> ""
        Temp = GetYuzs(Right(MyNumber, 3))
        If Temp <> "" Then Manat = Temp & Place(Count) & Manat
        If Len(MyNumber) > 3 Then
            MyNumber = Left(MyNumber, Len(MyNumber) - 3)
        Else
            MyNumber = ""
        End If
        Count = Count + 1
    Loop
    Select Case Manat
        Case ""
            Manat = " "
        Case "Bir"
            Manat = "Bir Manat"
         Case Else
            Manat = Manat & " Manat"
    End Select
    Dim qapik As String
    qapik = ThisWorkbook.Sheets("musteri").Range("A13").Value
    Select Case Qepik
        Case ""
            Qepik = " "
        Case "Bir"
            Qepik = " Bir " & qapik
              Case Else
            Qepik = " " & Qepik & " " & qapik
    End Select
    SpellNumberAzn = Manat & Qepik
End Function
 
Function GetYuzs(ByVal MyNumber)
    Dim Result As String
    Dim hundred As String
    hundred = ThisWorkbook.Sheets("musteri").Range("A14").Value
    If Val(MyNumber) = 0 Then Exit Function
    MyNumber = Right("000" & MyNumber, 3)
    ' Convert the Yuzs place.
    If Mid(MyNumber, 1, 1) <> "0" Then
        Result = GetDigit(Mid(MyNumber, 1, 1)) & " " & hundred & " "
    End If
    ' Convert the tens and Birs place.
    If Mid(MyNumber, 2, 1) <> "0" Then
        Result = Result & GetTens(Mid(MyNumber, 2))
    Else
        Result = Result & GetDigit(Mid(MyNumber, 3))
    End If
    GetYuzs = Result
End Function


Function GetTens(TensText)
        Dim uch As String
        Dim dord As String
        Dim besh As String
        Dim alti As String
        Dim sekkiz As String
        Dim qirx As String
        Dim elli As String
        Dim altmish As String
        Dim yetmish As String
        Dim seksan As String
        uch = ThisWorkbook.Sheets("musteri").Range("A3").Value
        dord = ThisWorkbook.Sheets("musteri").Range("A4").Value
        besh = ThisWorkbook.Sheets("musteri").Range("A5").Value
        alti = ThisWorkbook.Sheets("musteri").Range("A6").Value
        sekkiz = ThisWorkbook.Sheets("musteri").Range("A7").Value
        
        qirx = ThisWorkbook.Sheets("musteri").Range("A8").Value
        elli = ThisWorkbook.Sheets("musteri").Range("A9").Value
        altmish = ThisWorkbook.Sheets("musteri").Range("A10").Value
        yetmish = ThisWorkbook.Sheets("musteri").Range("A11").Value
        seksan = ThisWorkbook.Sheets("musteri").Range("A12").Value
        
    Dim Result As String
    Result = "" ' Null out the temporary function value.
    If Val(Left(TensText, 1)) = 1 Then   ' If value between 10-19...
        Select Case Val(TensText)
            Case 10: Result = "On"
            Case 11: Result = "Onbir"
            Case 12: Result = "Oniki"
            Case 13: Result = "On" & uch
            Case 14: Result = "On" & dord
            Case 15: Result = "On" & besh
            Case 16: Result = "On" & alti
            Case 17: Result = "Onyeddi"
            Case 18: Result = "On" & sekkiz
            Case 19: Result = "Ondoqquz"
            Case Else
        End Select
    Else ' If value between 20-99...
        Select Case Val(Left(TensText, 1))
            Case 2: Result = "Iyirmi "
            Case 3: Result = "Otuz "
            Case 4: Result = qirx & " "
            Case 5: Result = elli & " "
            Case 6: Result = altmish & " "
            Case 7: Result = yetmish & " "
            Case 8: Result = seksan & " "
            Case 9: Result = "Doqsan "
            Case Else
        End Select
        Result = Result & GetDigit _
            (Right(TensText, 1))  ' Retrieve Birs place.
    End If
    GetTens = Result
End Function
 
Function GetDigit(Digit)
    Dim uch As String
    Dim dord As String
    Dim besh As String
    Dim alti As String
    Dim sekkiz As String
    Dim qirx As String
    Dim elli As String
    Dim altmish As String
    Dim yetmish As String
    Dim seksan As String
    Dim iki As String
    
    iki = ThisWorkbook.Sheets("musteri").Range("A2").Value
    uch = ThisWorkbook.Sheets("musteri").Range("A3").Value
    dord = ThisWorkbook.Sheets("musteri").Range("A4").Value
    besh = ThisWorkbook.Sheets("musteri").Range("A5").Value
    alti = ThisWorkbook.Sheets("musteri").Range("A6").Value
    sekkiz = ThisWorkbook.Sheets("musteri").Range("A7").Value
    
    qirx = ThisWorkbook.Sheets("musteri").Range("A8").Value
    elli = ThisWorkbook.Sheets("musteri").Range("A9").Value
    altmish = ThisWorkbook.Sheets("musteri").Range("A10").Value
    yetmish = ThisWorkbook.Sheets("musteri").Range("A11").Value
    seksan = ThisWorkbook.Sheets("musteri").Range("A12").Value
        
    Select Case Val(Digit)
        Case 1: GetDigit = "Bir"
        Case 2: GetDigit = iki
        Case 3: GetDigit = uch
        Case 4: GetDigit = dord
        Case 5: GetDigit = besh
        Case 6: GetDigit = alti
        Case 7: GetDigit = "Yeddi"
        Case 8: GetDigit = sekkiz
        Case 9: GetDigit = "Doqquz"
        Case Else: GetDigit = ""
    End Select
End Function







