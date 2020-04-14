Attribute VB_Name = "Module3"
Global hundreds As Integer
Global remainder As Integer
Global result As Integer
Global tens As Integer
Global num As Integer
Global power_value(1 To 5) As Currency
Global power_name(1 To 5) As String
Global digits As Integer
Global i As Integer
Global Naira As Currency
Global Kobo As Integer
Global Naira_result As String
Global Kobo_result As String
'Global num As Currency


' Return words for this value between 1 and 999.
Public Function Words_1_999(ByVal num As Integer) As String
Dim hundreds As Integer
Dim remainder As Integer
Dim result As String

    hundreds = num \ 100
    remainder = num - hundreds * 100

    If hundreds > 0 Then
        result = Words_1_19(hundreds) & " hundred "
    End If

    If remainder > 0 Then
        result = result & Words_1_99(remainder)
    End If

    Words_1_999 = Trim$(result)
End Function
' Return a word for this value between 1 and 99.
Public Function Words_1_99(ByVal num As Integer) As String
Dim result As String
Dim tens As Integer

    tens = num \ 10

    If tens <= 1 Then
        ' 1 <= num <= 19
        result = result & " " & Words_1_19(num)
    Else
        ' 20 <= num
        ' Get the tens digit word.
        Select Case tens
            Case 2
                result = "twenty"
            Case 3
                result = "thirty"
            Case 4
                result = "forty"
            Case 5
                result = "fifty"
            Case 6
                result = "sixty"
            Case 7
                result = "seventy"
            Case 8
                result = "eighty"
            Case 9
                result = "ninety"
        End Select

        ' Add the ones digit number.
        result = result & " " & Words_1_19(num - tens * 10)
    End If

    Words_1_99 = Trim$(result)
End Function
' Return a word for this value between 1 and 19.
Public Function Words_1_19(ByVal num As Integer) As String
    Select Case num
        Case 1
            Words_1_19 = "one"
        Case 2
            Words_1_19 = "two"
        Case 3
            Words_1_19 = "three"
        Case 4
            Words_1_19 = "four"
        Case 5
            Words_1_19 = "five"
        Case 6
            Words_1_19 = "six"
        Case 7
            Words_1_19 = "seven"
        Case 8
            Words_1_19 = "eight"
        Case 9
            Words_1_19 = "nine"
        Case 10
            Words_1_19 = "ten"
        Case 11
            Words_1_19 = "eleven"
        Case 12
            Words_1_19 = "twelve"
        Case 13
            Words_1_19 = "thirteen"
        Case 14
            Words_1_19 = "fourteen"
        Case 15
            Words_1_19 = "fifteen"
        Case 16
            Words_1_19 = "sixteen"
        Case 17
            Words_1_19 = "seventeen"
        Case 18
            Words_1_19 = "eightteen"
        Case 19
            Words_1_19 = "nineteen"
    End Select
End Function
' Return a string of words to represent the
' integer part of this value.
Public Function Words_1_all(ByVal num As Currency) As String
Dim power_value(1 To 5) As Currency
Dim power_name(1 To 5) As String
Dim digits As Integer
Dim result As String
Dim i As Integer

    ' Initialize the power names and values.
    power_name(1) = "trillion": power_value(1) = 1000000000000#
    power_name(2) = "billion":  power_value(2) = 1000000000
    power_name(3) = "million":  power_value(3) = 1000000
    power_name(4) = "thousand": power_value(4) = 1000
    power_name(5) = "":         power_value(5) = 1

    For i = 1 To 5
        ' See if we have digits in this range.
        If num >= power_value(i) Then
            ' Get the digits.
            digits = Int(num / power_value(i))

            ' Add the digits to the result.
            If Len(result) > 0 Then result = result & ", "
            result = result & _
                Words_1_999(digits) & _
                " " & power_name(i)

            ' Get the number without these digits.
            num = num - digits * power_value(i)
        End If
    Next i

    Words_1_all = Trim$(result)
End Function
' Return a string of words to represent this
' currency value in Naira and Kobo.
Public Function Words_Money(ByVal num As Currency) As String
Dim Naira As Currency
Dim Kobo As Integer
Dim Naira_result As String
Dim Kobo_result As String

    ' Naira.
    Naira = Int(num)
    Naira_result = Words_1_all(Naira)
    If Len(Naira_result) = 0 Then Naira_result = "zero"

    If Naira_result = "one" Then
        Naira_result = Naira_result & " Naira"
    Else
        Naira_result = Naira_result & " Naira"
    End If

    ' Kobo.
    Kobo = CInt((num - Naira) * 100#)
    Kobo_result = Words_1_all(Kobo)
    If Len(Kobo_result) = 0 Then Kobo_result = ""

    If Kobo_result = "one" Then
        Kobo_result = Kobo_result & "  only"
    Else
        Kobo_result = Kobo_result & "  Only"
    End If

    ' Combine the results.
    Words_Money = Naira_result & Kobo_result
End Function



