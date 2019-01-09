Attribute VB_Name = "SpeakAndSpell"
Option Explicit


Function NumberToWords(ByVal Mystring As String, Optional IsInteger As Boolean = True) As String

Dim MyArray() As String

Mystring = NumChunkCleanString(Mystring)

MyArray = Split(Mystring, "|")

Select Case UBound(MyArray, 1)

Case 0:
    NumberToWords = NumChunkWords(MyArray(0))
Case 1:
NumberToWords = NumChunkWords(MyArray(0))
    If Not IsInteger Then NumberToWords = NumberToWords & " POINT " & NumChunkDecimal(MyArray(1))
Case Else:
    NumberToWords = "Cannot Calculate"
End Select


End Function

Private Function NumChunkDecimal(Mystring As String) As String

Dim LP As Integer                       ':: GENERIC LOOP VARIABLE ::
Dim LUVdouble As Double       ' ::DICTIONARY LOOKUP POWER ::

For LP = 1 To Len(Mystring)

    LUVdouble = 1 / (10 ^ LP)    ' RAISE 10 TO THE POWER OF LOOP ITERATION
    
    If Mid(Mystring, LP, 1) <> "0" Then ' :: ONLY RETURN DATA IF NOT ZERO AT POSITION ::
            'RETURN TEXT VALUE OF POWER FROM
            NumChunkDecimal = NumChunkDecimal & NumChunkName(Mid(Mystring, LP, 1)) ' NUMBER NAME
            NumChunkDecimal = NumChunkDecimal & " " & NumChunkName(LUVdouble)  ' POSITION NAME
                If Mid(Mystring, LP, 1) <> "1" Then
                    NumChunkDecimal = NumChunkDecimal & "S "
                Else
                    NumChunkDecimal = NumChunkDecimal & " "
                End If
            
End If

Next LP
NumChunkDecimal = UCase(Trim(NumChunkDecimal)) ' :: UPPER AND TRIM ::

End Function

Private Function NumChunkCleanString(Mystring As String) As String

Dim LP As Integer ' :: GENERIC LOOP VARIABLE ::

' :: RETURN CLEANSED STRING

Mystring = Replace(Mystring, " ", "") ' :: REMOVE SPACES
Mystring = Replace(Mystring, ",", "") ' :: REMOVE COMMAS
Mystring = Replace(Mystring, "'", "") ' :: REMOVE APOSTROPHES
Mystring = StrReverse(Replace(StrReverse(Mystring), ".", "|", , 1))   ' :: POINT TO PIPE FOR FIRST POINT
Mystring = Replace(Mystring, Chr(46), "")  ' :: NULL ALL ADDITIONAL POINTS ::
Mystring = Replace(Mystring, Chr(133), "")  ' :: NULL ALL ADDITIONAL POINTS ::

' :: REMOVE ANY ADDITIONAL DASHES ::
If Left(Mystring, 1) = "-" And Len(Mystring) > 1 Then Mystring = "-" & Replace(Right(Mystring, Len(Mystring) - 1), "-", "")

':: REMOVE NON-NUMERIC CHARACTERS ::

For LP = 1 To Len(Mystring)

    If NumericDict(Asc(Mid(Mystring, LP, 1))) = "Y" Then NumChunkCleanString = NumChunkCleanString & Mid(Mystring, LP, 1)

Next LP


End Function



Private Function NumChunkWords(Mystring As String) As String

Dim LP As Integer       ' :: GENERIC LOOP VARIABLE ::
Dim MyChunks As Integer ' :: CHUNKS OF 3 NUMBERS ::
Dim Outstring As String ' ::  RAW OUTPUT STRING ::
Dim PadString As String '  :: ZERO-PADDED NUMBER DIVISIBLE BY THREE ::
Dim MyPowerLUV As Double  ' :: POWER OF THREE-DIGIT PORTION ::
Dim MyPos As Integer  ' :: STRING POSITION CURSOR ::

' :: only use integer portion

PadString = NumChunkPadMe(Mystring) ' zero-padded string to multiple of three

MyChunks = NumChunkCount(Mystring)  ' Count of three digit chunks

MyPos = 1
    
             ' :: CASE WHEN VALUE IS 0 ::
    If MyChunks = 1 And PadString = "000" Then
            NumChunkWords = "ZERO"
            Exit Function
    End If
    
    For LP = 1 To MyChunks
            
        MyPowerLUV = 10 ^ ((MyChunks - LP) * 3)  ' LOOKUP FOR POWER BRACKET ::
            
    
        If MyPos = MyChunks Then
            Outstring = Outstring & NumChunkWordBlock(Mid(PadString, (3 * LP) - 2, 3)) & " "

        End If
    
        If MyPos <> MyChunks Then
            If Val(Mid(PadString, (3 * LP) - 2, 3)) > 0 Then
                Outstring = Outstring & NumChunkWordBlock(Mid(PadString, (3 * LP) - 2, 3)) & " "
                Outstring = Outstring & NumChunkName(MyPowerLUV) & " "
            End If
        End If
        
        MyPos = MyPos + 1
        
        
    Next LP

NumChunkWords = UCase(Trim(Outstring))

End Function

Private Function NumChunkWordBlock(Mystring As String) As String

Dim LP As Index ' GENERIC LOOP
Dim NumHund, NumTen, NumUnit As Double '  Portions

NumChunkWordBlock = "" ' :: MAKE BLANK ::

NumHund = Val(Left(Mystring, 1)) ' HUNDREDS
NumTen = Val(Mid(Mystring, 2, 1)) ' TENS
NumUnit = Val(Right(Mystring, 1)) ' UNITS


' :: DO HUNDREDS FIRST ::

If NumHund > 0 Then
    
    NumChunkWordBlock = NumChunkName(NumHund) & " " & NumChunkName(100)

End If

':: ADD IN THE 'AND' WHERE TENS OR UNITS PORTION IS NOT ZERO

If NumTen + NumUnit > 0 And NumHund > 0 = True Then

NumChunkWordBlock = NumChunkWordBlock & " AND"

End If

':: CHECK FOR > 19

If (NumTen * 10) + NumUnit > 19 Then

    NumChunkWordBlock = NumChunkWordBlock & " " & NumChunkName(NumTen * 10)
    
    If NumUnit > 0 Then  ':: HAS UNITS?
        NumChunkWordBlock = NumChunkWordBlock & " " & NumChunkName(NumUnit)
    End If

End If

':: CHECK FOR <= 19

If (NumTen * 10) + NumUnit <= 19 And (NumTen * 10) + NumUnit > 0 Then

    NumChunkWordBlock = NumChunkWordBlock & " " & NumChunkName((NumTen * 10) + NumUnit)
    
End If


End Function


Private Function NumChunkCount(Mystring As String) As Integer

':: COUNT OF BLOCKS OF THREE DIGITS IN NUMBER ::

NumChunkCount = Len(Format(Val(Mystring), "#,###")) - Len(Format(Val(Mystring), "#")) + 1

End Function


Private Function NumChunkPadMe(ByVal Mystring As String) As String

':: PAD STRING TO NEAREST MULTIPLE OF THREE ( ROUND UP!! )

Dim MyPad As String

NumChunkPadMe = "" & ZeroPad(Mystring, NumChunkCount(Mystring) * 3)

End Function

Private Function NumericDict(ByVal MyInteger As Integer) As String

' :: RETURN ' Y ' OR  ' N ' IF VALID CHARACTER ::
Dim MyDict As Dictionary
   
   
' :: CHECK THE ASCII CODE PASSED AS MyInteger IS VALID NUMERIC CHARACTER ( OR VALID DELIMITER )

Set MyDict = New Dictionary
MyDict(48) = "Y" 'ZERO
MyDict(49) = "Y" 'ONE
MyDict(50) = "Y" 'TWO
MyDict(51) = "Y" 'THREE
MyDict(52) = "Y" 'FOUR
MyDict(53) = "Y" 'FIVE
MyDict(54) = "Y" 'SIX
MyDict(55) = "Y" 'SEVEN
MyDict(56) = "Y" 'EIGHT
MyDict(57) = "Y" 'NINE
MyDict(45) = "Y" 'DASH
MyDict(46) = "Y" 'POINT
MyDict(124) = "Y" 'PIPE

If MyDict.EXISTS(MyInteger) Then
    
    NumericDict = MyDict(MyInteger)

Else
    
    NumericDict = "N"

End If

End Function

Private Function NumChunkName(ByVal MyInteger As Double) As String

'
' ::  DICTIONARY TO CONVERT DOUBLE INPUT TO ENGLISH EQUIVALENT  ::
'

Dim MyDict As Dictionary
   
    Set MyDict = New Dictionary
     
        MyDict(0) = "zero"
        MyDict(1) = "one"
        MyDict(2) = "two"
        MyDict(3) = "three"
        MyDict(4) = "four"
        MyDict(5) = "five"
        MyDict(6) = "six"
        MyDict(7) = "seven"
        MyDict(8) = "eight"
        MyDict(9) = "nine"
        MyDict(10) = "ten"
        MyDict(11) = "eleven"
        MyDict(12) = "twelve"
        MyDict(13) = "thirteen"
        MyDict(14) = "fourteen"
        MyDict(15) = "fifteen"
        MyDict(16) = "sixteen"
        MyDict(17) = "seventeen"
        MyDict(18) = "eighteen"
        MyDict(19) = "nineteen"
        MyDict(20) = "twenty"
        MyDict(30) = "thirty"
        MyDict(40) = "forty"
        MyDict(50) = "fifty"
        MyDict(60) = "sixty"
        MyDict(70) = "seventy"
        MyDict(80) = "eighty"
        MyDict(90) = "ninety"
        MyDict(100) = "hundred"
        MyDict(1000) = "thousand"
        MyDict(1000000) = "million"
        MyDict(1000000000) = "billion"
        MyDict(1000000000000#) = "trillion"
        MyDict(1E+15) = "quadrillion"
        MyDict(1E+18) = "quintillion"
        MyDict(1E+21) = "Sextillion"
        MyDict(1E+24) = "Septillion"
        MyDict(1E+27) = "Octillion"
        MyDict(1E+30) = "Nonillion"
        MyDict(1E+33) = "Decillion"
        MyDict(0.1) = "Tenth"
        MyDict(0.01) = "Hundredth"
        MyDict(0.001) = "Thousandth"
        MyDict(0.0001) = "Ten-Thousandth"
        MyDict(0.00001) = "Hundred-Thousandth"
        MyDict(0.000001) = "Millionth"
        MyDict(0.0000001) = "Ten-Millionth"
        MyDict(0.00000001) = "Hundred-Millionth"
        MyDict(0.000000001) = "Billionth"
        MyDict(0.0000000001) = "Ten-Billionth"
        MyDict(0.00000000001) = "Hundred-Billionth"
        MyDict(0.000000000001) = "Trillionth"
        MyDict(0.0000000000001) = "Ten-Trillionth"
        MyDict(0.00000000000001) = "Hundreed-Trillionth"
        MyDict(0.000000000000001) = "Quadrillionth"
        MyDict(1E-16) = "Ten-Quadrillionth"
        MyDict(1E-17) = "Hundred-Quadrillionth"
        MyDict(1E-18) = "Quitillionth"
        MyDict(1E-19) = "Ten-Quitillionth"
        MyDict(1E-20) = "Hundred-Quitillionth"
        MyDict(1E-21) = "Sextillionth"
        MyDict(1E-22) = "Ten-Sextillionth"
        MyDict(1E-23) = "Hundred-Sextillionth"
        MyDict(1E-24) = "Septillionth"

    

If MyDict.EXISTS(MyInteger) Then
    
    NumChunkName = MyDict(MyInteger)

Else
    
    NumChunkName = ""

End If

End Function

