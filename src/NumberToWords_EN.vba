Option Explicit

' ==========================================================
' CONVERT NUMBER TO ENGLISH WORDS
' ==========================================================

Function NumberToEnglishWordsWithFraction(ByVal WholePart As Double, _
                                          ByVal FractionPart As Long) As String

    Dim Groups As Variant
    Dim i As Long
    Dim n As Double ' Double to prevent Overflow
    Dim s As String
    Dim part As String
    
    Groups = Array("", " thousand", " million", " billion", " trillion")
    
    If WholePart = 0 Then
        s = "zero"
    Else
        i = 0
        
        Do While WholePart > 0
            
            ' Overflow protection for numbers > 2.1 billion
            n = WholePart - Int(WholePart / 1000) * 1000
            
            If n <> 0 Then
                part = ThreeDigitsToEnglish(CLng(n))
                
                If part <> "" Then
                    part = part & Groups(i)
                    If s <> "" Then
                        s = part & " " & s
                    Else
                        s = part
                    End If
                End If
            End If
            
            WholePart = Int(WholePart / 1000)
            i = i + 1
            
        Loop
    End If
    
    s = Trim(s)
    
    ' ==========================================================
    ' FRACTIONS AS /100
    ' ==========================================================
    If FractionPart > 0 Then
        s = s & " and " & Right("00" & CStr(FractionPart), 2) & "/100"
    End If
    
    NumberToEnglishWordsWithFraction = s

End Function


' ==========================================================
' CONVERT 3-DIGIT NUMBER 0-999
' ==========================================================
Private Function ThreeDigitsToEnglish(ByVal n As Long) As String

    ' Use Static arrays for memory optimization and speed
    Static Ones As Variant
    Static Tens As Variant
    Dim s As String
    
    If IsEmpty(Ones) Then
        Ones = Array("", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine", _
                     "ten", "eleven", "twelve", "thirteen", "fourteen", "fifteen", "sixteen", _
                     "seventeen", "eighteen", "nineteen")
                     
        Tens = Array("", "", "twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety")
    End If
    
    ' ==========================================================
    ' HUNDREDS
    ' ==========================================================
    If n >= 100 Then
        s = Ones(Int(n / 100)) & " hundred"
        n = n Mod 100
        If n > 0 Then s = s & " "
    End If
    
    ' ==========================================================
    ' TENS AND ONES
    ' ==========================================================
    If n >= 20 Then
        s = s & Tens(Int(n / 10))
        n = n Mod 10
        If n > 0 Then
            s = s & "-" & Ones(n) ' e.g., twenty-one
        End If
    ElseIf n > 0 Then
        s = s & Ones(n)
    End If
    
    ThreeDigitsToEnglish = Trim(s)

End Function


' ==========================================================
' WORD MACRO ENTRY POINT
' ==========================================================
Sub ConvertSelectionToEnglishWords()

    Dim SelText As String, CleanSel As String
    Dim NumPart As String, RemainderText As String
    Dim CleanNum As String, fracText As String, txt As String
    Dim Parts() As String
    
    Dim WholeNum As Double
    Dim FractionPart As Long
    Dim SpacePos As Long
    
    ' 1. GET SELECTION & CLEAN (remove accidental paragraph breaks/spaces)
    SelText = Selection.Text
    CleanSel = Replace(Replace(SelText, vbCr, ""), vbLf, "")
    CleanSel = Trim(CleanSel)
    
    If CleanSel = "" Then
        MsgBox "Please select a number!", vbExclamation, "Selection Error"
        Exit Sub
    End If
    
    ' 2. ISOLATE NUMBER AND KEEP TRAILING TEXT
    SpacePos = InStr(CleanSel, " ")
    
    If SpacePos > 0 Then
        NumPart = Left(CleanSel, SpacePos - 1)
        RemainderText = Mid(CleanSel, SpacePos) ' Keeps trailing text, e.g., " USD"
    Else
        NumPart = CleanSel
        RemainderText = ""
    End If
    
    ' 3. REMOVE THOUSANDS SEPARATORS (English format uses commas for thousands)
    CleanNum = Replace(NumPart, ",", "")
    
    ' 4. SPLIT INTEGER / FRACTION (English format uses dots for decimals)
    Parts = Split(CleanNum, ".")
    
    If Not IsNumeric(Parts(0)) Then
        MsgBox "The selection is not recognized as a valid number!", vbExclamation, "Data Error"
        Exit Sub
    End If
    
    WholeNum = CDbl(Parts(0))
    
    ' 5. CALCULATE FRACTION
    FractionPart = 0
    If UBound(Parts) > 0 Then
        fracText = Parts(1)
        
        If IsNumeric(fracText) Then
            If Len(fracText) = 1 Then
                FractionPart = CInt(fracText) * 10
            Else
                FractionPart = CInt(Left(fracText, 2))
            End If
        End If
    End If
    
    ' 6. CONVERT
    txt = NumberToEnglishWordsWithFraction(WholeNum, FractionPart)
    
    ' 7. REPLACE SELECTION (Keep the original number format, add words, keep trailing text)
    Selection.Text = NumPart & " " & txt & RemainderText

End Sub
