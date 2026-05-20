Attribute VB_Name = "QuantFunctionsV1.7"
'Version 1.7

Sub absopen()
'
' absopen Macro
' by Mark Blakely
' Keyboard Shortcut: Ctrl+m
'
    Dim name As FileDialog
    
    Set name = Application.FileDialog(msoFileDialogFilePicker)
    
    Dim vrtSelectedItem As Variant
    
    With name
    
    If .Show = -1 Then
    
    For Each vrtSelectedItem In .SelectedItems
    
    
    Workbooks.OpenText Filename:= _
        vrtSelectedItem, Origin _
        :=437, StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote _
        , ConsecutiveDelimiter:=True, Tab:=True, Semicolon:=False, Comma:=False _
        , Space:=True, Other:=False, FieldInfo:=Array(Array(1, 1), Array(2, 1), Array _
        (3, 1)), TrailingMinusNumbers:=True
        
     Next vrtSelectedItem
    Else
    End If
    End With
 
    Set fd = Nothing
    
End Sub
Function LinestLabled(known_ys As Variant, known_xs As Variant, cnst As Boolean, stats As Boolean) As Variant
'Returns Linest with lables in adjacent cells.

Dim linefit As Variant
linefit = Application.WorksheetFunction.LinEst(known_ys, known_xs, cnst, stats)

'Dim test(8, 2) As Variant
'test = [{1,2{3,4},{5,6},{7,8},{9,10},{11,12},{13,14},{15,16}]

Dim linefitLabeled(4, 3) As Variant
Dim i As Long
Dim j As Long

linefitLabeled(0, 0) = "m"
linefitLabeled(0, 1) = linefit(1, 1)
linefitLabeled(0, 2) = linefit(1, 2)
linefitLabeled(0, 3) = "b"
linefitLabeled(1, 0) = "um"
linefitLabeled(1, 1) = linefit(2, 1)
linefitLabeled(1, 2) = linefit(2, 2)
linefitLabeled(1, 3) = "ub"
linefitLabeled(2, 0) = "R" & ChrW(178)
linefitLabeled(2, 1) = linefit(3, 1)
linefitLabeled(2, 2) = linefit(3, 2)
linefitLabeled(2, 3) = "sy"
linefitLabeled(3, 0) = "F"
linefitLabeled(3, 1) = linefit(4, 1)
linefitLabeled(3, 2) = linefit(4, 2)
linefitLabeled(3, 3) = "DF"
linefitLabeled(4, 0) = "ssReg"
linefitLabeled(4, 1) = linefit(5, 1)
linefitLabeled(4, 2) = linefit(5, 2)
linefitLabeled(4, 3) = "ssResid"

RSet linefitLabeled(0, 0) = "m"

LinestLabled = linefitLabeled
End Function
Function Pwr(x)
Dim ten As Double
ten = Application.Power(10, x)

End Function
Function AlphaFrac(ph, Hlevel, pKaArray As Variant) As Double
'Returns alpha fraction for an weak acid or base species at a given pH and protonation level.
'Author: Nathan Boland

'Calculate [H+]
Dim H As Double
H = 10 ^ (-1 * ph)

'Take pKaArray input and convert to an array of Kas
Dim Ks() As Double
Dim size As Integer
Dim arr As Variant
arr = pKaArray

'Determine length of pKaArray whether a value or an array
If IsArray(pKaArray) Then size = UBound(arr) Else size = 1
ReDim Ks(size) As Double
Dim i As Integer
Ks(0) = 1

For i = 1 To size
    Ks(i) = 10 ^ (-1 * pKaArray(i))
Next i

'Calculate beta array
Dim x As Integer
Dim Betas() As Double
ReDim Betas(size + 1) As Double
For x = 1 To size + 1
    If Betas(x) = 0 Then Betas(x) = 1
    For y = 0 To x - 1
        Betas(x) = Betas(x) * Ks(y)
    Next y
Next x

'Calculate array of alpha fraction terms of appropriate size and format
Dim arrD() As Double
ReDim arrD(size + 1) As Double
Dim j As Integer

For j = j To size + 1
    arrD(j) = H ^ (size + 1 - j) * Betas(j)
Next j

'Calculate alpha numerator and denominator
Dim numerator As Double
Dim denominator As Double
Dim D As Double

numerator = arrD(size + 1 - Hlevel)
For D = 0 To (size + 1)
    denominator = denominator + arrD(D)
Next D

AlphaFrac = numerator / denominator

End Function
Function UncInX(y, k, m, s_y, rangex, rangey) As Double
'Returns the uncertainty in an unknown X value determined from a calibration curve and a measured Y value using Equation 4.27
'Source: Harris, D. C.; Quantitative Chemical Analysis, 9th ed.; W.H. Freeman and Company: New York, 2016

Dim n As Integer
Dim ave_y As Double
Dim devsqr_x As Double
Dim abs_m As Double
Dim sqr As Double
Dim test As Double

n = Application.Count(rangey)
ave_y = Application.Average(rangey)
devsqr_x = Application.DevSq(rangex)
sqr = (1 / k + 1 / n + ((y - ave_y) ^ 2 / (m ^ 2 * devsqr_x))) ^ 0.5
abs_m = Abs(m)

UncInX = s_y / abs_m * sqr

End Function
Function SciExp(Expon As Double, Optional SigFig As Integer) As String
'Returns number in text scientific notation format (for labels). Can provide 3(default) to 1 sig figs.

Dim meExpon As String
Dim mantissa As String
Dim exponent As String
Dim i As Integer
Dim sign As String
Dim e As String
Dim Dig As String
Dim u As Long

'divide Excel scientific notation into mantissa and exponent
    meExpon = Format(Expon, "Scientific")
    mantissa = Left(meExpon, InStr(meExpon, "E") - 1)
    exponent = Right(meExpon, Len(meExpon) - InStr(meExpon, "E") - 1)

'trim mantissa by sig figs specified (between 1 and 3), default sig figs is 3
    If SigFig = 2 Then
        mantissa = Left(mantissa, 3)
    ElseIf SigFig = 1 Then
        mantissa = Left(mantissa, 1)
    ElseIf IsMissing(SigFig) Then
        SigFig = 3
    Else
    End If

'If no exponent then just return the mantissa
    If exponent = 0 Then
        SciExp = mantissa
        Exit Function
    End If
    
'Ignore "+" exponent, create superscript "-" exponent
    sign = Mid(meExpon, InStr(meExpon, "E") + 1, 1)
    If sign = "+" Then
        sign = ""
    Else
        sign = ChrW(8315)
    End If
    
'Trim leading zeroes from exponent
    If InStr(exponent, "0") = 1 Then
        exponent = Right(exponent, 1)
    End If
    

'Convert exponent to superscript characters
    e = ""

    For i = 1 To Len(exponent)
        Dig = CStr(Mid(exponent, i, 1))

        Select Case Dig
            Case "0"
                u = 8304
            Case "1"
                u = 185
            Case "2"
                u = 178
            Case "3"
                u = 179
            Case "4"
                u = 8308
            Case "5"
                u = 8309
            Case "6"
                u = 8310
            Case "7"
                u = 8311
            Case "8"
                u = 8312
            Case "9"
                u = 8313
            Case "-"
                u = 8315
            Case "+"
                u = 8314
        End Select
        
        e = e & ChrW(u)
    Next i

'output
    SciExp = mantissa & ChrW(215) & "10" & sign & e
End Function
Function f()
'Returns Faraday's Constant in C/mol
'Source: 2014 CODATA Values from http://physics.nist.gov/cuu/constants/index.html (Accessed August 2018)

f = 96485.33289

End Function
Function e()
'Returns the elementary charge in Coulombs
'Source: 2014 CODATA Values from http://physics.nist.gov/cuu/constants/index.html (Accessed August 2018)

e = 1.6021766208E-19

End Function
Function k()
'Returns Boltzmann's constant in J/K
'Source: 2014 CODATA Values from http://physics.nist.gov/cuu/constants/index.html (Accessed August 2018)

k = 1.380648524E-23

End Function
Function H()
'Returns Planck's constant in J*s
'Source: 2014 CODATA Values from http://physics.nist.gov/cuu/constants/index.html (Accessed August 2018)

H = 6.62607004E-34

End Function
Function GasJ()
'Returns the Gas constant in J/(mol*K)
'Source: 2014 CODATA Values from http://physics.nist.gov/cuu/constants/index.html (Accessed August 2018)

GasJ = 8.3144598

End Function
Function GasLatm()
'Returns the Gas constant in L*atm /(mol*K)
'Source: 2014 CODATA Values from http://physics.nist.gov/cuu/constants/index.html (Accessed August 2018). Converted from GasJ()

GasLatm = 0.082057338

End Function
Function Avogadro()
'Returns Avogadro's number
'Source: 2014 CODATA Values from http://physics.nist.gov/cuu/constants/index.html (Accessed August 2018)

Avogadro = 6.022140857E+23

End Function
Function QUADRATIC(A, b, c, Optional result)
'Returns the one of two x values from solving the
'quadratic equation (a*x^2 + bx + c = 0)

Dim x As Double
Dim xpos As Double
Dim xneg As Double
Dim root As Double
Dim r As Double

If IsMissing(result) Then result = 0

' Calculate
r = (b ^ 2 - (4 * A * c)) ^ 0.5
xpos = ((-(b) + r) / (2 * A))
xneg = ((-(b) - r) / (2 * A))
If result = 0 Then x = xpos Else x = xneg
QUADRATIC = x

End Function
Function DAVIESG(i, z, Optional A, Optional b, Optional temp)
'Returns the activity coefficient (gamma)
'calculated using Davies Equation where A = 0.50886 by default
'and B = 0.3 by default (vs. 0.2) in final term
'Source: W. Stumm and J. J. Morgan, Aquatic Chemistry: Chemical Equilibria and Rates in Natural Waters, John Wiley & Sons, Inc., New York, NY, 3rd edn., 1996.

Dim log_g As Double
Dim a_varT As Double

'Calculate Temperature adjusted A value using A regression fit of Helgeson/Kirham data (25-225 deg. C) and Harened/Owen data (0-25 deg. C)

If IsMissing(temp) Then temp = 25

a_varT = 0.50886 - 0.0008 * (temp - 25) + 0.00001 * (temp - 25) * (temp - 25)

If IsMissing(A) Then A = a_varT
If IsMissing(b) Then b = 0.3
' Calculate
log_g = -1 * A * z ^ 2 * (i ^ 0.5 / (1 + i ^ 0.5) - b * i)
DAVIESG = 10 ^ log_g

End Function
Function GINV(alpha, n, Optional tails)
'Returns the critical Grubbs value for a data
'set. This value can be used to determine
'whether a given data point in a data set
'is an outlier. The value is the number of
'standard deviations away from the mean at
'which point values are considered outliers.
'Source: http://www.itl.nist.gov/div898/handbook/eda/section3/eda35h1.htm

Dim alpha_crit As Double
Dim tcrit As Double

' Use 0.05 as the default value of alpha
If IsMissing(alpha) Then alpha = 0.05
If IsMissing(tails) Then tails = 1
' Calculate critical Grubb's value
alpha_crit = 2 * alpha / (tails * n)
tcrit = Application.Tinv(alpha_crit, n - 2)
GINV = (n - 1) / ((n) ^ 0.5) * (tcrit ^ 2 / (n - 2 + tcrit ^ 2)) ^ 0.5

End Function
Sub AddUDFToCustomCategory()
    ' Provide guidance when function is used. Is NOT compatible with all versions of Excel.
Dim arrCI As Variant
Dim arrGINV As Variant
Dim arrDAVIESG As Variant
Dim arrQUAD As Variant
Dim arrSciExp As Variant
Dim arrUncInX As Variant
Dim arrAlphaFrac As Variant

arrCI = Array("range of x values in regression", "range of y values in regression", "x value at which to calculate confidence limit", "alpha (optional)")
arrGINV = Array("alpha for confidence level (default 0.05 for 95%)", "number of data points", "number of tails in test (1 or 2, default 1)")
arrDAVIESG = Array("ionic strength of solution in mol/L", "charge of species", "A coefficient (default is 0.50886 for water at 25 degrees C. Note Harris lists 0.51.)", "B Coefficient (default is 0.3 as recommended by Hering and Morel 1993 and is used by Harris, 0.2 was originally published by Davies 1962, but he later recommended 0.3)", "temperature in deg. C. (default is 25) This value is used for correcting the A coefficient according to a regression of data from Helgeson/Kirham (25-225 deg C) and Harned/Owen (0-25 deg C).")
arrQUAD = Array("a", "b", "c", "Optional. Result = 0 is default and returns addition result, specify in result = 1 returns subtraction result.")
arrSciExp = Array("takes an input i and makes it a 10^i (where i is superscripted)")
arrUncInX = Array("(Average) Unknown Y value", "Number of measurements of unknown", "slope of calibration curve", "standard error in y from calibration curve", "range of x values from calibration data", "range of y values from calibration data")
arrAlphaFrac = Array("pH", "protonation level of desired species", "Range of cells containing pKa values. pKas must be in a continuous column starting with pKa1.")

Application.MacroOptions Macro:="LCI", Description:="Returns the lower confidence limit y value for a linear regression", Category:="Science", ArgumentDescriptions:=arrCI
Application.MacroOptions Macro:="UCI", Description:="Returns the upper confidence limit y value for a linear regression", Category:="Science", ArgumentDescriptions:=arrCI
Application.MacroOptions Macro:="GINV", Description:="Returns the critical Grubb's Test value", Category:="Science", ArgumentDescriptions:=arrGINV
Application.MacroOptions Macro:="DAVIESG", Description:="Returns the activity coefficient according to Davies Equation", Category:="Science", ArgumentDescriptions:=arrDAVIESG
Application.MacroOptions Macro:="QUADRATIC", Description:="Returns one root of a quadratic equation of the form a*x^2 + bx + c = 0, addition root is given by default (optional result = 0), the subtraction root is given when result = 1.", Category:="Science", ArgumentDescriptions:=arrQUAD
Application.MacroOptions Macro:="F", Description:="Returns Faraday's constant", Category:="Science"
Application.MacroOptions Macro:="e", Description:="Returns the elementary charge", Category:="Science"
Application.MacroOptions Macro:="k", Description:="Returns Boltzmann's constant", Category:="Science"
Application.MacroOptions Macro:="h", Description:="Returns Planck's constant", Category:="Science"
Application.MacroOptions Macro:="GasJ", Description:="Returns the gas constant R in J/(mol*K)", Category:="Science"
Application.MacroOptions Macro:="GasLatm", Description:="Returns the gas constant R in L*atm/(mol*K)", Category:="Science"
Application.MacroOptions Macro:="Avogadro", Description:="Returns Avogadro's Number", Category:="Science"
Application.MacroOptions Macro:="SciExp", Description:="Returns a number as an exponent for scientific notation (x10^)", Category:="Science", ArgumentDescriptions:=arrSciExp
Application.MacroOptions Macro:="UncInX", Description:="Calculates uncertainty in x from measurements of an unknown and calibration data. Uses Eqn 4-27 from Quantitative Methods of Analysis, 8th Ed. by Daniel C. Harris.", Category:="Science", ArgumentDescriptions:=arrUncInX
Application.MacroOptions Macro:="AlphaFrac", Description:="Calculates the fraction (alpha) of an acid/base species at a given pH.", Category:="Science", ArgumentDescriptions:=arrAlphaFrac
End Sub

Function NOWTIME()
NOWTIME = ActiveCell.value = Format(Now(), "h:mm:ss")
End Function

Function SEinSS(SumSquares, YcalcRange, ParameterRange) As Double

Ycs = Application.Count(YcalcRange)
Ps = Application.Count(ParameterRange)

SEinSS = sqr(SumSquares / (Ycs - Ps))

End Function

