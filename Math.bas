Attribute VB_Name = "Math"
'Math module by arfu
Option Explicit
#If Win64 Then
    #If VBA7 Then
        Public Const MAX_INTEGER As LongLong = 2 ^ 63 - 1
    #Else
        Public Const MAX_INTEGER As Long = 2 ^ 63 - 1
    #End If
#Else
    Public Const MAX_INTEGER As Long = 2 ^ 31 - 1
#End If
Public Const PI As Double = 3.14159265359, E As Double = 2.71828182846, PI2 As Double = 1.57079632679, TAU As Double = 6.28318530718, GRatio As Double = 1.61803398875

Public Function IsPrime(ByVal X As Long) As Boolean
    Dim c As Integer
    IsPrime = True
    For c = 2 To Abs(X) - 1
        If Abs(X) Mod c = 0 Then
            IsPrime = False
            Exit Function
        End If
    Next
    If X = 0 Then IsPrime = False
End Function


Public Function Odd(ByVal Number As Long) As Long
    Odd = Number * 2 + 1
End Function


Public Function isDivisible(ByVal Number#, Optional ByVal DividedBy# = 2) As Boolean
    isDivisible = Number Mod DividedBy = 0
End Function


Public Function Evaluate(ByVal String1 As String) As Double
    On Error Resume Next
    Dim Excel As Object: Set Excel = CreateObject("Excel.Application")
    Evaluate = Excel.Evaluate(String1)
End Function

    
Public Function Pow(ByVal X#, Optional ByVal y# = 2) As Double
    Pow = (X ^ y)
End Function


Public Function Root(ByVal X#, Optional ByVal y As Double = 2) As Double
    Root = Abs(X) ^ (1 / y)
End Function


Public Function RandomNum(Optional ByVal Minimum As Single, Optional ByVal Maximum As Single = 1, Optional ByVal Float As Integer, Optional RandomizeNumber As Variant) As Single
    If IsMissing(RandomizeNumber) Then
        Randomize
    Else
        Randomize RandomizeNumber
    End If
    RandomNum = Round((Maximum - Minimum) * Rnd + Minimum, Float)
End Function


Public Function Ceil(ByVal X#) As Long
    Ceil = IIf(Round(X, 0) >= X, Round(X, 0), Round(X, 0) + 1)
End Function


Public Function Trunc(ByVal X#) As Long
    Trunc = IIf(X > 0, Int(X), -Int(-X))
End Function


Public Function Floor(ByVal X#) As Long
    Floor = IIf(Round(X, 0) <= X, Round(X, 0), Round(X, 0) - 1)
End Function


Public Function Delta(ByVal a#, Optional ByVal b# = 0, Optional ByVal c# = 0) As Double
    Delta = b ^ 2 - 4 * a * c
End Function


Public Function Bhask(ByVal a#, Optional ByVal b# = 0, Optional ByVal c# = 0)
    If Delta(a, b, c) < 0 Then Exit Function
    Bhask = Array((-b + Sqr(Delta(a, b, c))) / (2 * a), (-b - Sqr(Delta(a, b, c))) / (2 * a))
End Function


Public Function Min(ParamArray X() As Variant) As Double
    Dim i%
    For i = LBound(X) To UBound(X)
        If i = 0 Or X(i) < Min Then Min = X(i)
    Next
End Function


Public Function Max(ParamArray X() As Variant) As Double
    Dim i%
    For i = LBound(X) To UBound(X)
        If i = 0 Or X(i) > Max Then Max = X(i)
    Next
End Function


Public Function GCD(ByVal a As Long, ByVal b As Long) As Long
    Dim remainder As Long
    If a = 0 Or b = 0 Then Exit Function
    Do
      remainder = Abs(a) Mod Abs(b)
      a = Abs(b)
      b = remainder
    Loop Until remainder = 0
    GCD = a
End Function


Public Function LCM(ByVal a As Long, ByVal b As Long) As Long
    If a = 0 Or b = 0 Then Exit Function
    LCM = (Abs(a) * Abs(b)) \ GCD(a, b)
End Function


Public Function Fact(ByVal N As Long, Optional ByVal StepValue As Long = 1) As Long
    Fact = 1
    For N = N To 1 Step -Abs(StepValue)
        Fact = Fact * N
    Next
End Function


Public Function Fibonacci(ByVal N As Long) As Long
    If N <= 0 Then Exit Function
    Fibonacci = IIf(N = 1, 1, Fibonacci(N - 1) + Fibonacci(N - 2))
End Function


Public Function Mean(ParamArray X() As Variant) As Double
    Dim i%
    For i = LBound(X) To UBound(X)
        Mean = Mean + X(i)
    Next
    Mean = Mean / (UBound(X) + 1)
End Function


Public Function Median(ParamArray X() As Variant) As Double
    Median = X(0)
    If UBound(X) = 0 Then Exit Function
    Median = IIf(UBound(X) Mod 2, (X(UBound(X) \ 2) + X(UBound(X) \ 2 + 1)) / 2, X(UBound(X) \ 2))
End Function


Public Function Variance(ByVal N1#, ByVal N2#) As Double
    Variance = (Mean(N1, N2) - N1) ^ 2 + (Mean(N1, N2) - N2) ^ 2
End Function


Public Function Mid(ByVal X1#, ByVal X2#) As Double
    Mid = (X1 + X2) / 2
End Function


Public Function FindA(ByVal X1#, ByVal X2#, ByVal Y1#, ByVal Y2#) As Double
    If X1 = X2 Then Exit Function
    FindA = (Y1 - Y2) / (X1 - X2)
End Function


Public Function Lerp(ByVal X1#, ByVal X2#, ByVal Y1#, ByVal Y2#, ByVal X#) As Double
    If X1 = X2 Then Exit Function
    Lerp = Y1 + (X - X1) * (Y2 - Y1) / (X2 - X1)
End Function


Public Function LineLineIntersect(ByVal X1#, ByVal Y1#, ByVal X2#, ByVal Y2#, ByVal x3#, ByVal y3#, ByVal x4#, ByVal y4#)
    Dim X As Double, y As Double
    If (X1 - X2) * (y3 - y4) = (Y1 - Y2) * (x3 - x4) Then Exit Function
    X = ((X1 * Y2 - Y1 * X2) * (x3 - x4) - (X1 - X2) * (x3 * y4 - y3 * x4)) / ((X1 - X2) * (y3 - y4) - (Y1 - Y2) * (x3 - x4))
    y = ((X1 * Y2 - Y1 * X2) * (y3 - y4) - (Y1 - Y2) * (x3 * y4 - y3 * x4)) / ((X1 - X2) * (y3 - y4) - (Y1 - Y2) * (x3 - x4))
    LineLineIntersect = Array(X, y)
End Function


Public Function Distance(ByVal X1#, ByVal X2#, ByVal Y1#, ByVal Y2#, Optional ByVal Sqrt As Boolean = True) As Double
    Distance = IIf(Sqrt, Sqr((X2 - X1) ^ 2 + (Y2 - Y1) ^ 2), (X2 - X1) ^ 2 + (Y2 - Y1) ^ 2)
End Function


Public Function Distance2(ByVal X1#, ByVal X2#, ByVal Y1#, ByVal Y2#, ByVal Z1#, ByVal Z2#, Optional ByVal Sqrt As Boolean = True) As Double
    Distance2 = IIf(Sqrt, Sqr((X2 - X1) ^ 2 + (Y2 - Y1) ^ 2 + (Z2 - Z1) ^ 2), (X2 - X1) ^ 2 + (Y2 - Y1) ^ 2 + (Z2 - Z1) ^ 2)
End Function


Public Function Hypot(ByVal X#, ByVal y#) As Double
    Hypot = Sqr(X ^ 2 + y ^ 2)
End Function


Public Function LogN(ByVal X#, ByVal y#) As Double
    LogN = Log(X) / Log(y)
End Function


Public Function ATn2(ByVal X#, ByVal y#) As Double
    ATn2 = IIf(X > 0, Atn(y / X), IIf(X < 0, Atn(y / X) + PI * Sgn(y) + IIf(y = 0, PI, 0), PI / 2 * Sgn(y)))
End Function


Public Function Sec(ByVal X#) As Double
    Sec = 1 / Cos(X)
End Function


Public Function Cosec(ByVal X#) As Double
    Cosec = 1 / Sin(X)
End Function


Public Function Cotan(ByVal X#) As Double
    Cotan = 1 / Tan(X)
End Function


Public Function Radians(ByVal Degrees#) As Double
    Radians = Degrees * 180 / PI
End Function


Public Function Degrees(ByVal Radians#) As Double
    Radians = Radians * PI / 180
End Function


Public Function ASin(ByVal X#) As Double
    ASin = Atn(X / Sqr(-X * X + 1))
End Function


Public Function ACos(ByVal X#) As Double
    ACos = Atn(-X / Sqr(-X * X + 1)) + 2 * Atn(1)
End Function


Public Function ASec(ByVal X#) As Double
    ASec = Atn(X / Sqr(X * X - 1)) + Sgn((X) - 1) * (2 * Atn(1))
End Function


Public Function ACosec(ByVal X#) As Double
    ACosec = Atn(X / Sqr(X * X - 1)) + (Sgn(X) - 1) * (2 * Atn(1))
End Function


Public Function ACotan(ByVal X#) As Double
    ACotan = Atn(X) + 2 * Atn(1)
End Function


Public Function HSin(ByVal X#) As Double
    HSin = (Exp(X) - Exp(-X)) / 2
End Function


Public Function HCos(ByVal X#) As Double
    HCos = (Exp(X) + Exp(-X)) / 2
End Function


Public Function HTan(ByVal X#) As Double
    HTan = (Exp(X) - Exp(-X)) / (Exp(X) + Exp(-X))
End Function


Public Function HSec(ByVal X#) As Double
    HSec = 2 / (Exp(X) + Exp(-X))
End Function


Public Function HCosec(ByVal X#) As Double
    HCosec = 2 / (Exp(X) - Exp(-X))
End Function


Public Function HCotan(ByVal X#) As Double
    HCotan = (Exp(X) + Exp(-X)) / (Exp(X) - Exp(-X))
End Function


Public Function HASin(ByVal X#) As Double
    HASin = Log(X + Sqr(X * X + 1))
End Function


Public Function HACos(ByVal X#) As Double
    HACos = Log(X + Sqr(X * X - 1))
End Function


Public Function HATan(ByVal X#) As Double
    HATan = Log((1 + X) / (1 - X)) / 2
End Function


Public Function HASec(ByVal X#) As Double
    HASec = Log((Sqr(-X * X + 1) + 1) / X)
End Function


Public Function HACosec(ByVal X#) As Double
    HACosec = Log((Sgn(X) * Sqr(X * X + 1) + 1) / X)
End Function


Public Function HACotan(ByVal X#) As Double
    HACotan = Log((X + 1) / (X - 1)) / 2
End Function

Public Function LawCos(ByVal b As Double, ByVal c As Double, ByVal Angle As Double) As Double
    LawCos = b ^ 2 + c ^ 2 - 2 * c * Cos(Angle)
End Function

