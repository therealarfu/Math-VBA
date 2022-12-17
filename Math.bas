Attribute VB_Name = "Math"
Option Explicit
'Math module by arfu
Public Const PI As Double = 3.14159265359, E As Double = 2.71828182846, PI2 As Double = PI / 2, TAU As Double = PI * 2, GRatio As Double = 1.61803398875


Public Function IsPrime(ByVal x#) As Boolean
    Dim c#, i%
    IsPrime = True
    For c = 2 To x - 1
        If isDivisible(x, c) = True Then IsPrime = False
    Next
End Function


Public Function isDivisible(ByVal x#, Optional ByVal y# = 2) As Boolean
    isDivisible = x Mod y = 0
End Function


Function Evaluate(ByVal String1 As String) As Double
    On Error Resume Next
    Dim Excel As Object: Set Excel = CreateObject("Excel.Application")
    Evaluate = Excel.Evaluate(String1)
End Function


Public Function Pow(ByVal x#, Optional ByVal y# = 2) As Double
    Pow = (x ^ y)
End Function


Public Function Root(ByVal x#, Optional ByVal y As Double = 2) As Double
    Root = x ^ (1 / y)
End Function


Public Function RandNum(Optional ByVal Minimum As Single, Optional ByVal Maximum As Single = 1, Optional ByVal Float As Integer, Optional RandomizeNumber As Variant) As Single
    If IsMissing(RandomizeNumber) Then
        Randomize
    Else
        Randomize RandomizeNumber
    End If
    RandNum = Round((Maximum - Minimum) * Rnd + Minimum, Float)
End Function


Public Function Ceil(ByVal x#) As Long
    Ceil = IIf(Round(x, 0) >= x, Round(x, 0), Round(x, 0) + 1)
End Function


Public Function Trunc(ByVal x#) As Long
    Trunc = IIf(x > 0, Int(x), Int(x * -1) * -1)
End Function


Public Function Floor(ByVal x#) As Long
    Floor = IIf(Round(x, 0) <= x, Round(x, 0), Round(x, 0) - 1)
End Function


Public Function Delta(ByVal a#, Optional ByVal b# = 0, Optional ByVal c# = 0) As Double
    Delta = b ^ 2 - 4 * a * c
End Function


Public Function Bhask(ByVal a#, Optional ByVal b# = 0, Optional ByVal c# = 0)
    If Delta(a, b, c) < 0 Then Exit Function
    Bhask = Array((-b + Root(Delta(a, b, c))) / (2 * a), (-b - Root(Delta(a, b, c))) / (2 * a))
End Function


Public Function Min(ParamArray x() As Variant) As Double
    Dim i%
    For i = LBound(x) To UBound(x)
        If i = 0 Or x(i) < Min Then Min = x(i)
    Next
End Function


Public Function Max(ParamArray x() As Variant) As Double
    Dim i%
    For i = LBound(x) To UBound(x)
        If i = 0 Or x(i) > Max Then Max = x(i)
    Next
End Function


Public Function GCD(ByVal a As Long, ByVal b As Long) As Long
    Dim remainder As Long
    a = Abs(a)
    b = Abs(b)
    If a = 0 Or b = 0 Then Exit Function
    Do
      remainder = a Mod b
      a = b
      b = remainder
    Loop Until remainder = 0
    GCD = a
End Function


Public Function LCM(ByVal a As Long, ByVal b As Long) As Long
    a = Abs(a)
    b = Abs(b)
    If a = 0 Or b = 0 Then Exit Function
    
    LCM = (a * b) \ GCD(a, b)
End Function


Public Function Fact(ByVal N As Long, Optional ByVal StepValue As Long = 1) As LongLong
    Fact = 1
    For N = N To 1 Step -Abs(StepValue)
        Fact = Fact * N
    Next
End Function


Public Function Fibonacci(ByVal N As Long) As Long
    If N <= 0 Then
      Fibonacci = 0
    ElseIf N = 1 Then
      Fibonacci = 1
    Else
      Fibonacci = Fibonacci(N - 1) + Fibonacci(N - 2)
    End If
End Function


Public Function Mean(ParamArray x() As Variant) As Double
    Dim i%
    For i = LBound(x) To UBound(x)
        Mean = Mean + x(i)
    Next
    Mean = Mean / (UBound(x) + 1)
End Function


Public Function Median(ParamArray x() As Variant) As Double
    Median = x(0)
    If UBound(x) = 0 Then Exit Function
    Median = IIf(UBound(x) Mod 2, (x(UBound(x) \ 2) + x(UBound(x) \ 2 + 1)) / 2, x(UBound(x) \ 2))
End Function


Public Function Variance(ByVal N1#, ByVal N2#) As Double
    Variance = (Mean(N1, N2) - N1) ^ 2 + (Mean(N1, N2) - N2) ^ 2
End Function


Public Function XMid(ByVal x1#, ByVal x2#) As Double
    XMid = (x1 + x2) / 2
End Function


Public Function YMid(ByVal y1#, ByVal y2#) As Double
    YMid = (y1 + y2) / 2
End Function


Public Function FindA(ByVal x1#, ByVal x2#, ByVal y1#, ByVal y2#) As Double
    If x1 = x2 Then Exit Function
    FindA = (y1 - y2) / (x1 - x2)
End Function


Public Function Lerp(ByVal x1#, ByVal x2#, ByVal y1#, ByVal y2#, ByVal x#) As Double
    Lerp = y1 + (x - x1) * (y2 - y1) / (x2 - x1)
End Function


Public Function LineLineIntersect(ByVal x1#, ByVal y1#, ByVal x2#, ByVal y2#, ByVal x3#, ByVal y3#, ByVal x4#, ByVal y4#)
    Dim x As Double, y As Double
    If (x1 - x2) * (y3 - y4) = (y1 - y2) * (x3 - x4) Then Exit Function
    x = ((x1 * y2 - y1 * x2) * (x3 - x4) - (x1 - x2) * (x3 * y4 - y3 * x4)) / ((x1 - x2) * (y3 - y4) - (y1 - y2) * (x3 - x4))
    y = ((x1 * y2 - y1 * x2) * (y3 - y4) - (y1 - y2) * (x3 * y4 - y3 * x4)) / ((x1 - x2) * (y3 - y4) - (y1 - y2) * (x3 - x4))
    LineLineIntersect = Array(x, y)
End Function


Public Function Distance(ByVal x1#, ByVal x2#, ByVal y1#, ByVal y2#, Optional ByVal Sqrt As Boolean = True) As Double
    Distance = IIf(Sqrt, Root((x2 - x1) ^ 2 + (y2 - y1) ^ 2), (x2 - x1) ^ 2 + (y2 - y1) ^ 2)
End Function


Public Function Distance2(ByVal x1#, ByVal x2#, ByVal y1#, ByVal y2#, ByVal Z1#, ByVal Z2#, Optional ByVal Sqrt As Boolean = True) As Double
    Distance2 = IIf(Sqrt, Root((x2 - x1) ^ 2 + (y2 - y1) ^ 2 + (Z2 - Z1) ^ 2), (x2 - x1) ^ 2 + (y2 - y1) ^ 2 + (Z2 - Z1) ^ 2)
End Function


Public Function Hypot(ByVal x#, ByVal y#) As Double
    Hypot = Root(x ^ 2 + y ^ 2)
End Function


Public Function LogN(ByVal x#, ByVal y#) As Double
    LogN = Log(x) / Log(y)
End Function


Public Function ATan2(ByVal x#, ByVal y#) As Double
    ATan2 = IIf(x > 0, Atn(y / x), IIf(x < 0, Atn(y / x) + PI * Sgn(y) + IIf(y = 0, PI, 0), PI2 * Sgn(y)))
End Function


Public Function Sec(ByVal x#) As Double
    Sec = 1 / Cos(x)
End Function


Public Function Cosec(ByVal x#) As Double
    Cosec = 1 / Sin(x)
End Function


Public Function Cotan(ByVal x#) As Double
    Cotan = 1 / Tan(x)
End Function


Public Function Radians(ByVal x#) As Double
    Radians = x * 180 / PI
End Function


Public Function Degrees(ByVal x#) As Double
    Degrees = x * PI / 180
End Function


Public Function ASin(ByVal x#) As Double
    ASin = Atn(x / Root(-x * x + 1))
End Function


Public Function ACos(ByVal x#) As Double
    ACos = Atn(-x / Root(-x * x + 1)) + 2 * Atn(1)
End Function


Public Function ASec(ByVal x#) As Double
    ASec = Atn(x / Root(x * x - 1)) + Sgn((x) - 1) * (2 * Atn(1))
End Function


Public Function ACosec(ByVal x#) As Double
    ACosec = Atn(x / Root(x * x - 1)) + (Sgn(x) - 1) * (2 * Atn(1))
End Function


Public Function ACotan(ByVal x#) As Double
    ACotan = Atn(x) + 2 * Atn(1)
End Function


Public Function HSin(ByVal x#) As Double
    HSin = (Exp(x) - Exp(-x)) / 2
End Function


Public Function HCos(ByVal x#) As Double
    HCos = (Exp(x) + Exp(-x)) / 2
End Function


Public Function HTan(ByVal x#) As Double
    HTan = (Exp(x) - Exp(-x)) / (Exp(x) + Exp(-x))
End Function


Public Function HSec(ByVal x#) As Double
    HSec = 2 / (Exp(x) + Exp(-x))
End Function


Public Function HCosec(ByVal x#) As Double
    HCosec = 2 / (Exp(x) - Exp(-x))
End Function


Public Function HCotan(ByVal x#) As Double
    HCotan = (Exp(x) + Exp(-x)) / (Exp(x) - Exp(-x))
End Function


Public Function HASin(ByVal x#) As Double
    HASin = Log(x + Root(x * x + 1))
End Function


Public Function HACos(ByVal x#) As Double
    HACos = Log(x + Root(x * x - 1))
End Function


Public Function HATan(ByVal x#) As Double
    HATan = Log((1 + x) / (1 - x)) / 2
End Function


Public Function HASec(ByVal x#) As Double
    HASec = Log((Root(-x * x + 1) + 1) / x)
End Function


Public Function HACosec(ByVal x#) As Double
    HACosec = Log((Sgn(x) * Root(x * x + 1) + 1) / x)
End Function


Public Function HACotan(ByVal x#) As Double
    HACotan = Log((x + 1) / (x - 1)) / 2
End Function

