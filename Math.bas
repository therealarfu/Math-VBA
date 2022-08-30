Attribute VB_Name = "Math"
Option Explicit
'Math module by arfu
Public Const PI As Double = 3.14159265359, E As Double = 2.71828182846, PI2 As Double = PI / 2, TAU As Double = PI * 2
Public Function IsPrime(ByVal X#) As Boolean
    Dim c#, i%
    IsPrime = True
    For c = 2 To X - 1
        If isDivisible(X, c) = True Then IsPrime = False
    Next
End Function
Public Function isDivisible(ByVal X#, Optional ByVal Y# = 2) As Boolean
    isDivisible = X Mod Y = 0
End Function
Function Evaluate(ByVal String1 As String) As Double
    On Error Resume Next
    Dim Excel As Object: Set Excel = CreateObject("Excel.Application")
    Evaluate = Excel.Evaluate(String1)
End Function
Public Function Pow(ByVal X#, Optional ByVal Y# = 2) As Double
    Pow = (X ^ Y)
End Function
Public Function Root(ByVal X#, Optional ByVal Y As Double = 2) As Double
    On Error Resume Next
    Root = X ^ (1 / Y)
End Function
Public Function Random() As Double
    Random = 1 * rnd + 0
End Function
Public Function RandNum(Optional ByVal Min As Single = 0, Optional ByVal Max As Single = 1, Optional ByVal Float As Integer = 0) As Double
    RandNum = Round(Max * rnd + Min, Float)
End Function
Public Function Ceil(ByVal X#) As Long
    Ceil = IIf(Round(X, 0) >= X, Round(X, 0), Round(X, 0) + 1)
End Function
Public Function Trunc(ByVal X#) As Long
    Trunc = IIf(X > 0, Int(X), Int(X * -1) * -1)
End Function
Public Function Floor(ByVal X#) As Long
    Floor = IIf(Round(X, 0) <= X, Round(X, 0), Round(X, 0) - 1)
End Function
Public Function Delta(ByVal a#, Optional ByVal b# = 0, Optional ByVal c# = 0) As Double
    Delta = b ^ 2 - 4 * a * c
End Function
Public Function Bhask(ByVal a#, Optional ByVal b# = 0, Optional ByVal c# = 0, Optional ByVal X As Boolean) As Double
    Dim D As Double: D = b ^ 2 - 4 * a * c
    If D < 0 Then Exit Function
    Bhask = IIf(X, (-b + Root(D)) / (2 * a), (-b - Root(D)) / (2 * a))
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
Public Function Factorial(X As Long) As LongLong
    Dim i As Long
    Factorial = 1
    For i = 1 To X
        Factorial = Factorial * i
    Next
End Function
Public Function Mean(ParamArray X() As Variant) As Double
    Dim i%
    For i = LBound(X) To UBound(X)
        Mean = Mean + X(i)
    Next
    Mean = Mean / (UBound(X) + 1)
End Function
Public Function XMid(ByVal X1#, ByVal X2#) As Double
    XMid = (X1 + X2) / 2
End Function
Public Function YMid(ByVal Y1#, ByVal Y2#) As Double
    YMid = (Y1 + Y2) / 2
End Function
Public Function Distance(ByVal X1#, ByVal X2#, ByVal Y1#, ByVal Y2#) As Double
    Distance = Root((X2 - X1) ^ 2 + (Y2 - Y1) ^ 2)
End Function
Public Function Distance2(ByVal X1#, ByVal X2#, ByVal Y1#, ByVal Y2#, ByVal Z1#, ByVal Z2#) As Double
    Distance2 = Root((X2 - X1) ^ 2 + (Y2 - Y1) ^ 2 + (Z2 - Z1) ^ 2)
End Function
Public Function Hypot(ByVal X#, ByVal Y#) As Double
    Hypot = Root(X ^ 2 + Y ^ 2)
End Function
Public Function LogN(ByVal X#, ByVal Y#) As Double
    LogN = Log(X) / Log(Y)
End Function
Public Function CosLaw(ByVal b#, ByVal c#, ByVal Angle#) As Double
    CosLaw = Root(b ^ 2 + c ^ 2 - 2 * b * c * Cos(Degrees(Angle)))
End Function
Public Function ATan2(ByVal X#, ByVal Y#) As Double
    ATan2 = IIf(X > 0, Atn(Y / X), IIf(X < 0, Atn(Y / X) + PI * Sgn(Y) + IIf(Y = 0, PI, 0), PI2 * Sgn(Y)))
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
Public Function Radians(ByVal X#) As Double
    Radians = X * 180 / PI
End Function
Public Function Degrees(ByVal X#) As Double
    Degrees = X * PI / 180
End Function
Public Function ASin(ByVal X#) As Double
    ASin = Atn(X / Root(-X * X + 1))
End Function
Public Function ACos(ByVal X#) As Double
    ACos = Atn(-X / Root(-X * X + 1)) + 2 * Atn(1)
End Function
Public Function ASec(ByVal X#) As Double
    ASec = Atn(X / Root(X * X - 1)) + Sgn((X) - 1) * (2 * Atn(1))
End Function
Public Function ACosec(ByVal X#) As Double
    ACosec = Atn(X / Root(X * X - 1)) + (Sgn(X) - 1) * (2 * Atn(1))
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
    HASin = Log(X + Root(X * X + 1))
End Function
Public Function HACos(ByVal X#) As Double
    HACos = Log(X + Root(X * X - 1))
End Function
Public Function HATan(ByVal X#) As Double
    HATan = Log((1 + X) / (1 - X)) / 2
End Function
Public Function HASec(ByVal X#) As Double
    HASec = Log((Root(-X * X + 1) + 1) / X)
End Function
Public Function HACosec(ByVal X#) As Double
    HACosec = Log((Sgn(X) * Root(X * X + 1) + 1) / X)
End Function
Public Function HACotan(ByVal X#) As Double
    HACotan = Log((X + 1) / (X - 1)) / 2
End Function
