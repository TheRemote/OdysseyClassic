Attribute VB_Name = "modWeather"
Public Declare Function InitSnow Lib "odysseydll" (ByVal TickCount As Long) As Long
Public Declare Function Snow16 Lib "odysseydll" (ByRef Surface As Any, ByVal TickCount As Long) As Long
Public Declare Function Snow32 Lib "odysseydll" (ByRef Surface As Any, ByVal TickCount As Long) As Long
Public Declare Function InitRain Lib "odysseydll" (ByVal TickCount As Long) As Long
Public Declare Function Rain16 Lib "odysseydll" (ByRef Surface As Any, ByVal TickCount As Long) As Long
Public Declare Function Rain32 Lib "odysseydll" (ByRef Surface As Any, ByVal TickCount As Long) As Long
