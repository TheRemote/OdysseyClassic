Attribute VB_Name = "imagefx"
Option Explicit
'**********************************
'* CODE BY: PATRICK MOORE (ZELDA) *
'* Feel free to re-distribute or  *
'* Use in your own projects.      *
'* Giving credit to me would be   *
'* nice :)                        *
'*                                *
'* Please vote for me if you find *
'* this code useful :]   -Patrick *
'**********************************
'
'PS: Please look for more submissions to PSC by me
'    shortly.  I've recently been working on a lot
'    :))  All my submissions are under author name
'    "Patrick Moore (Zelda)"

Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public rRed As Long, rBlue As Long, rGreen As Long

Private Type Pixel
    Pix As Long
    Top As Long
    Left As Long
    Right As Long
    Bottom As Long
End Type

Private Type RGBVals
    Red As Long
    Blue As Long
    Green As Long
End Type

Function RGBfromLONG(LongCol As Long)
'// This Function by Zel
' Get The Red, Blue And Green Values Of A Colour From The Long Value
Dim Blue As Double, Green As Double, Red As Double, GreenS As Double, BlueS As Double
Blue = Fix((LongCol / 256) / 256)
Green = Fix((LongCol - ((Blue * 256) * 256)) / 256)
Red = Fix(LongCol - ((Blue * 256) * 256) - (Green * 256))
rRed = Red: rBlue = Blue: rGreen = Green
End Function

Function GetRandomNumber(Upper As Integer, Lower As Integer) As Integer
'Get a random number
Randomize
GetRandomNumber = Int((Upper) * Rnd)
End Function

Public Sub Noise(PicBox, Intensity As Integer)
'Add noise to a picture
Dim X As Long, W As Integer, H As Integer, Num As Integer, Num2 As Integer
PicBox.ScaleMode = 3
W = PicBox.ScaleWidth
H = PicBox.ScaleHeight
For X = 1 To Intensity * 50
    Randomize
    Num = Int(Rnd * W - 1) + 1
    Randomize
    Num2 = Int(Rnd * H - 1) + 1
    SetPixel PicBox.hDC, Num, Num2, GetPixel(PicBox.hDC, Num2, Num)
Next X
End Sub
Public Sub Pixelate(PicBox, size As Integer)
'Pixelate a picture
Dim W As Long, H As Long, NumC As Integer
Dim Color As Long, CA As Integer
Dim C(1 To 100) As Long, S As Integer
Dim G As Long, R As Long, B As Long
PicBox.ScaleMode = 3
For H = 0 To PicBox.ScaleHeight - 2 Step size
    For W = 0 To PicBox.ScaleWidth - 2 Step size
        NumC = 1
        For S = 1 To size
            C(NumC) = GetPixel(PicBox.hDC, W, H)
            NumC = NumC + 1
            C(NumC) = GetPixel(PicBox.hDC, W + S, H)
            NumC = NumC + 1
            C(NumC) = GetPixel(PicBox.hDC, W + S, H + S)
            NumC = NumC + 1
            C(NumC) = GetPixel(PicBox.hDC, W, H + S)
            NumC = NumC + 1
        Next S
        For CA = 1 To NumC
            RGBfromLONG C(CA)
            G = G + rGreen
            R = R + rRed
            B = B + rBlue
        Next CA
        R = R / NumC
        G = G / NumC
        B = B / NumC
        Color = RGB(R, G, B)
        
        For S = 0 To size
            PicBox.Line (W + S, H)-(W + S, H + size), Color, BF
        Next S
    Next W
Next H
End Sub
Public Sub Lighten(Percent As Integer, PicBox)
'Lighten a picture
Dim newVal As Integer, H As Long, W As Long, K As Integer
Dim C As Long
Dim opRed As Long, opBlue As Long, opGreen As Long

newVal = Percent * 5
PicBox.ScaleMode = 3

For W = 0 To PicBox.ScaleWidth
    For H = 0 To PicBox.ScaleHeight
        C = GetPixel(PicBox.hDC, W, H)
        RGBfromLONG C
        opRed = rRed
        opGreen = rGreen
        opBlue = rBlue
        rRed = rRed + newVal
        If rRed > -1 And rRed < 256 Then opRed = rRed
        
        rGreen = rGreen + newVal
        If rGreen > -1 And rGreen < 256 Then opGreen = rGreen
        rBlue = rBlue + newVal
        If rBlue > -1 And rBlue < 256 Then opBlue = rBlue
        If rRed <> 1000 Then
           C = RGB(opRed, opGreen, opBlue)
           SetPixel PicBox.hDC, W, H, C
        End If
    Next H
Next W
End Sub


Public Sub Darken(Percent As Integer, PicBox)
'Darken a picture
Dim newVal As Integer, H As Long, W As Long, K As Integer
Dim C As Long
Dim icRed As Long, icBlue As Long, icGreen As Long
Dim opRed As Long, opBlue As Long, opGreen As Long

newVal = Percent * -5
PicBox.ScaleMode = 3

For H = 0 To PicBox.ScaleHeight
    For W = 0 To PicBox.ScaleWidth
        C = GetPixel(PicBox.hDC, W, H)
        RGBfromLONG C
        opRed = rRed
        opBlue = rBlue
        opGreen = rGreen
        rRed = rRed + newVal
        If rRed > -1 And icRed < 256 Then opRed = rRed
        
        rGreen = rGreen + newVal
        If rGreen > -1 And rGreen < 256 Then opGreen = rGreen
        rBlue = rBlue + newVal
        If rBlue > -1 And rBlue < 256 Then opBlue = rBlue
        If rRed <> 1000 Then
            If opRed < 0 Then opRed = 0
            If opGreen < 0 Then opGreen = 0
            If opBlue < 0 Then opBlue = 0
           C = RGB(opRed, opGreen, opBlue)
           SetPixel PicBox.hDC, W, H, C
        End If
    Next W
Next H
End Sub


Public Sub GrayScale(PicBox)
'Turn a color image to greyscale
Dim AveCol As Integer, A As Integer
Dim Y As Long, X As Long

PicBox.ScaleMode = 3
For Y = 0 To PicBox.ScaleHeight
    For X = 0 To PicBox.ScaleWidth
        AveCol = 0
        A = 0
        RGBfromLONG GetPixel(PicBox.hDC, X, Y)
        AveCol = AveCol + rGreen: A = A + 1
        If AveCol <= 0 Then AveCol = 0
        AveCol = (AveCol / A)
        SetPixel PicBox.hDC, X, Y, RGB(AveCol, AveCol, AveCol)
    Next X
Next Y
End Sub


Function LightenPixel(pixelLong As Long, Percent As Integer)
'Lighten only one pixel
Dim newVal As Integer, C As Long, opRed As Long, opGreen As Long, opBlue As Long
newVal = Percent * 5
C = pixelLong
RGBfromLONG C
opRed = rRed
opGreen = rGreen
opBlue = rBlue
rRed = rRed + newVal
If rRed > -1 And rRed < 256 Then opRed = rRed

rGreen = rGreen + newVal
If rGreen > -1 And rGreen < 256 Then opGreen = rGreen
rBlue = rBlue + newVal
If rBlue > -1 And rBlue < 256 Then opBlue = rBlue
If rRed <> 1000 Then
    C = RGB(opRed, opGreen, opBlue)
    LightenPixel = C
End If
End Function


Function DarkenPixel(pixelLong As Long, Percent As Integer) As Long
'Darken only one pixel
Dim newVal As Integer, C As Long, opRed As Long, opGreen As Long, opBlue As Long
newVal = Percent * -5
C = pixelLong
RGBfromLONG C
opRed = rRed
opGreen = rGreen
opBlue = rBlue
rRed = rRed + newVal
If rRed > -1 And rRed < 256 Then opRed = rRed

rGreen = rGreen + newVal
If rGreen > -1 And rGreen < 256 Then opGreen = rGreen
rBlue = rBlue + newVal
If rBlue > -1 And rBlue < 256 Then opBlue = rBlue
If rRed <> 1000 Then
    If opRed < 0 Then opRed = 0
    If opGreen < 0 Then opGreen = 0
    If opBlue < 0 Then opBlue = 0
    C = RGB(opRed, opGreen, opBlue)
    DarkenPixel = C
End If
End Function


Public Sub Blur(PicBox As PictureBox, Intensity As Integer)
'Blur a picture
Dim W As Long, H As Long, NumC As Long
Dim Color As Long, CA As Long, size As Integer
Dim C(1 To 100) As Long, S As Long, I As Long
Dim G As Long, R As Long, B As Long
PicBox.ScaleMode = 3
size = 1

For I = 1 To Intensity
    For W = 0 To PicBox.ScaleWidth - 2 Step size
        For H = 0 To PicBox.ScaleHeight - 2 Step size
            NumC = 1
            For S = 1 To size
                C(NumC) = GetPixel(PicBox.hDC, W, H)
                NumC = NumC + 1
                C(NumC) = GetPixel(PicBox.hDC, W + S, H)
                NumC = NumC + 1
                C(NumC) = GetPixel(PicBox.hDC, W + S, H + S)
                NumC = NumC + 1
                C(NumC) = GetPixel(PicBox.hDC, W, H + S)
                NumC = NumC + 1
            Next S
            For CA = 1 To NumC
                RGBfromLONG C(CA)
                G = G + rGreen
                R = R + rRed
                B = B + rBlue
            Next CA
            If G > 0 And R > 0 And B > 0 Then
                R = R / NumC
                G = G / NumC
                B = B / NumC
            Else
                R = 0
                G = 0
                B = 0
            End If
            Color = RGB(R, G, B)
            
            For S = 0 To size
                PicBox.Line (W + S, H)-(W + S, H + size), Color, BF
            Next S
        Next H
    Next W
Next I
End Sub

Function InvertPixel(colorLong As Long) As Long
'Invert the color of a pixel
Dim opRed As Long, opGreen As Long, opBlue As Long
RGBfromLONG colorLong

InvertPixel = RGB(255 - rRed, 255 - rGreen, 255 - rBlue)
End Function
Public Sub Invert(PicBox)
'Invert the image of a picturebox
Dim newVal As Integer, H As Long, W As Long, K As Integer
Dim C As Long
Dim opRed As Long, opBlue As Long, opGreen As Long

PicBox.ScaleMode = 3

For H = 0 To PicBox.ScaleHeight
    For W = 0 To PicBox.ScaleWidth

        C = GetPixel(PicBox.hDC, W, H)
        RGBfromLONG C
        opRed = 255 - rRed
        opGreen = 255 - rGreen
        opBlue = 255 - rBlue
        C = RGB(opRed, opGreen, opBlue)
        SetPixel PicBox.hDC, W, H, C
    Next W
Next H
End Sub

Public Sub FlipHorizontal(PicBox As PictureBox)
'Flip the picturebox (like when you look into a
'mirror).
'
'BTW: I know that BitBlt does a MUCH easier job
'with this, but I was experimenting with Get/Set Pixel
'and thought beginners or people not used to Get/Set Pixel
'would find it useful :]

Dim W As Long, H As Long, Num As Integer
Dim OldColor(0 To 1000, 0 To 1000) As Long, cColor As Long
PicBox.ScaleMode = 3
Num = PicBox.ScaleWidth / 2
For W = PicBox.ScaleWidth / 2 To PicBox.ScaleWidth
    For H = 0 To PicBox.ScaleHeight
        cColor = GetPixel(PicBox.hDC, Num, H)
        OldColor(W - (PicBox.ScaleWidth / 2), H) = GetPixel(PicBox.hDC, W, H)
        SetPixel PicBox.hDC, W, H, cColor
    Next H
    Num = Num - 1
    DoEvents
Next W


Num = PicBox.ScaleWidth / 2
For W = 0 To PicBox.ScaleWidth / 2
    For H = 0 To PicBox.ScaleHeight
        SetPixel PicBox.hDC, Num, H, OldColor(W, H)
    Next H
    Num = Num - 1
    DoEvents
Next W
End Sub

Public Sub Colorize(BaseColor As Long, PicBox As PictureBox)
'Colorize a picture
Dim newVal As Integer, H As Long, W As Long, K As Integer
Dim C As Long
Dim icRed As Long, icBlue As Long, icGreen As Long
Dim opRed As Long, opBlue As Long, opGreen As Long

Dim origRed As Long, origBlue As Long, origGreen As Long
RGBfromLONG BaseColor
origRed = rRed
origBlue = rBlue
origGreen = rGreen

PicBox.ScaleMode = 3
GrayScale PicBox

For H = 0 To PicBox.ScaleHeight
    For W = 0 To PicBox.ScaleWidth
        C = GetPixel(PicBox.hDC, W, H)
        RGBfromLONG C
        opRed = rRed
        opBlue = rBlue
        opGreen = rGreen
        
        
        If rRed > 0 Then rRed = origRed * (rRed / 255)
        If rRed > -1 And icRed < 256 Then opRed = rRed
        
        If rGreen > 0 Then rGreen = origGreen * (rGreen / 255)
        If rGreen > -1 And rGreen < 256 Then opGreen = rGreen
        
        If rBlue > 0 Then rBlue = origBlue * (rBlue / 255)
        If rBlue > -1 And rBlue < 256 Then opBlue = rBlue

        If rRed <> 1000 Then
            If opRed < 0 Then opRed = 0
            If opGreen < 0 Then opGreen = 0
            If opBlue < 0 Then opBlue = 0
           C = RGB(opRed, opGreen, opBlue)
           SetPixel PicBox.hDC, W, H, C
        End If
    Next W
Next H
End Sub

Public Sub Merge(picBox1 As PictureBox, picBox2 As PictureBox, OutputPic As PictureBox)
Dim H As Long, W As Long
Dim Col1 As Long, Col2 As Long
Dim MergeColR As Long, MergeColG As Long, MergeColB As Long

picBox1.ScaleMode = 3
picBox2.ScaleMode = 3
OutputPic.ScaleMode = 3
For H = 0 To picBox1.ScaleHeight
    For W = 0 To picBox1.ScaleWidth
        Col1 = GetPixel(picBox1.hDC, W, H)
        Col2 = GetPixel(picBox2.hDC, W, H)
        RGBfromLONG Col1
        If rRed < 0 Then rRed = 0
        If rGreen < 0 Then rGreen = 0
        If rBlue < 0 Then rBlue = 0
        MergeColR = rRed
        MergeColG = rGreen
        MergeColB = rBlue
        RGBfromLONG Col2
        If rRed < 0 Then rRed = 0
        If rGreen < 0 Then rGreen = 0
        If rBlue < 0 Then rBlue = 0
        MergeColR = (MergeColR + rRed) / 2
        MergeColG = (MergeColG + rGreen) / 2
        MergeColB = (MergeColB + rBlue) / 2
        SetPixel OutputPic.hDC, W, H, RGB(MergeColR, MergeColG, MergeColB)
    Next W
Next H
End Sub


Public Sub Gradient(PicBox, StartColor As Long, EndColor As Long)
Dim W As Long, Color As Long
Dim GradR As Integer, GradB As Integer, GradG As Integer
Dim B1 As Integer, B2 As Integer
Dim G1 As Integer, G2 As Integer
Dim R1 As Integer, R2 As Integer

'Determine Red, Green, and Blue values
'for the first color
RGBfromLONG StartColor
B1 = rBlue
G1 = rGreen
R1 = rRed

'Determine Red, Green, and Blue values
'for the last color
RGBfromLONG EndColor
B2 = rBlue
G2 = rGreen
R2 = rRed

PicBox.ScaleMode = 3
For W = 0 To PicBox.ScaleWidth
    GradR = ((R2 - R1) / PicBox.ScaleWidth * W) + R1
    GradG = ((G2 - G1) / PicBox.ScaleWidth * W) + G1
    GradB = ((B2 - B1) / PicBox.ScaleWidth * W) + B1
    Color = RGB(GradR, GradG, GradB)
    PicBox.Line (W, 0)-(W, PicBox.ScaleHeight), Color, B
Next W
End Sub



Public Sub GradientMerge(PicBox, StartColor As Long, EndColor As Long)
Dim W As Long, Color As Long, H As Long
Dim GradR As Integer, GradB As Integer, GradG As Integer
Dim B1 As Integer, B2 As Integer
Dim G1 As Integer, G2 As Integer
Dim R1 As Integer, R2 As Integer

'Determine Red, Green, and Blue values
'for the first color
RGBfromLONG StartColor
B1 = rBlue
G1 = rGreen
R1 = rRed

'Determine Red, Green, and Blue values
'for the last color
RGBfromLONG EndColor
B2 = rBlue
G2 = rGreen
R2 = rRed

GrayScale PicBox
PicBox.ScaleMode = 3
For W = 0 To PicBox.ScaleWidth
    For H = 0 To PicBox.ScaleHeight
        GradR = ((R2 - R1) / PicBox.ScaleWidth * W) + R1
        GradG = ((G2 - G1) / PicBox.ScaleWidth * W) + G1
        GradB = ((B2 - B1) / PicBox.ScaleWidth * W) + B1
        
        RGBfromLONG GetPixel(PicBox.hDC, W, H)
        Color = RGB((GradR + rRed) / 2, (GradG + rGreen) / 2, (GradB + rBlue) / 2)
        SetPixel PicBox.hDC, W, H, Color
    Next H
Next W
End Sub


Public Sub ReplaceColor(FindColor As Long, ReplaceColor As Long, PicBox As PictureBox)
'Replace a color in one picturebox with another
Dim newVal As Integer, H As Long, W As Long, K As Integer
Dim C As Long

PicBox.ScaleMode = 3

For H = 0 To PicBox.ScaleHeight
    For W = 0 To PicBox.ScaleWidth
        C = GetPixel(PicBox.hDC, W, H)
        If C = FindColor Then C = ReplaceColor
        SetPixel PicBox.hDC, W, H, C
    Next W
Next H
End Sub
Sub Text_3DGradient(PicBox As PictureBox, Caption As String, Forecolor As Long, BackColor As Long, Optional Distance As Integer = 10, Optional Left As Integer = 11, Optional Top As Integer = 11)
Dim B1 As Long, G1 As Long, R1 As Long
Dim B2 As Long, G2 As Long, R2 As Long
Dim GradB As Long, GradG As Long, GradR As Long
Dim W As Long, Color As Long

'Add text, with a drop shadow, to a picturebox
If PicBox.ScaleMode <> 3 Then PicBox.ScaleMode = 3

'Determine Red, Green, and Blue values
'for the first color
RGBfromLONG BackColor
B1 = rBlue
G1 = rGreen
R1 = rRed

'Determine Red, Green, and Blue values
'for the last color
RGBfromLONG Forecolor
B2 = rBlue
G2 = rGreen
R2 = rRed

PicBox.ScaleMode = 3

For W = 0 To Distance
    GradR = ((R2 - R1) / Distance * W) + R1
    GradG = ((G2 - G1) / Distance * W) + G1
    GradB = ((B2 - B1) / Distance * W) + B1
    Color = RGB(GradR, GradG, GradB)
    PicBox.CurrentX = Left + W
    PicBox.CurrentY = Top + W
    PicBox.Forecolor = Color
    PicBox.Print Caption
Next W
End Sub

Sub Text_DropShadow(PicBox As PictureBox, Caption As String, TextColor As Long, Optional ShadowColor As Long = "&H00C0C0C0", Optional Intensity As Integer = 2, Optional Distance As Integer = 4, Optional Left As Integer = 11, Optional Top As Integer = 11)
'Add text, with a drop shadow, to a picturebox
If PicBox.ScaleMode <> 3 Then PicBox.ScaleMode = 3

PicBox.Forecolor = PicBox.BackColor
PicBox.CurrentX = Left + Distance: PicBox.CurrentY = Top + Distance

PicBox.Forecolor = ShadowColor
PicBox.Print Caption
Blur PicBox, 4 - Intensity

PicBox.Forecolor = PicBox.BackColor
PicBox.CurrentX = Left: PicBox.CurrentY = Top
PicBox.Forecolor = TextColor
PicBox.Print Caption
End Sub

Sub Text_BorderOnly(pBox As PictureBox, Caption As String)
Dim X As Long, Y As Long, Pi As Pixel, Color As Long
Color = pBox.BackColor
pBox.BackColor = vbRed
pBox.Forecolor = vbBlack
pBox.Print Caption
For X = 1 To pBox.Width - 1
    For Y = 1 To pBox.Height - 1
        Pi.Pix = GetPixel(pBox.hDC, X, Y)
        Pi.Bottom = GetPixel(pBox.hDC, X, Y + 1)
        Pi.Left = GetPixel(pBox.hDC, X - 1, Y)
        Pi.Right = GetPixel(pBox.hDC, X + 1, Y)
        Pi.Top = GetPixel(pBox.hDC, X, Y - 1)
        If Pi.Left <> vbRed And Pi.Right <> vbRed And Pi.Bottom <> vbRed Then
            If Pi.Top <> vbRed Then SetPixel pBox.hDC, X, Y, vbWhite
        End If
    Next Y
Next X

ReplaceColor vbRed, Color, pBox
End Sub


Sub Text_OuterGlow(PicBox As PictureBox, Caption As String, TextColor As Long, Optional GlowColor As Long = "16777215", Optional Distance As Integer = 4, Optional Left As Integer = 11, Optional Top As Integer = 11)
'Add text, with an outer glow, to a picturebox

'// SETUP PICTUREBOX
If PicBox.ScaleMode <> 3 Then PicBox.ScaleMode = 3
PicBox.Forecolor = PicBox.BackColor

'// SOUTHERN PORTION OF GLOW
PicBox.CurrentX = Left: PicBox.CurrentY = Top + Distance
PicBox.Forecolor = GlowColor
PicBox.Print Caption

'// NORTHERN PORTION OF GLOW
PicBox.CurrentX = Left: PicBox.CurrentY = Top - Distance
PicBox.Forecolor = GlowColor
PicBox.Print Caption

'// EASTERN PORTION OF GLOW
PicBox.CurrentX = Left + Distance: PicBox.CurrentY = Top
PicBox.Forecolor = GlowColor
PicBox.Print Caption

'// WESTERN PORTION OF GLOW
PicBox.CurrentX = Left - Distance: PicBox.CurrentY = Top
PicBox.Forecolor = GlowColor
PicBox.Print Caption

'// APPLY BLUR TO SHADOW
Blur PicBox, 2

'// PRINT REGULAR TEXT
PicBox.Forecolor = PicBox.BackColor
PicBox.CurrentX = Left: PicBox.CurrentY = Top
PicBox.Forecolor = TextColor
PicBox.Print Caption
End Sub

Sub Stroke(PicBox As PictureBox, Color As Long)
Dim X As Long, Y As Long, PixColors As Pixel
Dim BGColor As Long, W As Long
BGColor = PicBox.BackColor
PicBox.ScaleMode = 3
For X = 1 To PicBox.ScaleWidth - 2
    For Y = 1 To PicBox.ScaleHeight - 2
        PixColors.Bottom = GetPixel(PicBox.hDC, X, Y + 1)
        PixColors.Left = GetPixel(PicBox.hDC, X - 1, Y)
        PixColors.Right = GetPixel(PicBox.hDC, X + 1, Y)
        PixColors.Top = GetPixel(PicBox.hDC, X, Y - 1)
        PixColors.Pix = GetPixel(PicBox.hDC, X, Y)
        If PixColors.Pix <> BGColor Then GoTo Continue
        If PixColors.Bottom <> Color And PixColors.Bottom <> BGColor Then SetPixel PicBox.hDC, X, Y - W, Color: GoTo Continue
        If PixColors.Left <> Color And PixColors.Left <> BGColor Then SetPixel PicBox.hDC, X + W, Y, Color: GoTo Continue
        If PixColors.Top <> Color And PixColors.Top <> BGColor Then SetPixel PicBox.hDC, X, Y + W, Color: GoTo Continue
        If PixColors.Right <> Color And PixColors.Right <> BGColor Then SetPixel PicBox.hDC, X - W, Y, Color
Continue:
    Next Y
Next X
End Sub

Public Sub TVScanLines(PicBox As PictureBox)
'Colorize a picture
Dim H As Long, W As Long, C As Long
PicBox.ScaleMode = 3
'GrayScale PicBox

For H = 0 To PicBox.ScaleHeight
    For W = 0 To PicBox.ScaleWidth
        C = GetPixel(PicBox.hDC, W, H)
        If IsOdd(H) = True Then
            C = DarkenPixel(C, 7)
        Else
            C = LightenPixel(C, 4)
        End If
        SetPixel PicBox.hDC, W, H, C
    Next W
Next H
End Sub

Function IsOdd(Num) As Boolean
Dim RNum As String
RNum = Right(Num, 1)
If RNum = "1" Or RNum = "3" Or RNum = "5" Or RNum = "7" Or RNum = "9" Then
    IsOdd = True
Else
    IsOdd = False
End If
End Function
