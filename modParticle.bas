Attribute VB_Name = "modParticle"
Public Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long


Public Particle(0 To 2000) As typeP
Type typeP
    X As Variant
    Y As Variant
    
    Life As Integer
    Dead As Boolean
    
    Gravity As Variant
    
    Age As Integer
    ImageIndex As Integer
End Type

Public Const PSPEED = 0.01
