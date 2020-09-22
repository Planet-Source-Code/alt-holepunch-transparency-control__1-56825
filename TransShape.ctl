VERSION 5.00
Begin VB.UserControl TransShape 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackStyle       =   0  'Transparent
   CanGetFocus     =   0   'False
   ClientHeight    =   690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   765
   FillStyle       =   0  'Solid
   ForeColor       =   &H8000000F&
   ScaleHeight     =   46
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   51
   ToolboxBitmap   =   "TransShape.ctx":0000
End
Attribute VB_Name = "TransShape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**************************************************************************************************
' TransShape.ctl - Put transparent regions on your form using pre-defined shapes.
'
'  Copyright © 2004, Alan Tucker, All Rights Reserved
'  Contact alan_usa@hotmail.com for usage restrictions
'**************************************************************************************************
Option Explicit

'**************************************************************************************************
' TransShape.ctl Constant Declares
'**************************************************************************************************
Private Const RGN_AND = &H1&
Private Const RGN_OR = &H2&
Private Const RGN_XOR = &H3&
Private Const RGN_DIFF = &H4&
Private Const RGN_COPY = &H5&
Private Const NULLREGION = &H1&
Private Const SIMPLEREGION = &H2&
Private Const COMPLEXREGION = &H3&
Private Const WINDING = 2
Private Const SM_CYCAPTION = 4
Private Const SM_CYMENU = 15
Private Const SM_CXFRAME = 32
Private Const SM_CYFRAME = 33

'**************************************************************************************************
' TransShape.ctl Struct\Enum Declarations
'**************************************************************************************************
Public Enum eShape
     [Rectangular]
     [RoundRectangular]
     [Elliptical]
     [TriangleTop]
     [TriangleRight]
     [TriangleBottom]
     [TriangleLeft]
     [Star]
     [Diamond]
     [CustomPolygon]
End Enum ' eShape

Public Type POINTAPI
     X As Long
     Y As Long
End Type ' POINTAPI

Private Type RECT
     Left As Long
     Top As Long
     Right As Long
     Bottom As Long
End Type ' RECT

'**************************************************************************************************
' TransShape.ctl API Declarations
'**************************************************************************************************
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, _
     lpPoint As POINTAPI) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, _
     ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, _
     ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, _
     ByVal nPolyFillMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, _
     ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, _
     ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function Ellipse Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, _
     ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, _
     lpRect As RECT) As Long
Private Declare Function GetWindowRgn Lib "user32" (ByVal hwnd As Long, _
     ByVal hRgn As Long) As Long
Private Declare Function OffsetRgn Lib "gdi32" (ByVal hRgn As Long, ByVal X As Long, _
     ByVal Y As Long) As Long
Private Declare Function Polygon Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, _
     ByVal nCount As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, _
     ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, _
     ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, _
     ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, _
       ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
   
'**************************************************************************************************
' TransShape.ctl Default Property Variables
'**************************************************************************************************
Const m_def_Shape = 0

'**************************************************************************************************
' TransShape.ctl Property Variables
'**************************************************************************************************
Dim m_hDC As Long
Dim m_IsAmbient As Boolean
Dim m_Shape As eShape

'**************************************************************************************************
' TransShape.ctl Property Get\Let\Set
'**************************************************************************************************
Private Property Get IsAmbient() As Boolean
     IsAmbient = m_IsAmbient
End Property ' Get IsAmbient

Private Property Let IsAmbient(New_IsAmbient As Boolean)
     m_IsAmbient = New_IsAmbient
     With UserControl
          If Not (m_IsAmbient) Then
               .BackStyle = 1
          Else
               .BackStyle = 0
          End If
     End With
End Property ' Let IsAmbient

Public Property Get Shape() As eShape
     Shape = m_Shape
End Property ' Get Shape

Public Property Let Shape(New_Shape As eShape)
     m_Shape = New_Shape
     ' Draw selected shape
     DrawShape New_Shape
     PropertyChanged "Shape"
End Property ' Let Shape

'**************************************************************************************************
' TransShape.ctl Intrinsic Methods
'**************************************************************************************************
Private Sub UserControl_InitProperties()
     Shape = m_def_Shape
End Sub ' UserControl_InitProperties

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
     Shape = PropBag.ReadProperty("Shape", m_def_Shape)
     IsAmbient = Ambient.UserMode
End Sub ' UserControl_ReadProperties

Private Sub UserControl_Resize()
     ' redraw shape
     DrawShape Shape
End Sub ' UserControl_Resize

Private Sub UserControl_Show()
     IsAmbient = Ambient.UserMode
End Sub ' UserControl_Show

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
     PropBag.WriteProperty "Shape", Shape, m_def_Shape
     IsAmbient = Ambient.UserMode
End Sub ' UserControl_WriteProperties

'**************************************************************************************************
' TransShape.ctl Private Methods
'**************************************************************************************************
Private Sub DrawShape(shp As eShape)
     Dim lhRgnEx As Long
     Dim lhRgnExType As Long
     Dim hR2 As Long
     Dim hR1 As Long
     Dim hR3 As Long
     Dim hR4 As Long
     Dim lhRgn As Long
     Dim lRtn As Long
     Dim rc As RECT
     Dim pt(1 To 5) As POINTAPI
     With UserControl
          On Error Resume Next
          If UserControl.Parent.MDIChild Then
               ' Create region encompassing the screen with mdi apps
               hR1 = CreateRectRgn(0, 0, Screen.Width \ Screen.TwipsPerPixelX, _
                    Screen.Height \ Screen.TwipsPerPixelY)
          Else
               ' set up rectangle
               rc.Left = 0
               rc.Top = -((.Parent.Height - .Parent.ScaleHeight) \ Screen.TwipsPerPixelY)
               rc.Right = .Parent.Width
               rc.Bottom = .Parent.Height
               ' create region
               hR1 = CreateRectRgn(rc.Left, rc.Top, rc.Right, rc.Bottom)
          End If
          ' process selected shape
          Select Case shp
               Case 0, 1, 2 ' Rectangle, RoundRectangle, Ellipse
                    pt(1).X = 0
                    pt(1).Y = 0
                    pt(2).X = .ScaleWidth
                    pt(2).Y = .ScaleHeight
                    If shp = 0 Then
                         ' If running
                         If IsAmbient Then
                              ' Create our rectangle region
                              hR2 = CreateRectRgn(pt(1).X, pt(1).Y, pt(2).X, pt(2).Y)
                              ' Offset region so that it will be placed at the postion
                              ' of the ctl
                              OffsetRegion hR2
                              ' create an empty region to accept the combined region
                              lhRgn = CreateRectRgn(0, 0, 0, 0)
                         Else ' we're in design, only create the shape
                              Cls
                              lRtn = Rectangle(hDC, 0, 0, .ScaleWidth, .ScaleHeight)
                         End If
                    ElseIf shp = 1 Then
                         ' If running
                         If IsAmbient Then
                              ' Create our round rectangle region
                              hR2 = CreateRoundRectRgn(pt(1).X, pt(1).Y, pt(2).X, _
                                   pt(2).Y, 20, 20)
                              ' Offset region so that it will be placed at the postion
                              ' of the ctl
                              OffsetRegion hR2
                              ' create an empty region to accept the combined region
                              lhRgn = CreateRoundRectRgn(0, 0, 0, 0, 20, 20)
                         Else ' we're in design, only create the shape
                              Cls
                              lRtn = RoundRect(hDC, 0, 0, .ScaleWidth, .ScaleHeight, 20, 20)
                         End If
                    ElseIf shp = 2 Then
                         ' If running
                         If IsAmbient Then
                              ' Create our elliptic region
                              hR2 = CreateEllipticRgn(pt(1).X, pt(1).Y, pt(2).X, pt(2).Y)
                              ' Offset region so that it will be placed at the postion
                              ' of the ctl
                              OffsetRegion hR2
                              ' create an empty region to accept the combined region
                              lhRgn = CreateEllipticRgn(0, 0, 0, 0)
                         Else ' we're in design, only create the shape
                              Cls
                              lRtn = Ellipse(hDC, 0, 0, .ScaleWidth, .ScaleHeight)
                         End If
                    End If
               Case 3, 4, 5, 6 ' Triangles
                    If shp = 3 Then
                         pt(1).X = 0
                         pt(1).Y = .ScaleHeight
                         pt(2).X = .ScaleWidth \ 2
                         pt(2).Y = 0
                         pt(3).X = .ScaleWidth
                         pt(3).Y = .ScaleHeight
                    ElseIf shp = 4 Then
                         pt(1).X = 0
                         pt(1).Y = 0
                         pt(2).X = 0
                         pt(2).Y = .ScaleHeight - 1
                         pt(3).X = .ScaleWidth
                         pt(3).Y = .ScaleHeight \ 2
                    ElseIf shp = 5 Then
                         pt(1).X = 0
                         pt(1).Y = 0
                         pt(2).X = .ScaleWidth \ 2
                         pt(2).Y = .ScaleHeight
                         pt(3).X = .ScaleWidth
                         pt(3).Y = 0
                    ElseIf shp = 6 Then
                         pt(1).X = 0
                         pt(1).Y = .ScaleHeight \ 2
                         pt(2).X = .ScaleWidth
                         pt(2).Y = 0
                         pt(3).X = .ScaleWidth
                         pt(3).Y = .ScaleHeight
                    End If
                    ' If running
                    If IsAmbient Then
                         ' Create our polygon
                         hR2 = CreatePolygonRgn(pt(1), 3, 2)
                         ' Offset region so that it will be placed at the postion
                         ' of the ctl
                         OffsetRegion hR2
                         ' create an empty region to accept the combined region
                         lhRgn = CreatePolygonRgn(pt(1), 3, 2)
                    Else ' we're in design, only create the shape
                         Cls
                         lRtn = Polygon(hDC, pt(1), 3)
                    End If
               Case 7 ' Star
                    ' Plot star points
                    PlotStar .ScaleWidth, .ScaleHeight, pt
                    ' if running
                    If IsAmbient Then
                         ' create our star-shaped region
                         hR2 = CreatePolygonRgn(pt(1), 5, 2)
                         ' Offset region so that it will be placed at the position
                         ' of the ctl
                         OffsetRegion hR2
                         ' create an empty region to accept the combined region
                         lhRgn = CreatePolygonRgn(pt(1), 5, 2)
                    Else ' we're in design, only create the shape
                         Cls
                         lRtn = Polygon(.hDC, pt(1), 5)
                    End If
               Case 8 ' Diamond
                    pt(1).X = 0
                    pt(1).Y = .ScaleHeight \ 2
                    pt(2).X = .ScaleWidth \ 2
                    pt(2).Y = .ScaleHeight
                    pt(3).X = .ScaleWidth
                    pt(3).Y = .ScaleHeight \ 2
                    pt(4).X = .ScaleWidth \ 2
                    pt(4).Y = 0
                    ' If running
                    If IsAmbient Then
                         ' Create our polygon
                         hR2 = CreatePolygonRgn(pt(1), 4, 2)
                         ' Offset region so that it will be placed at the position
                         ' of the ctl
                         OffsetRegion hR2
                         ' create an empty region to accept the combined region
                         lhRgn = CreatePolygonRgn(pt(1), 4, 2)
                    Else ' we're in design, only create the shape
                         Cls
                         lRtn = Polygon(.hDC, pt(1), 4)
                    End If
               Case 9 ' Custom
                    Cls
                    DeleteObject hR1
                    Exit Sub
          End Select
     End With
     ' if running...clean up and set final region
     If IsAmbient Then
          ' combine our regions
          hR3 = CombineRgn(lhRgn, hR1, hR2, RGN_DIFF)
          ' see if a region has already been set
          If (GetExistingRegion(lhRgnEx, pt)) > False Then
               ' combine with existing region so we can continue
               ' to add shapes to the form.  Otherwise, we would
               ' get only one shape.
               hR4 = CombineRgn(lhRgn, lhRgn, lhRgnEx, RGN_AND)
               ' Delete the temporary pre-existing region object
               DeleteObject lhRgnEx
          End If
          ' Delete our other temporary regions
          DeleteObject hR1
          DeleteObject hR2
          DeleteObject hR3
          DeleteObject hR4
          ' Set the final region on our parent form
          SetWindowRgn UserControl.Parent.hwnd, lhRgn, True
     End If
End Sub ' DrawShape

Private Sub PlotStar(ByVal lWidth As Long, ByVal lHeight As Long, ByRef pt() As POINTAPI)
     Const PI = 3.14159265
     Dim ucX As Single
     Dim ucY As Single
     Dim ucWid As Single
     Dim ucHgt As Single
     Dim ucTheta As Single
     Dim ucdTheta As Single
     Dim lLoop As Long
     ' calc center between height and width
     ucX = (0 + lWidth) / 2
     ucY = (0 + lHeight) / 2
     ' get absolute values
     ucWid = Abs(0 - lWidth) / 2
     ucHgt = Abs(0 - lHeight) / 2
     ' Calculate and connect the star's points.
     ucTheta = 90 * PI / 180
     ucdTheta = 2 * 72 * PI / 180
     ' set current X and Y
     UserControl.CurrentX = ucX + ucWid * Cos(ucTheta)
     UserControl.CurrentY = ucY - ucHgt * Sin(ucTheta)
     ' Loop, plot, and store star points
     For lLoop = 1 To 5
          ucTheta = ucTheta + ucdTheta
          pt(lLoop).X = ucX + ucWid * Cos(ucTheta)
          pt(lLoop).Y = ucY - ucHgt * Sin(ucTheta)
     Next
End Sub ' PlotStar

Private Function GetExistingRegion(ByRef hRgnEx As Long, ByRef pt() As POINTAPI) As Long
     Dim hRgnExType As Long
     Select Case m_Shape
          Case 0, 1, 2
               hRgnEx = CreatePolygonRgn(pt(1), 2, 2)
          Case 3, 4, 5, 6
               hRgnEx = CreatePolygonRgn(pt(1), 3, 2)
          Case 7
               ' see if the region has already been set
               hRgnEx = CreatePolygonRgn(pt(1), 5, 2)
          Case 8
               ' see if the region has already been set
               hRgnEx = CreatePolygonRgn(pt(1), 4, 2)
          Case 9
               hRgnEx = CreatePolygonRgn(pt(1), 8, 2)
     End Select
     ' determine complex, simple, null, or error
     hRgnExType = GetWindowRgn(UserControl.Parent.hwnd, hRgnEx)
     ' If 2 or better, we have a defined region
     If hRgnExType <= 1 Then
          ' else a region has not been defined yet so
          ' delete our temp region
          DeleteObject hRgnEx
          ' lose the region handle
          hRgnEx = False
          ' Empty array
          Erase pt
     Else
          ' set to other than 0
          GetExistingRegion = 1
     End If
End Function ' GetExistingRegion

Private Sub OffsetRegion(ByVal lRgn As Long)
     Dim lWtDiff As Long
     Dim lHtDiff As Long
     Dim lOffsetX As Long
     Dim lOffsetY As Long
     Dim lRtn As Long
     Dim pt As POINTAPI
     With UserControl
          If Not (UserControl.Parent.MDIChild) Then
               ' determine how much real estate of the width is border and divide
               ' it by two since half is left and half is right.  Necessary to account
               ' for the border when plotting the offset on the x axis.
               lWtDiff = ((.Parent.Width - .Parent.ScaleWidth) \ Screen.TwipsPerPixelX) \ 2
               ' determine where our control is sited on the x axis
               lOffsetX = (Extender.Left \ Screen.TwipsPerPixelX) + _
                    (.Parent.Width - .Parent.ScaleWidth) \ Screen.TwipsPerPixelX
               ' Subtract the border real estate
               lOffsetX = lOffsetX - lWtDiff
               ' determine real estate of the border and titlebar so we can account
               ' for it when offsetting the Y axis
               lHtDiff = (.Parent.Height - .Parent.ScaleHeight) \ Screen.TwipsPerPixelY
               ' find out control position on the Y axis
               lOffsetY = (Extender.Top \ Screen.TwipsPerPixelY) + _
                    ((.Parent.Height - .Parent.ScaleHeight) \ Screen.TwipsPerPixelY)
               ' we have the Y offset but need to subtract half the border.
               lOffsetY = lOffsetY - lWtDiff
               ' now use the API to offset the region
               lRtn = OffsetRgn(lRgn, lOffsetX, lOffsetY)
          Else
               pt.X = (Extender.Left \ Screen.TwipsPerPixelX) + GetSystemMetrics(SM_CXFRAME)
               pt.Y = (Extender.Top \ Screen.TwipsPerPixelY) + GetSystemMetrics(SM_CYCAPTION) + _
                    GetSystemMetrics(SM_CYFRAME)
               lRtn = OffsetRgn(lRgn, pt.X, pt.Y)
          End If
     End With
End Sub ' OffsetRegion

'**************************************************************************************************
' TransShape.ctl Public Methods
'**************************************************************************************************
Public Function DrawCustomShape(xPts() As Variant, yPts() As Variant, ByVal lpntCnt As Long) As Long
     Dim lRtn As Long
     Dim lhRgnEx As Long
     Dim lhRgnExType As Long
     Dim hR2 As Long
     Dim hR1 As Long
     Dim hR3 As Long
     Dim hR4 As Long
     Dim lhRgn As Long
     Dim rc As RECT
     Dim pt() As POINTAPI
     Dim lLoop As Long
     ' are xpts and ypts an array
     If IsArray(xPts) And IsArray(yPts) Then
          ' convert points to pointapi
          ReDim pt(1 To lpntCnt)
          For lLoop = 1 To lpntCnt
               pt(lLoop).X = xPts(lLoop - 1)
               pt(lLoop).Y = yPts(lLoop - 1)
          Next
          If IsAmbient Then
               If UserControl.Parent.MDIChild Then
                    ' Create region encompassing the screen with mdi apps
                    hR1 = CreateRectRgn(0, 0, Screen.Width \ Screen.TwipsPerPixelX, _
                         Screen.Height \ Screen.TwipsPerPixelY)
               Else
                    With UserControl
                    ' set up rectangle
                         rc.Left = 0
                         rc.Top = -((.Parent.Height - .Parent.ScaleHeight) \ Screen.TwipsPerPixelY)
                         rc.Right = .Parent.Width
                         rc.Bottom = .Parent.Height
                    End With
                    ' create region
                    hR1 = CreateRectRgn(rc.Left, rc.Top, rc.Right, rc.Bottom)
               End If
               ' Create our polygonal region
               hR2 = CreatePolygonRgn(pt(1), lpntCnt, 2)
               ' Offset region so that it will be placed at the postion
               ' of the ctl
               OffsetRegion hR2
               ' create an empty region to accept the combined region
               lhRgn = CreatePolygonRgn(pt(1), lpntCnt, 2)
               ' combine our regions
               hR3 = CombineRgn(lhRgn, hR1, hR2, RGN_DIFF)
               ' see if a region has already been set
               If (GetExistingRegion(lhRgnEx, pt)) > False Then
                    ' combine with existing region so we can continue
                    ' to add shapes to the form.  Otherwise, we would
                    ' get only one shape.
                    hR4 = CombineRgn(lhRgn, lhRgn, lhRgnEx, RGN_AND)
                    ' Delete the temporary pre-existing region object
                    DeleteObject lhRgnEx
               End If
               ' Delete our other temporary regions
               DeleteObject hR1
               DeleteObject hR2
               DeleteObject hR3
               DeleteObject hR4
               ' Set the final region on our parent form
               SetWindowRgn UserControl.Parent.hwnd, lhRgn, True
          End If
     End If
End Function ' DrawCustomShape
