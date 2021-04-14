VERSION 5.00
Begin VB.UserControl uAOProgress 
   BackStyle       =   0  'Transparent
   ClientHeight    =   945
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2295
   ScaleHeight     =   945
   ScaleWidth      =   2295
   ToolboxBitmap   =   "uAOProgress.ctx":0000
   Begin VB.Timer tDanger 
      Interval        =   250
      Left            =   1200
      Top             =   360
   End
   Begin VB.Timer tTimer 
      Interval        =   10
      Left            =   1680
      Top             =   360
   End
   Begin VB.Label lblStat 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "999/999"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   615
   End
   Begin VB.Label lblShadowText 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "999/999"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   480
      TabIndex        =   1
      Top             =   0
      Width           =   615
   End
   Begin VB.Shape shpStat 
      BorderColor     =   &H000000FF&
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   300
      Left            =   0
      Top             =   0
      Width           =   1365
   End
   Begin VB.Shape shpAdd 
      BorderColor     =   &H000000FF&
      FillColor       =   &H00808000&
      FillStyle       =   0  'Solid
      Height          =   300
      Left            =   0
      Top             =   0
      Width           =   765
   End
   Begin VB.Shape shpSub 
      BorderColor     =   &H000000FF&
      FillColor       =   &H00808000&
      FillStyle       =   0  'Solid
      Height          =   300
      Left            =   0
      Top             =   0
      Width           =   1605
   End
   Begin VB.Shape shpBack 
      BorderColor     =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   300
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2085
   End
End
Attribute VB_Name = "uAOProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%                                                     %
'%                   AO PROGRESS v1.5                  %
'%               Copyright © 2021 by ^[GS]^            %
'%                    www.GS-ZONE.org                  %
'%                                                     %
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%  Este control permite realizar barras de            %
'%  progreso facilmente.                               %
'%                                                     %
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%  Changelog:                                         %
'%   13/04/2021 - Se agrego el color danger y se       %
'%                mejoraron las animaciones. (^[GS]^)  %
'%   25/04/2013 - Se agrego la posibilidad de usar     %
'%                un color de fondo solido. (^[GS]^)   %
'%   03/09/2012 - Mejora de rendimiento.               %
'%                Se agrego el % al mantener el        %
'%                sobre el valor. (^[GS]^)             %
'%   25/08/2012 - Se finalizo una primera versión,     %
'%                sencilla, son Shapes y animación     %
'%                de cambio de valor. (^[GS]^)         %
'%   23/07/2012 - Se inicio el proyecto. (^[GS]^)      %
'%                                                     %
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

Option Explicit

Private iMax As Long
Private iMin As Long
Private iValue As Long
Private iNewValue As Long
Private bShowText As Boolean
Private sCustomText As String
Private bAnimate As Boolean
Private bEnabled As Boolean
Private bUseBackground As Boolean
Private lForeColor As Long
Private lBackColor As Long
Private lBackAddColor As Long
Private lBackSubColor As Long
Private lBorderColor As Long
Private lBackgroundColor As Long
Private fTextFont As Font
Private iMinDanger As Long
Private lBackDangerColor As Long
Private lBackgroundDangerColor As Long
Private lShadowTextColor As Long

Private MouseOverText As String
Private bAnimating As Boolean

Private Sub DrawStat()
'*************************************************
'Author: ^[GS]^
'Last modified: 13/04/2021
'*************************************************

On Error Resume Next
    
    If bEnabled = False Then Exit Sub
    
    If LenB(MouseOverText) <> 0 Then MouseOverText = vbNullString
    
    If iNewValue < MinDanger Then
        tDanger.Enabled = True
    Else
        tDanger.Enabled = False
        shpStat.FillColor = lBackColor
        shpBack.FillColor = lBackgroundColor
        lblStat.ForeColor = lForeColor
        lblShadowText.ForeColor = lShadowTextColor
    End If
    
    If bAnimate = False Then
        iNewValue = iValue
        If LenB(sCustomText) = 0 Then
            lblStat.Caption = iNewValue & "/" & iMax
            lblShadowText.Caption = lblStat.Caption
        End If
        shpAdd.Visible = False
        shpSub.Visible = False
        shpStat.Width = (((iNewValue / 100) / (iMax / 100)) * UserControl.Width)
    Else
        If iNewValue = iValue Then
            tTimer.Enabled = False
        Else
            tTimer.Enabled = True
        End If
        Dim lDif As Long
        lDif = Abs(iValue - iNewValue)
        shpAdd.Visible = False
        shpSub.Visible = False
        If iNewValue < iValue Then
            iNewValue = iNewValue + 1
            Select Case lDif
                Case Is > 500
                    iNewValue = iNewValue + (lDif / 8)
                Case Is > 100
                    iNewValue = iNewValue + (lDif / 14)
                Case Is > 10
                    iNewValue = iNewValue + (lDif / 18)
            End Select
            If iNewValue > iValue Then iNewValue = iValue
            bAnimating = True
            shpAdd.Width = (((iNewValue / 100) / (iMax / 100)) * UserControl.Width)
            shpAdd.Visible = True
        ElseIf iNewValue > iValue Then
            iNewValue = iNewValue - 1
            Select Case lDif
                Case Is > 500
                    iNewValue = iNewValue - (lDif / 8)
                Case Is > 100
                    iNewValue = iNewValue - (lDif / 14)
                Case Is > 10
                    iNewValue = iNewValue - (lDif / 18)
            End Select
            If iNewValue < iValue Then iNewValue = iValue
            bAnimating = True
            shpSub.Width = (((iNewValue / 100) / (iMax / 100)) * UserControl.Width)
            shpSub.Visible = True
        Else
            iNewValue = iValue
            bAnimating = False
            shpAdd.Visible = False
            shpSub.Visible = False
            shpStat.Width = (((iValue / 100) / (iMax / 100)) * UserControl.Width)
        End If
        If lDif > (iMax / 10) Or iValue < (iMax / 10) Then
            tTimer.Interval = 1
        Else
            tTimer.Interval = 30
        End If
        If LenB(sCustomText) = 0 Then
            lblStat.Caption = iNewValue & "/" & iMax
            lblShadowText.Caption = lblStat.Caption
        End If
        If shpSub.Visible Then
           shpStat.Width = (((iValue / 100) / (iMax / 100)) * UserControl.Width)
        End If
        shpBack.Refresh
    End If
    
End Sub




Private Sub lblStat_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*************************************************
'Author: ^[GS]^
'Last modified: 03/09/2012
'*************************************************

On Error Resume Next

    If LenB(MouseOverText) = 0 Then
        MouseOverText = Round(CDbl(iNewValue) * CDbl(100) / CDbl(iMax), 2) & "%"
    End If
    
    If LenB(CustomText) = 0 Then
        lblStat.Caption = MouseOverText
        lblShadowText.Caption = lblStat.Caption
    End If
    
   
End Sub

Private Sub lblStat_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*************************************************
'Author: ^[GS]^
'Last modified: 03/09/2012
'*************************************************

On Error Resume Next

    If LenB(CustomText) = 0 Then
        lblStat.Caption = iNewValue & "/" & iMax
        lblShadowText.Caption = lblStat.Caption
    End If
    
End Sub

Private Sub tDanger_Timer()
    If iValue < iMinDanger Then
        If LenB(tDanger.Tag) = 0 Then
            tDanger.Tag = "0"
            lblStat.ForeColor = lForeColor
            lblShadowText.ForeColor = lShadowTextColor
            shpStat.FillColor = lBackColor
            shpBack.FillColor = lBackgroundColor
        Else
            tDanger.Tag = vbNullString
            'lblStat.ForeColor = lBackColor
            'lblShadowText.ForeColor = lBackgroundDangerColor
            shpStat.FillColor = lBackDangerColor
            shpBack.FillColor = lBackgroundDangerColor
        End If
    End If
End Sub

Private Sub tTimer_Timer()
'*************************************************
'Author: ^[GS]^
'Last modified: 03/09/2012
'*************************************************

On Error Resume Next
    
    If bEnabled = False Then
        tTimer.Enabled = False
        Exit Sub
    End If
    If bAnimating = True Then
        Call DrawStat
    End If
    
End Sub

Private Sub ResizeLabel()
'*************************************************
'Author: ^[GS]^
'Last modified: 25/08/2012
'*************************************************

On Error Resume Next
    
    lblStat.Left = 0
    lblStat.Width = UserControl.Width
    lblStat.Top = (UserControl.Height / 2) - ((lblStat.Height / 2))
    lblShadowText.Left = lblStat.Left + 30
    lblShadowText.Width = lblStat.Width
    lblShadowText.Top = lblStat.Top + 30
    Call DrawStat
    
End Sub

Private Sub UserControl_InitProperties()
'*************************************************
'Author: ^[GS]^
'Last modified: 25/04/2013
'*************************************************

On Error Resume Next

    iMax = 100
    iMin = 1
    iValue = 1
    iMinDanger = 0
    bShowText = True
    sCustomText = vbNullString
    bEnabled = True
    bAnimate = True
    bUseBackground = False
    lBackgroundColor = RGB(0, 0, 0)
    lForeColor = RGB(255, 255, 255)
    lBackColor = RGB(100, 100, 100)
    lBackSubColor = RGB(75, 75, 75)
    lBackAddColor = RGB(125, 125, 125)
    lBackDangerColor = RGB(125, 0, 0)
    lBackgroundDangerColor = RGB(0, 0, 0)
    lBorderColor = RGB(200, 200, 200)
    
End Sub

Private Sub UserControl_Resize()
'*************************************************
'Author: ^[GS]^
'Last modified: 25/04/2013
'*************************************************

On Error Resume Next
    
    shpStat.Left = 0
    shpStat.Height = UserControl.Height
    shpBack.Height = UserControl.Height
    shpSub.Height = UserControl.Height
    shpAdd.Height = UserControl.Height
    shpBack.Width = UserControl.Width
    Call ResizeLabel
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'*************************************************
'Author: ^[GS]^
'Last modified: 13/04/2021
'*************************************************

On Error Resume Next
    
    Debug.Print PropBag.Contents
    
    With PropBag
        iMax = .ReadProperty("Max", 100)
        iMin = .ReadProperty("Min", 0)
        iMinDanger = .ReadProperty("MinDanger", 0)
        iValue = .ReadProperty("Value", 50)
        bEnabled = .ReadProperty("Enabled", True)
        bAnimate = .ReadProperty("Animate", True)
        bUseBackground = .ReadProperty("UseBackground", True)
        lShadowTextColor = .ReadProperty("ShadowTextColor", RGB(0, 0, 0))
        lBackgroundColor = .ReadProperty("BackgroundColor", RGB(0, 0, 0))
        lBackgroundDangerColor = .ReadProperty("BackgroundDangerColor", RGB(0, 0, 0))
        lForeColor = .ReadProperty("ForeColor", RGB(255, 255, 255))
        lBackColor = .ReadProperty("BackColor", RGB(100, 100, 100))
        lBackAddColor = .ReadProperty("BackAddColor", RGB(75, 75, 75))
        lBackSubColor = .ReadProperty("BackSubColor", RGB(125, 125, 125))
        lBackDangerColor = .ReadProperty("BackDangerColor", RGB(125, 0, 0))
        lBorderColor = .ReadProperty("BorderColor", RGB(200, 200, 200))
        bShowText = .ReadProperty("ShowText", True)
        sCustomText = .ReadProperty("CustomText", vbNullString)
        Set lblStat.Font = .ReadProperty("FONT", lblStat.Font)
        Set lblShadowText.Font = .ReadProperty("FONT", lblStat.Font)
    End With
    
    lblStat.ForeColor = lForeColor
    If LenB(sCustomText) > 0 Then
        lblStat.Caption = sCustomText
        lblShadowText.Caption = lblStat.Caption
    End If
    lblStat.Visible = bShowText
    lblShadowText.ForeColor = ShadowTextColor
    lblShadowText.Visible = lblStat.Visible
    shpStat.FillColor = lBackColor
    shpStat.BorderColor = lBorderColor
    shpBack.BorderColor = lBorderColor
    shpAdd.FillColor = lBackAddColor
    shpAdd.BorderColor = lBorderColor
    shpSub.FillColor = lBackSubColor
    shpSub.BorderColor = lBorderColor
    shpBack.BackColor = lBackgroundColor
    shpBack.Visible = bUseBackground
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'*************************************************
'Author: ^[GS]^
'Last modified: 13/04/2021
'*************************************************

On Error Resume Next
    
    With PropBag
        .WriteProperty "Max", iMax, 100
        .WriteProperty "Min", iMin, 0
        .WriteProperty "MinDanger", iMinDanger, 0
        .WriteProperty "Value", iValue, 50
        .WriteProperty "Enabled", bEnabled, True
        .WriteProperty "Animate", bAnimate, True
        .WriteProperty "UseBackground", bUseBackground, True
        .WriteProperty "ShadowTextColor", lShadowTextColor, RGB(0, 0, 0)
        .WriteProperty "BackgroundColor", lBackgroundColor, RGB(0, 0, 0)
        .WriteProperty "BackgroundDangerColor", lBackgroundDangerColor, RGB(0, 0, 0)
        .WriteProperty "ForeColor", lForeColor, RGB(255, 255, 255)
        .WriteProperty "BackColor", lBackColor, RGB(100, 100, 100)
        .WriteProperty "BackAddColor", lBackAddColor, RGB(125, 125, 125)
        .WriteProperty "BackDangerColor", lBackDangerColor, RGB(125, 0, 0)
        .WriteProperty "BackSubColor", lBackSubColor, RGB(75, 75, 75)
        .WriteProperty "BorderColor", lBorderColor, RGB(200, 200, 200)
        .WriteProperty "ShowText", bShowText, True
        .WriteProperty "CustomText", sCustomText, ""
        Call .WriteProperty("FONT", lblStat.Font)
        Call .WriteProperty("FONT", lblShadowText.Font)
    End With
    
End Sub

Public Property Get Enabled() As Boolean
'*************************************************
'Author: ^[GS]^
'Last modified: 25/08/2012
'*************************************************

On Error Resume Next
    
    Enabled = bEnabled
    
End Property

Public Property Let Enabled(ByVal NewValue As Boolean)
'*************************************************
'Author: ^[GS]^
'Last modified: 25/08/2012
'*************************************************

On Error Resume Next
    
    bEnabled = NewValue
    PropertyChanged "Enabled"
    
    UserControl.Enabled = False
    
End Property

Public Property Get Animado() As Boolean
'*************************************************
'Author: ^[GS]^
'Last modified: 25/08/2012
'*************************************************

On Error Resume Next
    
    Animado = bAnimate
    
End Property

Public Property Let Animado(ByVal NewValue As Boolean)
'*************************************************
'Author: ^[GS]^
'Last modified: 25/08/2012
'*************************************************

On Error Resume Next
    
    bAnimate = NewValue
    PropertyChanged "Animate"
    
    Call DrawStat
    
End Property

Public Property Get UseBackground() As Boolean
'*************************************************
'Author: ^[GS]^
'Last modified: 25/04/2013
'*************************************************

On Error Resume Next
    
    UseBackground = bUseBackground
    
End Property

Public Property Let UseBackground(ByVal NewValue As Boolean)
'*************************************************
'Author: ^[GS]^
'Last modified: 25/04/2013
'*************************************************

On Error Resume Next
    
    bUseBackground = NewValue
    PropertyChanged "UseBackground"
    
    shpBack.Visible = bUseBackground
    
End Property

Public Property Get Font() As Font
'*************************************************
'Author: ^[GS]^
'Last modified: 25/08/2012
'*************************************************

On Error Resume Next
    
    Set Font = lblStat.Font
    
End Property

Public Property Set Font(ByRef newFont As Font)
'*************************************************
'Author: ^[GS]^
'Last modified: 25/08/2012
'*************************************************

On Error Resume Next
    
    Set lblStat.Font = newFont
    Set lblShadowText.Font = newFont

    Call ResizeLabel

    PropertyChanged "FONT"
    
End Property

Public Property Get FontBold() As Boolean
'*************************************************
'Author: ^[GS]^
'Last modified: 25/08/2012
'*************************************************

On Error Resume Next
    
    FontBold = lblStat.FontBold
    
End Property

Public Property Let FontBold(ByVal NewValue As Boolean)
'*************************************************
'Author: ^[GS]^
'Last modified: 25/08/2012
'*************************************************

On Error Resume Next
    
    lblStat.FontBold = NewValue
    lblShadowText.FontBold = lblStat.FontBold
    
    Call ResizeLabel
    
End Property

Public Property Get FontItalic() As Boolean
'*************************************************
'Author: ^[GS]^
'Last modified: 25/08/2012
'*************************************************

On Error Resume Next
    
    FontItalic = lblStat.FontItalic
    
End Property

Public Property Let FontItalic(ByVal NewValue As Boolean)
'*************************************************
'Author: ^[GS]^
'Last modified: 25/08/2012
'*************************************************

On Error Resume Next
    
    lblStat.FontItalic = NewValue
    lblShadowText.FontItalic = lblStat.FontItalic

    Call ResizeLabel
    
End Property

Public Property Get FontUnderline() As Boolean
'*************************************************
'Author: ^[GS]^
'Last modified: 25/08/2012
'*************************************************

On Error Resume Next
    
    FontUnderline = lblStat.FontUnderline
    
End Property

Public Property Let FontUnderline(ByVal NewValue As Boolean)
'*************************************************
'Author: ^[GS]^
'Last modified: 25/08/2012
'*************************************************

On Error Resume Next
    
    lblStat.FontUnderline = NewValue
    lblShadowText.FontUnderline = lblStat.FontUnderline

    Call ResizeLabel
    
End Property

Public Property Get FontSize() As Integer
'*************************************************
'Author: ^[GS]^
'Last modified: 25/08/2012
'*************************************************

On Error Resume Next
    
    FontSize = lblStat.FontSize
    
End Property

Public Property Let FontSize(ByVal NewValue As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 25/08/2012
'*************************************************

On Error Resume Next
    
    lblStat.FontSize = NewValue
    lblShadowText.FontSize = lblStat.FontSize

    Call ResizeLabel
    
End Property

Public Property Get FontName() As String
'*************************************************
'Author: ^[GS]^
'Last modified: 25/08/2012
'*************************************************

On Error Resume Next
    
    FontName = lblStat.FontName
    
End Property

Public Property Let FontName(ByVal NewValue As String)
'*************************************************
'Author: ^[GS]^
'Last modified: 25/08/2012
'*************************************************

On Error Resume Next
    
    lblStat.FontName = NewValue
    lblShadowText.FontName = lblStat.FontName
    
    Call ResizeLabel
    
End Property

Public Property Get ForeColor() As OLE_COLOR
'*************************************************
'Author: ^[GS]^
'Last modified: 25/08/2012
'*************************************************

On Error Resume Next
    
    ForeColor = lForeColor
    
End Property

Public Property Let ForeColor(ByVal NewValue As OLE_COLOR)
'*************************************************
'Author: ^[GS]^
'Last modified: 25/08/2012
'*************************************************

On Error Resume Next
    
    lForeColor = NewValue
    PropertyChanged "ForeColor"
    
    lblStat.ForeColor = lForeColor
    
End Property

Public Property Get BackgroundColor() As OLE_COLOR
'*************************************************
'Author: ^[GS]^
'Last modified: 25/04/2013
'*************************************************

On Error Resume Next
    
    BackgroundColor = lBackgroundColor
    
End Property

Public Property Let BackgroundColor(ByVal NewValue As OLE_COLOR)
'*************************************************
'Author: ^[GS]^
'Last modified: 25/04/2013
'*************************************************

On Error Resume Next
    
    lBackgroundColor = NewValue
    PropertyChanged "BackgroundColor"
    
    shpBack.FillColor = lBackgroundColor
    
End Property

Public Property Get BackColor() As OLE_COLOR
'*************************************************
'Author: ^[GS]^
'Last modified: 25/08/2012
'*************************************************

On Error Resume Next
    
    BackColor = lBackColor
    
End Property

Public Property Let BackColor(ByVal NewValue As OLE_COLOR)
'*************************************************
'Author: ^[GS]^
'Last modified: 25/08/2012
'*************************************************

On Error Resume Next
    
    lBackColor = NewValue
    PropertyChanged "BackColor"
    
    shpStat.FillColor = lBackColor
    
End Property

Public Property Get BorderColor() As OLE_COLOR
'*************************************************
'Author: ^[GS]^
'Last modified: 25/08/2012
'*************************************************

On Error Resume Next
    
    BorderColor = lBorderColor
    
End Property

Public Property Let BorderColor(ByVal NewValue As OLE_COLOR)
'*************************************************
'Author: ^[GS]^
'Last modified: 25/04/2013
'*************************************************

On Error Resume Next
    
    lBorderColor = NewValue
    PropertyChanged "BorderColor"
    
    shpStat.BorderColor = lBorderColor
    shpBack.BackColor = lBorderColor
    shpAdd.BorderColor = lBorderColor
    shpSub.BorderColor = lBorderColor
    
End Property

Public Property Let Value(ByVal NewValue As Long)
'*************************************************
'Author: ^[GS]^
'Last modified: 25/08/2012
'*************************************************

On Error Resume Next
    
    If NewValue > iMax Then NewValue = iMax
    If NewValue < iMin Then NewValue = iMin
    iValue = NewValue
    
    PropertyChanged "Value"
    
    Call DrawStat
    
End Property

Public Property Get Value() As Long
'*************************************************
'Author: ^[GS]^
'Last modified: 25/08/2012
'*************************************************

On Error Resume Next
    
    Value = iValue
    
End Property

Public Property Let Max(ByVal NewValue As Long)
'*************************************************
'Author: ^[GS]^
'Last modified: 25/08/2012
'*************************************************

On Error Resume Next
    
    If NewValue < 1 Then NewValue = 1
    If NewValue <= iMin Then NewValue = iMin + 1
    iMax = NewValue
    
    If Value > iMax Then Value = iMax
    PropertyChanged "Max"
    
    Call DrawStat
    
End Property

Public Property Get Max() As Long
'*************************************************
'Author: ^[GS]^
'Last modified: 25/08/2012
'*************************************************

On Error Resume Next
    
    Max = iMax
    
End Property

Public Property Let Min(ByVal NewValue As Long)
'*************************************************
'Author: ^[GS]^
'Last modified: 25/08/2012
'*************************************************

On Error Resume Next
    
    If NewValue >= iMax Then NewValue = Max - 1
    If NewValue < 0 Then NewValue = 0
    iMin = NewValue
    If Value < iMin Then Value = iMin
    
    PropertyChanged "Min"
    
    Call DrawStat
    
End Property

Public Property Get Min() As Long
'*************************************************
'Author: ^[GS]^
'Last modified: 25/08/2012
'*************************************************

On Error Resume Next
    
    Min = iMin
    
End Property


Public Property Get ShowText() As Boolean
'*************************************************
'Author: ^[GS]^
'Last modified: 13/04/2021
'*************************************************

On Error Resume Next
    
    ShowText = bShowText
    
End Property

Public Property Let ShowText(ByVal NewValue As Boolean)
'*************************************************
'Author: ^[GS]^
'Last modified: 13/04/2021
'*************************************************

On Error Resume Next
    
    bShowText = NewValue
    PropertyChanged "ShowText"
    
    lblStat.Visible = bShowText
    
End Property

Public Property Get CustomText() As String
'*************************************************
'Author: ^[GS]^
'Last modified: 13/04/2021
'*************************************************

On Error Resume Next
    
    CustomText = sCustomText
    
End Property


Public Property Let CustomText(ByVal NewValue As String)
'*************************************************
'Author: ^[GS]^
'Last modified: 13/04/2021
'*************************************************

On Error Resume Next
    
    sCustomText = NewValue
    PropertyChanged "CustomText"
    
    lblStat.Caption = sCustomText
    
End Property


Public Property Get BackAddColor() As OLE_COLOR
'*************************************************
'Author: ^[GS]^
'Last modified: 13/04/2021
'*************************************************

On Error Resume Next
    
    BackAddColor = lBackAddColor
    
End Property


Public Property Let BackAddColor(ByVal NewValue As OLE_COLOR)
'*************************************************
'Author: ^[GS]^
'Last modified: 13/04/2021
'*************************************************

On Error Resume Next
    
    lBackAddColor = NewValue
    PropertyChanged "BackAddColor"
    
    shpAdd.FillColor = lBackAddColor
    
End Property

Public Property Get BackSubColor() As OLE_COLOR
'*************************************************
'Author: ^[GS]^
'Last modified: 13/04/2021
'*************************************************

On Error Resume Next
    
    BackSubColor = lBackSubColor
    
End Property


Public Property Let BackSubColor(ByVal NewValue As OLE_COLOR)
'*************************************************
'Author: ^[GS]^
'Last modified: 13/04/2021
'*************************************************

On Error Resume Next
    
    lBackSubColor = NewValue
    PropertyChanged "BackSubColor"
    
    shpSub.FillColor = lBackSubColor
    
End Property

Public Property Get BackDangerColor() As OLE_COLOR
'*************************************************
'Author: ^[GS]^
'Last modified: 13/04/2021
'*************************************************

On Error Resume Next
    
    BackDangerColor = lBackDangerColor
    
End Property


Public Property Let BackDangerColor(ByVal NewValue As OLE_COLOR)
'*************************************************
'Author: ^[GS]^
'Last modified: 13/04/2021
'*************************************************

On Error Resume Next
    
    lBackDangerColor = NewValue
    PropertyChanged "BackDangerColor"

End Property


Public Property Get MinDanger() As Long
'*************************************************
'Author: ^[GS]^
'Last modified: 13/04/2021
'*************************************************

On Error Resume Next
    
    MinDanger = iMinDanger
    
End Property


Public Property Let MinDanger(ByVal NewValue As Long)
'*************************************************
'Author: ^[GS]^
'Last modified: 13/04/2021
'*************************************************

On Error Resume Next
    
    iMinDanger = NewValue
    PropertyChanged "MinDanger"

End Property


Public Property Get BackgroundDangerColor() As OLE_COLOR
'*************************************************
'Author: ^[GS]^
'Last modified: 13/04/2021
'*************************************************

On Error Resume Next
    
    BackgroundDangerColor = lBackgroundDangerColor
    
End Property


Public Property Let BackgroundDangerColor(ByVal NewValue As OLE_COLOR)
'*************************************************
'Author: ^[GS]^
'Last modified: 13/04/2021
'*************************************************

On Error Resume Next
    
    lBackgroundDangerColor = NewValue
    PropertyChanged "BackgroundDangerColor"

End Property


Public Property Get ShadowTextColor() As OLE_COLOR
'*************************************************
'Author: ^[GS]^
'Last modified: 13/04/2021
'*************************************************

On Error Resume Next
    
    ShadowTextColor = lShadowTextColor
    
End Property


Public Property Let ShadowTextColor(ByVal NewValue As OLE_COLOR)
'*************************************************
'Author: ^[GS]^
'Last modified: 13/04/2021
'*************************************************

On Error Resume Next
    
    lShadowTextColor = NewValue
    PropertyChanged "ShadowTextColor"

End Property

