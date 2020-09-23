Attribute VB_Name = "Module1"
'
' These are used by the MianForm and the TextBoxProps forms to communicate.
' When the MainForm loads the default values are stored into element 1 and 2 of the
' appropriate arrays. Element 2 of the arrays holds the current working value. When
' the user is entering TextBox properties the values are held in element 3 of the
' arrays. When the user is happy with the values and commits to them, then they are
' transferred to element 2. The values then become the current working values.
' This may not be the best way for two forms to communicate but it is simple and
' controled.
' The Tbp prefix is to indicate Text Box Properties.
'
Public TbpDefault As Integer
Public TbpCurrent As Integer
Public TbpTemp As Integer
'
Public TbpText(1 To 3) As String
Public TbpAlignment(1 To 3) As Integer
Public TbpFontName(1 To 3) As String
Public TbpFontSize(1 To 3) As Integer
Public TbpFontColor(1 To 3) As Long
Public TbpFontNormal(1 To 3) As Boolean
Public TbpFontBold(1 To 3) As Boolean
Public TbpFontItalic(1 To 3) As Boolean
Public TbpFontUnderlined(1 To 3) As Boolean
Public TbpFontShadow(1 To 3) As Boolean
Public TbpFontEmbossed(1 To 3) As Boolean
Public TbpLineWidth(1 To 3) As Single
Public TbpLineStyle(1 To 3) As Single
Public TbpLineColor(1 To 3) As Long
Public TbpLeft(1 To 3) As Single
Public TbpTop(1 To 3) As Single
Public TbpWidth(1 To 3) As Single
Public TbpHeight(1 To 3) As Single
Public TbpFlag As Boolean
'
' These are used to keep track of program organization.
'
Public TheSlideName As String
Public MainHeight As Single
Public MainWidth As Single
Public ShapeCnt As Integer
'
' These are the PowerPoint objects.
'
Public PptApp As PowerPoint.Application
Public Present As PowerPoint.Presentation
Public Slds As PowerPoint.Slides
Public Sld As PowerPoint.Slide
Public Shp As PowerPoint.Shape
'
Public Sub TbpSetDefaults()
 '
 ' Store a reasonable set of defaults. (Tbp is short for TextBox Properties.)
 ' See commentary in Module1 for further explanation.
 '
 TbpText(TbpDefault) = "Sample text."
 TbpAlignment(TbpDefault) = ppAlignCenter
 TbpFontName(TbpDefault) = "Arial"
 TbpFontSize(TbpDefault) = 36
 TbpFontColor(TbpDefault) = RGB(0, 0, 0)
 TbpFontNormal(TbpDefault) = True
 TbpFontBold(TbpDefault) = False
 TbpFontItalic(TbpDefault) = False
 TbpFontUnderlined(TbpDefault) = False
 TbpFontShadow(TbpDefault) = False
 TbpFontEmbossed(TbpDefault) = False
 TbpLineWidth(TbpDefault) = 3.5
 TbpLineStyle(TbpDefault) = msoLineThickThin
 TbpLineColor(TbpDefault) = RGB(0, 0, 0)
 TbpLeft(TbpDefault) = 36
 TbpTop(TbpDefault) = 36
 TbpWidth(TbpDefault) = 288
 TbpHeight(TbpDefault) = 50
 '
End Sub
'
Public Sub TbpDefaultToCurrent()
 '
 ' Transfer the default values into the current values, i.e., do a reset.
 ' See commentary in Module1 for further explanation.
 '
 TbpMove TbpDefault, TbpCurrent
 '
End Sub
'
Public Sub TbpTempToCurrent()
 '
 ' Save the temporary working values into the current.
 ' See commentary in Module1 for further explanation.
 '
 TbpMove TbpTemp, TbpCurrent
 '
End Sub
'
Public Sub TbpDefaultToTemp()
 '
 ' Reset the tempory working values to thier default values.
 ' See commentary in Module1 for further explanation.
 '
  TbpMove TbpDefault, TbpTemp
 '
End Sub
'
Public Sub TbpCurrentToTemp()
 '
 ' Transfer the current vaalues to the temporary working set.
 ' See commentary in Module1 for further explanation.
 '
 TbpMove TbpCurrent, TbpTemp
 '
End Sub
'
Public Sub TbpMove(TbpFrom As Integer, TbpTo As Integer)
 '
 TbpText(TbpTo) = TbpText(TbpFrom)
 TbpAlignment(TbpTo) = TbpAlignment(TbpFrom)
 TbpFontName(TbpTo) = TbpFontName(TbpFrom)
 TbpFontSize(TbpTo) = TbpFontSize(TbpFrom)
 TbpFontColor(TbpTo) = TbpFontColor(TbpFrom)
 TbpFontNormal(TbpTo) = TbpFontNormal(TbpFrom)
 TbpFontBold(TbpTo) = TbpFontBold(TbpFrom)
 TbpFontItalic(TbpTo) = TbpFontItalic(TbpFrom)
 TbpFontUnderlined(TbpTo) = TbpFontUnderlined(TbpFrom)
 TbpFontShadow(TbpTo) = TbpFontShadow(TbpFrom)
 TbpFontEmbossed(TbpTo) = TbpFontEmbossed(TbpFrom)
 TbpLineWidth(TbpTo) = TbpLineWidth(TbpFrom)
 TbpLineStyle(TbpTo) = TbpLineStyle(TbpFrom)
 TbpLineColor(TbpTo) = TbpLineColor(TbpFrom)
 TbpLeft(TbpTo) = TbpLeft(TbpFrom)
 TbpTop(TbpTo) = TbpTop(TbpFrom)
 TbpWidth(TbpTo) = TbpWidth(TbpFrom)
 TbpHeight(TbpTo) = TbpHeight(TbpFrom)
 '
End Sub
