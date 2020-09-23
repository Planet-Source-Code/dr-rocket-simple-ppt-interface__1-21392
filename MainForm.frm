VERSION 5.00
Begin VB.Form MainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simple PowerPoint Interface"
   ClientHeight    =   3195
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   9165
   ControlBox      =   0   'False
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   9165
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton MiniatureBtn 
      Caption         =   "Show miniature"
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   2580
      Width           =   1575
   End
   Begin VB.ListBox TextBoxLst 
      Height          =   2205
      Left            =   4680
      TabIndex        =   5
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton AddTextBoxBtn 
      Caption         =   "Add a Textbox."
      Height          =   375
      Left            =   7440
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton ShowHideBtn 
      Caption         =   "Show/Hide"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2580
      Width           =   1575
   End
   Begin VB.CommandButton SldDetailsBtn 
      Caption         =   "Show slide details."
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.ListBox SlideLst 
      Height          =   2205
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton AddSldBtn 
      Caption         =   "Add a slide."
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label MessageLine 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   255
      Left            =   4680
      TabIndex        =   6
      Top             =   2640
      Width           =   4335
   End
   Begin VB.Menu menuExit 
      Caption         =   "E&xit"
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' I have put this project together to teach myself how to use Microsoft PowerPoint
' as an automation server. This, for me, is an exercise and not an example of good
' software design. For example, I am using PowerPoint 9 as in Office 2000. A solid
' application would check to see what version, if any, is on the user's system and
' adjust its behavior accordingly. Some version 9 features are not available
' version 8.
'
' The program starts out by creating a new PowerPoint presentation.
' The user can add slides to the presentation by clicking on the appropriate button.
' The user can give the slide a name via the input box or click OK to use the default
' name. The name of the slide is added to the list on the left. When the user clicks
' a name in this list then, the corresponding slide becomes the current selection.
' To add a text box to a slide the user clicks any slide in the left list and then
' the Add Textbox button. This brings up a dialog that gathers the text box para-
' meters. (Note: if the user has not selected a slide then, the Add Textbox button
' ignores its click event.)
'
' In this program I am using list boxes to store the names of slides and text boxes
' as they are generated. This gives the user both a graphical interface to keep track
' of the PowerPoint objects in play and provides a programatic means to track the
' same objects. If the PowerPoint manipulations are to occur "behind the scenes" in an
' actual application, then the names could just as well have been stored in some other
' convenient data structure.
'
' The user can show/hide the PowerPoint application by clicking the Show/Hide button.
' When PowerPoint is made visible it is maximized. The "Alt-Tab" method is used to get
' back to this program and the Show/Hide button to minimize the PowerPoint application.
' If the user closes PowerPoint outside of this program then, all subsequent references
' to PowerPoint objects will be invalid.
'
' The Show miniature button does not work, in so far as I have not yet figured out
' how this might work. (Suggestions?)
'
' This program acts as a simplified user interface to PowerPoint. This is rather
' pointless since PowerPoint has an elaborate user interface. A more realistic
' application would automatically prepare a presentation based on data gathered
' from, say, an Access database. However, this is just an exercise, so simple is
' best.
'
' If anyone out there has a better way to do any of the things that this program does
' I would like to hear about it. (mail to: jimd@ftassociates.com)
'
' This program assumes that you have PowerPoint and Office on your system. Be sure to
' reference PowerPoint and Office. Use the Reference item on the Project menu to do
' this.
'
' (By the way, in all that follows "Tbp" is meant to suggest "TextBoxParameter.")
'
Sub BeginPowerPoint()
 '
 ' Begin the PowerPoint application in a minimized window.
 '
 Set PptApp = New PowerPoint.Application
 PptApp.Activate
 PptApp.WindowState = ppWindowMinimized
 '
 ' Add a new presentation and create a reference to it. Exhibit the system generated name.
 '
 Set Present = PptApp.Presentations.Add
 
 MsgBox Present.Name
 '
 ' Create a reference to the slides collection in the new presentation.
 '
 Set Slds = Present.Slides
 '
 ' Gather some useful information.
 '
 MainHeight = Present.PageSetup.SlideHeight
 MainWidth = Present.PageSetup.SlideWidth
 '
End Sub
'
Private Sub AddTextBoxBtn_Click()
 '
 ' The user has indicated that a text box should be added to a slide.
 '
 ' This is used to reference the slide to which the text box is to be added.
 '
 Dim S As PowerPoint.Slide
 '
 ' Add the text box only if the user has selected a slide from SlideLst.
 '
 If SlideLst.ListIndex > -1 Then
  '
  ' Create a refernce to the slide that will get the new text box.
  '
  Set S = Slds(SlideLst.Text)
  '
  ' The user needs to supply some properties for the new text box. Show modal so as to
  ' require property input.
  '
  TbpFlag = True
  TextBoxProps.Show vbModal, Me
  '
  ' Make the new text box.
  '
  If TbpFlag = True Then InsertTextBox S, S.Name & "TextBox" & Format(ShapeCnt, "000")
 Else
  '
  ' Inform the user that a slide selection is needed.
  '
  MsgBox "Select a slide first."
 End If
 '
End Sub
'
Sub InsertTextBox(ByVal S As PowerPoint.Slide, T As String)
 '
 ' Add a text box to the slide S and give it the name T.
 '
 ' This is a reference to the text box.
 '
 Dim Sh As PowerPoint.Shape
 '
 ' Use the user supplied properties to create the text box.
 '
 S.Shapes.AddTextbox(msoTextOrientationHorizontal, TbpLeft(TbpCurrent), TbpTop(TbpCurrent), _
  TbpWidth(TbpCurrent), TbpHeight(TbpCurrent)).Select
 PptApp.Windows(1).Selection.ShapeRange.Name = T
 PptApp.Windows(1).Selection.ShapeRange.TextFrame.WordWrap = msoTrue
 With PptApp.Windows(1).Selection.ShapeRange
  .Line.Visible = msoTrue
  .Line.Style = TbpLineStyle(TbpCurrent)
  .Line.Weight = TbpLineWidth(TbpCurrent)
  .Line.ForeColor.RGB = TbpLineColor(TbpCurrent)
  .Line.Visible = msoTrue
 End With
 PptApp.Windows(1).Selection.ShapeRange.TextFrame.TextRange.Characters(Start:=1, _
  Length:=0).Select
 With PptApp.Windows(1).Selection.TextRange
  .ParagraphFormat.Alignment = TbpAlignment(TbpCurrent)
  .Text = TbpText(TbpCurrent)
  With .Font
   .Name = TbpFontName(TbpCurrent)
   .Size = TbpFontSize(TbpCurrent)
   If TbpFontNormal(TbpCurrent) = True Then
    .Bold = msoFalse
    .Italic = msoFalse
    .Underline = msoFalse
    .Shadow = msoFalse
    .Emboss = msoFalse
   Else
    If TbpFontBold(TbpCurrent) = True Then .Italic = msoTrue
    If TbpFontItalic(TbpCurrent) = True Then .Bold = msoTrue
    If TbpFontUnderlined(TbpCurrent) = True Then .Underline = msoTrue
    If TbpFontShadow(TbpCurrent) = True Then .Shadow = msoTrue
    If TbpFontEmbossed(TbpCurrent) = True Then .Emboss = msoTrue
   End If
   .BaselineOffset = 0
   .AutoRotateNumbers = msoFalse
   .Color.RGB = TbpFontColor(TbpCurrent)
  End With
 End With
 '
 ' Keep track of how many shapes are out there.
 '
 ShapeCnt = ShapeCnt + 1
 '
 ' Revise the list of shapes (text boxes) that are associated with the selected slide.
 '
 If SlideLst.ListIndex > -1 Then
  PptApp.Windows(1).View.GotoSlide Index:=Slds(S.Name).SlideIndex
  With TextBoxLst
   .Clear
   For Each Sh In Slds(S.Name).Shapes
    .AddItem Sh.Name
   Next
  End With
 End If
 '
End Sub
'
Private Sub AddSldBtn_Click()
 '
 ' The user has indicated that a new slide should be added to the collection.
 '
 Dim Reply As String
 '
 ' Give the user the opportunity to name the new slide.
 '
 Reply = InputBox("Name the slide.", CStr(Slds.Count + 1))
 '
 ' Now add the new slide.
 '
 Slds.Add Slds.Count + 1, ppLayoutBlank
 '
 ' Make a reference to the slide and rename it.
 '
 Set Sld = Slds(Slds.Count)
 Sld.Name = Reply
 '
 ' Add a reference to the new slide into SlideLst.
 '
 SlideLst.AddItem Sld.Name
 SlideLst.Refresh
 '
End Sub
'
Private Sub Form_Load()
 '
 ' Initialize Powerpoint.
 '
 BeginPowerPoint
 ShapeCnt = 1
 '
 ' Initialize the text box properties (Tbp for short.)
 '
 TbpDefault = 1
 TbpCurrent = 2
 TbpTemp = 3
 TbpSetDefaults
 TbpDefaultToCurrent
 TbpCurrentToTemp
 '
End Sub
'
Private Sub SlideLst_Click()
 '
 ' The user has clicked SlideLst.
 '
 Dim MsgStr As String
 Dim S As PowerPoint.Shape
 '
 ' Only if a list item has been selected:
 '
 If SlideLst.ListIndex > -1 Then
  '
  ' Get the selected item.
  '
  TheSlideName = SlideLst.Text
  Slds(TheSlideName).Select
  'PptApp.Windows(1).View.GotoSlide Index:=Slds(TheSlideName).SlideIndex
  '
  ' Revise TextBoxLst to reflect the users selection.
  '
  With TextBoxLst
   .Clear
   For Each S In Slds(TheSlideName).Shapes
    .AddItem S.Name
   Next
  End With
  '
  ' Show the user something.
  '
  MsgStr = TheSlideName & " "
  MsgStr = MsgStr & " " & Slds(TheSlideName).Application.Name
  MsgStr = MsgStr & " " & Slds(TheSlideName).Name
  MessageLine.Caption = MsgStr
 Else
  '
  ' If this line executes, something odd has happened.
  '
  MessageLine.Caption = "There are no slides."
 End If
 '
End Sub
'
Private Sub menuExit_Click()
 '
 ' The user wants to end it. Make sure PowerPoint also ends.
 '
 PptApp.Quit
 Set PptApp = Nothing
 End
 '
End Sub
'
Private Sub MiniatureBtn_Click()
 '
 ' I wanted to show a miniature of the current slide here, but I have not figured out
 ' how to do this. The PowerPoint application's user interface will only show a
 ' miniature if the current slide is bigger than the view window or in black-and-white
 ' view.
 '
 MsgBox "Still under construction. Sorry."
End Sub
'
Private Sub ShowHideBtn_Click()
 '
 ' The user want to see or hide what is going on in PowerPoint.
 '
 If PptApp.WindowState = ppWindowNormal Or PptApp.WindowState = ppWindowMaximized Then
  PptApp.WindowState = ppWindowMinimized
 Else
  PptApp.WindowState = ppWindowNormal
 End If
 '
End Sub
'
Private Sub SldDetailsBtn_Click()
 '
 ' The user wants to see the slide names.
 '
 Dim S As PowerPoint.Slide
 
 If SlideLst.ListCount > 0 Then
  For Each S In Slds
   MsgBox "Slide name = " & S.Name & vbCrLf & "Slide Id = " & CStr(S.SlideID)
  Next S
 Else
  MsgBox "No slides to report on."
 End If
 '
End Sub
