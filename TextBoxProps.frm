VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form TextBoxProps 
   Caption         =   "Pick text box characteristics."
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9855
   LinkTopic       =   "Form2"
   ScaleHeight     =   5175
   ScaleWidth      =   9855
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox LineWidthLst 
      Height          =   450
      Left            =   6120
      TabIndex        =   40
      Top             =   120
      Width           =   3615
   End
   Begin VB.CommandButton DefaultBtn 
      Caption         =   "Default Values"
      Height          =   375
      Left            =   4920
      TabIndex        =   39
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CommandButton OkBtn 
      Caption         =   "OK"
      Height          =   375
      Left            =   6600
      TabIndex        =   38
      Top             =   4680
      Width           =   1455
   End
   Begin VB.TextBox HeightTxt 
      Height          =   285
      Left            =   6120
      TabIndex        =   35
      Text            =   "0"
      Top             =   4095
      Width           =   615
   End
   Begin VB.TextBox WidthTxt 
      Height          =   285
      Left            =   6120
      TabIndex        =   32
      Text            =   "0"
      Top             =   3600
      Width           =   615
   End
   Begin VB.TextBox TopTxt 
      Height          =   285
      Left            =   6120
      TabIndex        =   29
      Text            =   "0"
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox LeftTxt 
      Height          =   285
      Left            =   6120
      TabIndex        =   26
      Text            =   "0"
      Top             =   2625
      Width           =   615
   End
   Begin MSComCtl2.UpDown LeftUpDown 
      Height          =   285
      Left            =   6720
      TabIndex        =   25
      Top             =   2625
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   503
      _Version        =   393216
      BuddyControl    =   "LeftTxt"
      BuddyDispid     =   196615
      OrigLeft        =   6720
      OrigTop         =   2625
      OrigRight       =   6960
      OrigBottom      =   2910
      SyncBuddy       =   -1  'True
      Wrap            =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.PictureBox GeometrySamplePic 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   7200
      ScaleHeight     =   141
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   141
      TabIndex        =   24
      Top             =   2280
      Width           =   2175
      Begin VB.Shape TextBoxShape 
         Height          =   255
         Left            =   600
         Top             =   720
         Width           =   735
      End
      Begin VB.Shape SlideShape 
         Height          =   1335
         Left            =   120
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.ListBox AlignmentLst 
      Height          =   645
      Left            =   1200
      TabIndex        =   21
      Top             =   600
      Width           =   3615
   End
   Begin VB.ListBox LineStyleLst 
      Height          =   645
      Left            =   6120
      TabIndex        =   19
      Top             =   840
      Width           =   3615
   End
   Begin VB.CommandButton LineColorBtn 
      Caption         =   "..."
      Height          =   315
      Left            =   6120
      TabIndex        =   17
      Top             =   1770
      Width           =   315
   End
   Begin VB.Frame Frame3 
      Caption         =   "Font style:"
      Height          =   2340
      Left            =   2160
      TabIndex        =   9
      Top             =   2280
      Width           =   1515
      Begin VB.OptionButton FontStyleOpt 
         Caption         =   "Normal"
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Value           =   -1  'True
         Width           =   1065
      End
      Begin VB.CheckBox FontStyleChk 
         Caption         =   "Embossed"
         Height          =   240
         Index           =   4
         Left            =   150
         TabIndex        =   14
         Top             =   1875
         Width           =   1140
      End
      Begin VB.CheckBox FontStyleChk 
         Caption         =   "Shadow"
         Height          =   240
         Index           =   3
         Left            =   150
         TabIndex        =   13
         Top             =   1575
         Width           =   1215
      End
      Begin VB.CheckBox FontStyleChk 
         Caption         =   "Underlined"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   150
         TabIndex        =   12
         Top             =   1275
         Width           =   1140
      End
      Begin VB.CheckBox FontStyleChk 
         Caption         =   "Italic"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   150
         TabIndex        =   11
         Top             =   975
         Width           =   915
      End
      Begin VB.CheckBox FontStyleChk 
         Caption         =   "Bold"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   150
         TabIndex        =   10
         Top             =   675
         Width           =   915
      End
   End
   Begin VB.CommandButton FontColorBtn 
      Caption         =   "..."
      Height          =   315
      Left            =   1200
      TabIndex        =   7
      Top             =   2280
      Width           =   315
   End
   Begin VB.ComboBox FontSize1 
      Height          =   315
      Left            =   4080
      TabIndex        =   5
      Text            =   "Combo3"
      Top             =   1770
      Width           =   690
   End
   Begin VB.ComboBox FontSlctr1 
      Height          =   315
      Left            =   1200
      Sorted          =   -1  'True
      TabIndex        =   3
      Text            =   "Combo2"
      Top             =   1770
      Width           =   2715
   End
   Begin VB.TextBox ContentsTxt 
      Height          =   315
      Left            =   1200
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   120
      Width           =   3615
   End
   Begin VB.CommandButton DismissBtn 
      Caption         =   "Dismiss"
      Height          =   375
      Left            =   8280
      TabIndex        =   0
      Top             =   4680
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1080
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComCtl2.UpDown TopUpDown 
      Height          =   285
      Left            =   6720
      TabIndex        =   28
      Top             =   3120
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   503
      _Version        =   393216
      BuddyControl    =   "TopTxt"
      BuddyDispid     =   196614
      OrigLeft        =   6720
      OrigTop         =   3120
      OrigRight       =   6960
      OrigBottom      =   3405
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown WidthUpDown 
      Height          =   285
      Left            =   6720
      TabIndex        =   31
      Top             =   3600
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   503
      _Version        =   393216
      BuddyControl    =   "WidthTxt"
      BuddyDispid     =   196613
      OrigLeft        =   6720
      OrigTop         =   3600
      OrigRight       =   6960
      OrigBottom      =   3885
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown HeightUpDown 
      Height          =   285
      Left            =   6720
      TabIndex        =   34
      Top             =   4095
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   503
      _Version        =   393216
      BuddyControl    =   "HeightTxt"
      BuddyDispid     =   196612
      OrigLeft        =   6720
      OrigTop         =   4095
      OrigRight       =   6960
      OrigBottom      =   4380
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "Geometry:"
      Height          =   255
      Left            =   5160
      TabIndex        =   37
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "Height:"
      Height          =   255
      Left            =   5520
      TabIndex        =   36
      Top             =   4110
      Width           =   495
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Width:"
      Height          =   255
      Left            =   5520
      TabIndex        =   33
      Top             =   3615
      Width           =   495
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Top:"
      Height          =   255
      Left            =   5640
      TabIndex        =   30
      Top             =   3135
      Width           =   375
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Left:"
      Height          =   255
      Left            =   5640
      TabIndex        =   27
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Line width:"
      Height          =   255
      Left            =   5160
      TabIndex        =   23
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Alignment:"
      Height          =   255
      Left            =   360
      TabIndex        =   22
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label11 
      Caption         =   "Line style:"
      Height          =   255
      Left            =   5280
      TabIndex        =   20
      Top             =   840
      Width           =   735
   End
   Begin VB.Label LineColorSample 
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   6480
      TabIndex        =   18
      Top             =   1770
      Width           =   315
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Line color:"
      Height          =   240
      Left            =   5175
      TabIndex        =   16
      Top             =   1800
      Width           =   840
   End
   Begin VB.Label FontColorSample 
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1560
      TabIndex        =   8
      Top             =   2280
      Width           =   315
   End
   Begin VB.Label Label3 
      Caption         =   "Font color:"
      Height          =   240
      Left            =   360
      TabIndex        =   6
      Top             =   2310
      Width           =   840
   End
   Begin VB.Label Label2 
      Caption         =   "Font, font size:"
      Height          =   255
      Left            =   135
      TabIndex        =   4
      Top             =   1800
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "The text: "
      Height          =   255
      Left            =   135
      TabIndex        =   1
      Top             =   150
      Width           =   1065
   End
End
Attribute VB_Name = "TextBoxProps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' This form collects the parameters needed to specify a test box that can
' be added to a PowerPoint slide. It gathers the parameters from a set of
' global data items that are defined in Module1. The MainForm initializes
' these values to begin with. Every time that this form is re-displayed by
' the show method the previously gathered values are restored as a starting
' point.
'
' Most of this is standard "run the controls" stuff that is little real
' interest.
Dim LocalX As Single
Dim LocalY As Single
Dim LineWidthFlag As Boolean
'
Sub InitLocalValues()
 '
 ' Make the controls reflect the various TbpTemp values.
 '
 Dim i, j As Integer
 Dim Trip As Boolean
 '
 ContentsTxt = TbpText(TbpTemp)
 '
 If TbpAlignment(TbpTemp) = ppAlignCenter Then
  AlignmentLst.ListIndex = 0
 ElseIf TbpAlignment(TbpTemp) = ppAlignLeft Then
  AlignmentLst.ListIndex = 1
 Else
  AlignmentLst.ListIndex = 2
 End If
 '
 Trip = False
 For i = 0 To FontSlctr1.ListCount - 1
  FontSlctr1.ListIndex = i
  If FontSlctr1.Text = TbpFontName(TbpTemp) Then
   Trip = True
   Exit For
  End If
 Next i
 If Trip = False Then FontSlctr1.ListIndex = 0
 '
 Trip = False
 For i = 0 To FontSize1.ListCount - 1
  FontSize1.ListIndex = i
  If FontSize1.Text = CStr(TbpFontSize(TbpTemp)) Then
   Trip = True
   Exit For
  End If
 Next i
 If Trip = False Then FontSlctr1.ListIndex = 0
 '
 FontColorSample.BackColor = TbpFontColor(TbpTemp)
 '
 If TbpFontNormal(TbpTemp) = True Then
  FontStyleOpt = True
  FontStyleChk(0).Value = vbUnchecked
  FontStyleChk(1).Value = vbUnchecked
  FontStyleChk(2).Value = vbUnchecked
  FontStyleChk(3).Value = vbUnchecked
  FontStyleChk(4).Value = vbUnchecked
 Else
  FontStyleOpt = False
  If TbpFontBold(TbpTemp) = True Then FontStyleChk(0).Value = vbChecked
  If TbpFontItalic(TbpTemp) = True Then FontStyleChk(1).Value = vbChecked
  If TbpFontUnderlined(TbpTemp) = True Then FontStyleChk(2).Value = vbChecked
  If TbpFontShadow(TbpTemp) = True Then FontStyleChk(3).Value = vbChecked
  If TbpFontEmbossed(TbpTemp) = True Then FontStyleChk(4).Value = vbChecked
 End If
 '
 Trip = False
 For i = 0 To LineWidthLst.ListCount - 1
  If LineWidthLst.List(i) = CStr(TbpLineWidth(TbpTemp)) Then
   LineWidthLst.ListIndex = i
   Trip = True
   Exit For
  End If
 Next i
 If Trip = False Then LineWidthLst.ListIndex = 0
 '
 If TbpLineStyle(TbpTemp) = msoLineSingle Then LineStyleLst.ListIndex = 0
 If TbpLineStyle(TbpTemp) = msoLineThickBetweenThin Then LineStyleLst.ListIndex = 1
 If TbpLineStyle(TbpTemp) = msoLineThickThin Then LineStyleLst.ListIndex = 2
 If TbpLineStyle(TbpTemp) = msoLineThinThick Then LineStyleLst.ListIndex = 3
 If TbpLineStyle(TbpTemp) = msoLineThinThin Then LineStyleLst.ListIndex = 4
 '
 LineColorSample.BackColor = TbpLineColor(TbpTemp)
 '
 LeftUpDown.Value = TbpLeft(TbpTemp)
 TopUpDown.Value = TbpTop(TbpTemp)
 WidthUpDown.Value = TbpWidth(TbpTemp)
 HeightUpDown.Value = TbpHeight(TbpTemp)
 '
 LocalX = SlideShape.Width / MainWidth
 LocalY = SlideShape.Height / MainHeight
 TextBoxShape.Left = SlideShape.Left + (TbpLeft(TbpTemp) * LocalX)
 TextBoxShape.Top = SlideShape.Top + (TbpTop(TbpTemp) * LocalY)
 TextBoxShape.Width = LocalX * TbpWidth(TbpTemp)
 TextBoxShape.Height = LocalY * TbpHeight(TbpTemp)
 '
End Sub
'
Private Sub AlignmentLst_Click()
 '
 If AlignmentLst.Text = "ppAlignCenter" Then
  TbpAlignment(TbpTemp) = ppAlignCenter
 ElseIf AlignmentLst.Text = "ppAlignLeft" Then
  TbpAlignment(TbpTemp) = ppAlignLeft
 Else
  TbpAlignment(TbpTemp) = ppAlignRight
 End If
 '
End Sub
'
Private Sub ContentsTxt_Change()
 '
 TbpText(TbpTemp) = ContentsTxt.Text
 '
End Sub

Private Sub DefaultBtn_Click()
 '
 TbpCurrentToTemp
 InitLocalValues
 '
End Sub
'
Private Sub DismissBtn_Click()
 '
 ' Inform the MainForm that the user wants to cancel out of the text
 ' box creation.
 '
 TbpFlag = False
 '
 ' Disappear until needed again.
 '
 Me.Hide
 '
End Sub
'
Private Sub FontColorBtn_Click()
 '
On Error GoTo ErrHandler
 '
 CommonDialog1.Flags = cdlCCFullOpen
 CommonDialog1.ShowColor
 FontColorSample.BackColor = CommonDialog1.Color
 TbpFontColor(TbpTemp) = CommonDialog1.Color
 Exit Sub
 '
ErrHandler:

End Sub
'
Private Sub FontSize1_Click()
 '
 'TbpFontSize(TbpTemp) = CInt(FontSize1.Text)
 '
End Sub
'
Private Sub FontSlctr1_Click()
 '
 'TbpFontName(TbpTemp) = FontSlctr1.Text
 '
End Sub
'
Private Sub FontStyleChk_Click(Index As Integer)
 '
 Dim i As Integer
 Dim Flag As Boolean
 '
 TbpFontNormal(TbpTemp) = False
 TbpFontBold(TbpTemp) = False
 TbpFontItalic(TbpTemp) = False
 TbpFontUnderlined(TbpTemp) = False
 TbpFontShadow(TbpTemp) = False
 TbpFontEmbossed(TbpTemp) = False
 Flag = False
 '
 For i = 0 To FontStyleChk.Count - 1
  If FontStyleChk(i).Value = vbChecked Then
   Flag = True
   If i = 0 Then TbpFontBold(TbpTemp) = True
   If i = 1 Then TbpFontItalic(TbpTemp) = True
   If i = 2 Then TbpFontUnderlined(TbpTemp) = True
   If i = 3 Then TbpFontShadow(TbpTemp) = True
   If i = 4 Then TbpFontEmbossed(TbpTemp) = True
  End If
 Next i
 FontStyleOpt.Value = Not Flag
 '
End Sub
'
Private Sub FontStyleOpt_Click()
 '
 If FontStyleOpt.Value = True Then
  TbpFontBold(TbpTemp) = False
  TbpFontItalic(TbpTemp) = False
  TbpFontUnderlined(TbpTemp) = False
  TbpFontShadow(TbpTemp) = False
  TbpFontEmbossed(TbpTemp) = False
  For i = 0 To FontStyleChk.Count - 1
   FontStyleChk(i).Value = vbUnchecked
  Next i
 End If
 '
End Sub
'
Private Sub Form_Activate()
 '
 ' The show method used in MainForm causes this event to be fired.
 '
 ' First we copy the current set of text box parameters to the temporary ones.
 '
 TbpCurrentToTemp
 '
 ' Now set up the controls so as to reflect these values.
 '
 InitLocalValues
 '
End Sub

Private Sub Form_Load()
 '
 ' When the form is load the controls are set up.
 '
 Dim i As Integer
 Dim j As Single
 '
 ' Populate the font name and font size combos for the lead
 ' slide options tab.
 '
 With FontSlctr1
  For i = 0 To Screen.FontCount - 1
   .AddItem Screen.Fonts(i)
  Next i
  ' Set ListIndex to 0.
  .ListIndex = 0
 End With
 '
 With FontSize1
  '
  ' Populate the combo with sizes in
  ' increments of 2.
  '
  For i = 8 To 72 Step 2
   .AddItem i
  Next i
  ' Set ListIndex to size 12.
  .ListIndex = 2
 End With
 '
 With AlignmentLst
  .AddItem "ppAlignCenter"
  .AddItem "ppAlignLeft"
  .AddItem "ppAlignRight"
  .ListIndex = 0
 End With
 '
 With LineStyleLst
  .AddItem "msoLineSingle"
  '.AddItem "msoLineStyleMixed"
  .AddItem "msoLineThickBetweenThin"
  .AddItem "msoLineThickThin"
  .AddItem "msoLineThinThick"
  .AddItem "msoLineThinThin"
  .ListIndex = 3
 End With
 '
 With LineWidthLst
  For j = 0.25 To 6 Step 0.25
   .AddItem CStr(j)
  Next j
  .ListIndex = 13
 End With
 '
 LeftUpDown.Max = MainWidth
 TopUpDown.Max = MainHeight
 WidthUpDown.Max = MainWidth
 HeightUpDown.Max = MainHeight
 If MainWidth >= MainHeight Then
  SlideShape.Width = GeometrySamplePic.ScaleWidth - 4
  SlideShape.Height = SlideShape.Width * MainHeight / MainWidth
 Else
  SlideShape.Height = GeometrySamplePic.Height - 4
  SlideShape.Width = SlideShape.Height * MainWidth / MainHeight
 End If
 SlideShape.Left = (GeometrySamplePic.ScaleWidth - SlideShape.Width) / 2
 SlideShape.Top = (GeometrySamplePic.ScaleHeight - SlideShape.Height) / 2
 '
End Sub
'
Private Sub HeightUpDown_Change()
 '
 TbpHeight(TbpTemp) = CSng(HeightUpDown.Value)
 TextBoxShape.Height = LocalY * TbpHeight(TbpTemp)
 '
End Sub
'
Private Sub LeftUpDown_Change()
 '
 TbpLeft(TbpTemp) = CSng(LeftUpDown.Value)
 TextBoxShape.Left = SlideShape.Left + (TbpLeft(TbpTemp) * LocalX)
 '
End Sub
'
Private Sub LineColorBtn_Click()
 '
On Error GoTo ErrHandler
 '
 CommonDialog1.Flags = cdlCCFullOpen
 CommonDialog1.ShowColor
 LineColorSample.BackColor = CommonDialog1.Color
 TbpLineColor(TbpTemp) = CommonDialog1.Color
 Exit Sub
 '
ErrHandler:
 '
End Sub
'
Private Sub LineStyleLst_Click()
 '
 If LineStyleLst.Text = "msoLineSingle" Then
  TbpLineStyle(TbpTemp) = msoLineSingle
 'ElseIf LineStyleLst.Text = "msoLineStyleMixed" Then
 ' TbpLineStyle(TbpTemp) = msoLineStyleMixed
 ElseIf LineStyleLst.Text = "msoLineThickBetweenThin" Then
  TbpLineStyle(TbpTemp) = msoLineThickBetweenThin
 ElseIf LineStyleLst.Text = "msoLineThickThin" Then
  TbpLineStyle(TbpTemp) = msoLineThickThin
 ElseIf LineStyleLst.Text = "msoLineThinThick" Then
  TbpLineStyle(TbpTemp) = msoLineThinThick
 Else
  TbpLineStyle(TbpTemp) = msoLineThinThin
 End If
 '
End Sub
'
Private Sub LineWidthLst_Click()
 '
 TbpLineWidth(TbpTemp) = CSng(LineWidthLst.Text)
 '
End Sub
'
Private Sub OkBtn_Click()
 '
 ' The next two lines are candidates for clean up.
 '
 TbpFontName(TbpTemp) = FontSlctr1.Text
 TbpFontSize(TbpTemp) = CInt(FontSize1.Text)
 '
 ' Make the temporary working variable the current working set.
 '
 TbpTempToCurrent
 '
 ' Disappear until needed again.
 '
 Me.Hide
 '
End Sub
'
Private Sub TopUpDown_Change()
 '
 TbpTop(TbpTemp) = CSng(TopUpDown.Value)
 TextBoxShape.Top = SlideShape.Top + (TbpTop(TbpTemp) * LocalY)
 '
End Sub
'
Private Sub WidthUpDown_Change()
 '
 TbpWidth(TbpTemp) = CSng(WidthUpDown.Value)
 TextBoxShape.Width = LocalX * TbpWidth(TbpTemp)
 '
End Sub
