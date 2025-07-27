Attribute VB_Name = "NewMacros"
Sub DoAll()

SetMargins
DeleteCustomStyles
ResetCommonProperties

SetStyleNormal
SetStyleHeading
SetStyleListNumber
SetStyleListBullet

CleanListGalleries

SetOutlineNumberedListsGallery
SetNumberedListsGallery
SetBulletedListsGallery

End Sub

Sub SetMargins()
With ActiveDocument.PageSetup
 '.Orientation = wdOrientPortrait '0
 .LeftMargin = MillimetersToPoints(25) '70.85
 .RightMargin = MillimetersToPoints(15) '42.5
 .TopMargin = MillimetersToPoints(15) '42.5
 .BottomMargin = MillimetersToPoints(15) '42.5
End With

'Report
MsgBox "Margins Done", vbInformation
End Sub

Sub DeleteCustomStyles()
n = ActiveDocument.Styles.Count
'For...Next cycle is not suitable because .Styles.Count changes
'For Each cycle which is suggested by microsoft is also not suitable
i = 0
lblDeleteStyle:
 i = i + 1
 If ActiveDocument.Styles(i).BuiltIn = False Then
  ActiveDocument.Styles(i).Delete
  i = i - 1 'Now previous position is occupied by the next style
 End If
 If i < ActiveDocument.Styles.Count Then
    GoTo lblDeleteStyle
 End If

'Report
n = n - i
MsgBox n & " Custom Styles Deleted", vbInformation
End Sub

Sub ResetCommonProperties()
'On Error Resume Next
n = ActiveDocument.Styles.Count
For i = 1 To n
 Set theStyle = ActiveDocument.Styles(i)
 With theStyle.Font
 .Name = "Times New Roman"
 '.Size = 12
 .ColorIndex = wdAuto
 .Scaling = 100
 .Spacing = 0
 End With
 If theStyle.Type = wdStyleTypeParagraph _
 Or theStyle.Type = wdStyleTypeParagraphOnly Then
  With theStyle.ParagraphFormat
  .TabStops.ClearAll
  .FirstLineIndent = 0
  .LeftIndent = 0
  .LineSpacingRule = wdLineSpaceSingle
  .Space1
  .SpaceBefore = 0
  .SpaceAfter = 0
  End With
 End If
Next i

'Report
MsgBox "Common Properties Done", vbInformation
End Sub

Sub SetStyleNormal()
Set theStyle = ActiveDocument.Styles(wdStyleNormal)
With theStyle
 With .ParagraphFormat
 .Alignment = wdAlignParagraphJustify
 .FirstLineIndent = MillimetersToPoints(10)
 .LeftIndent = 0
 End With
End With

Set theStyle = ActiveDocument.Styles(wdStyleHeader)
With theStyle
 .Font.Bold = False
 .Font.Size = 10
 With theStyle.ParagraphFormat
 .Alignment = wdAlignParagraphRight
 End With
End With

Set theStyle = ActiveDocument.Styles(wdStyleFooter)
With theStyle
 .Font.Bold = False
 .Font.Size = 10
 With theStyle.ParagraphFormat
 .Alignment = wdAlignParagraphRight
 End With
End With

'Report
MsgBox "Normal Style Done", vbInformation
End Sub

Sub SetStyleHeading()

Set theStyle = ActiveDocument.Styles(wdStyleHeading1)
With theStyle
 .Font.Bold = True
 With theStyle.ParagraphFormat
 .Alignment = wdAlignParagraphCenter
 .FirstLineIndent = 0
 .LeftIndent = 0
 .SpaceBefore = 12
 .SpaceAfter = 6
 .KeepWithNext = True
 End With
End With

Set theStyle = ActiveDocument.Styles(wdStyleHeading2)
With theStyle
 .Font.Bold = True
 With theStyle.ParagraphFormat
 .Alignment = wdAlignParagraphCenter
 .FirstLineIndent = 0
 .LeftIndent = 0
 .SpaceBefore = 12
 .SpaceAfter = 6
 .KeepWithNext = True
 End With
End With

Set theStyle = ActiveDocument.Styles(wdStyleHeading3)
With theStyle
 .Font.Bold = True
 With theStyle.ParagraphFormat
 .Alignment = wdAlignParagraphCenter
 .FirstLineIndent = 0
 .LeftIndent = 0
 .SpaceBefore = 12
 .SpaceAfter = 6
 .KeepWithNext = True
 End With
End With

Set theStyle = ActiveDocument.Styles(wdStyleHeading4)
With theStyle
 .Font.Bold = True
 With theStyle.ParagraphFormat
 .Alignment = wdAlignParagraphCenter
 .FirstLineIndent = 0
 .LeftIndent = 0
 .SpaceBefore = 12
 .SpaceAfter = 6
 .KeepWithNext = True
 End With
End With

'Report
MsgBox "Heading Styles Done", vbInformation
End Sub

Sub SetStyleListNumber()

Set theStyle = ActiveDocument.Styles(wdStyleListNumber)
With theStyle.ParagraphFormat
 .Alignment = wdAlignParagraphJustify
 '.TabStops(MillimetersToPoints(10)).Alignment = wdAlignTabLeft
 .FirstLineIndent = MillimetersToPoints(-10)
 .LeftIndent = MillimetersToPoints(10)
End With

Set theStyle = ActiveDocument.Styles(wdStyleListNumber2)
With theStyle.ParagraphFormat
 .Alignment = wdAlignParagraphJustify
 '.TabStops(MillimetersToPoints(10)).Alignment = wdAlignTabLeft
 .FirstLineIndent = MillimetersToPoints(-10)
 .LeftIndent = MillimetersToPoints(10)
End With

Set theStyle = ActiveDocument.Styles(wdStyleListNumber3)
With theStyle.ParagraphFormat
 .Alignment = wdAlignParagraphJustify
 '.TabStops(MillimetersToPoints(10)).Alignment = wdAlignTabLeft
 .FirstLineIndent = MillimetersToPoints(-10)
 .LeftIndent = MillimetersToPoints(10)
End With

Set theStyle = ActiveDocument.Styles(wdStyleListNumber4)
With theStyle.ParagraphFormat
 .Alignment = wdAlignParagraphJustify
 '.TabStops(MillimetersToPoints(10)).Alignment = wdAlignTabLeft
 .FirstLineIndent = MillimetersToPoints(-10)
 .LeftIndent = MillimetersToPoints(10)
End With

Set theStyle = ActiveDocument.Styles(wdStyleListNumber5)
With theStyle.ParagraphFormat
 .Alignment = wdAlignParagraphJustify
 '.TabStops(MillimetersToPoints(10)).Alignment = wdAlignTabLeft
 .FirstLineIndent = MillimetersToPoints(-10)
 .LeftIndent = MillimetersToPoints(10)
End With

'Report
MsgBox "Numbered List Styles Done", vbInformation
End Sub

Sub SetStyleListBullet()
Dim theStyle As style

Set theStyle = ActiveDocument.Styles(wdStyleListBullet)
With theStyle.ParagraphFormat
 .Alignment = wdAlignParagraphJustify
 '.TabStops(MillimetersToPoints(10)).Alignment = wdAlignTabLeft
 .FirstLineIndent = MillimetersToPoints(-10)
 .LeftIndent = MillimetersToPoints(10)
End With

Set theStyle = ActiveDocument.Styles(wdStyleListBullet2)
With theStyle.ParagraphFormat
 .Alignment = wdAlignParagraphJustify
 '.TabStops(MillimetersToPoints(10)).Alignment = wdAlignTabLeft
 .FirstLineIndent = MillimetersToPoints(-10)
 .LeftIndent = MillimetersToPoints(10)
End With

Set theStyle = ActiveDocument.Styles(wdStyleListBullet3)
With theStyle.ParagraphFormat
 .Alignment = wdAlignParagraphJustify
 '.TabStops(MillimetersToPoints(10)).Alignment = wdAlignTabLeft
 .FirstLineIndent = MillimetersToPoints(-10)
 .LeftIndent = MillimetersToPoints(10)
End With

Set theStyle = ActiveDocument.Styles(wdStyleListBullet4)
With theStyle.ParagraphFormat
 .Alignment = wdAlignParagraphJustify
 '.TabStops(MillimetersToPoints(10)).Alignment = wdAlignTabLeft
 .FirstLineIndent = MillimetersToPoints(-10)
 .LeftIndent = MillimetersToPoints(10)
End With

Set theStyle = ActiveDocument.Styles(wdStyleListBullet5)
With theStyle.ParagraphFormat
 .Alignment = wdAlignParagraphJustify
 '.TabStops(MillimetersToPoints(10)).Alignment = wdAlignTabLeft
 .FirstLineIndent = MillimetersToPoints(-10)
 .LeftIndent = MillimetersToPoints(10)
End With

'Report
MsgBox "Bulleted List Styles Done", vbInformation
End Sub

Sub CleanListGalleries()
Dim lg As ListGallery
For Each lg In ListGalleries
 For i = 1 To 7
 lg.Reset (i)
 Next i
Next lg
End Sub

Sub SetOutlineNumberedListsGallery()

For i = 1 To 7
  For j = 1 To 9
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(i).ListLevels(j)
        .TrailingCharacter = wdTrailingNone
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TabPosition = wdUndefined
        .StartAt = 1
        With .Font
            .Bold = False
            .Italic = False
            .StrikeThrough = False
            .Subscript = False
            .Superscript = False
            .Shadow = False
            .Outline = False
            .Emboss = False
            .Engrave = False
            .AllCaps = False
            .Hidden = False
            .Underline = False
            .ColorIndex = wdAuto
            .Size = 12
            .Animation = False
            .DoubleStrikeThrough = False
            '.Name = ""
        End With
    End With
  Next j
Next i

With ListGalleries(wdOutlineNumberGallery).ListTemplates(1)
    With .ListLevels(1)
        .LinkedStyle = "Нумерованный список" 'wdStyleListNumber
        .NumberFormat = "%1." & Chr(160)
        .TextPosition = CentimetersToPoints(1)
        .ResetOnHigher = 0
    End With
    With .ListLevels(2)
        .LinkedStyle = "Нумерованный список 2" 'wdStyleListNumber2
        .NumberFormat = "%1.%2." & Chr(160)
        .TextPosition = CentimetersToPoints(1)
        .ResetOnHigher = 1
    End With
    With .ListLevels(3)
        .LinkedStyle = "Нумерованный список 3" 'wdStyleListNumber3
        .NumberFormat = "%1.%2.%3." & Chr(160)
        .TextPosition = CentimetersToPoints(1)
        .ResetOnHigher = 2
    End With
    With .ListLevels(4)
        .LinkedStyle = "Нумерованный список 4" 'wdStyleListNumber4
        .NumberFormat = "%1.%2.%3.%4." & Chr(160)
        .TextPosition = CentimetersToPoints(1.5)
        .ResetOnHigher = 3
    End With
    With .ListLevels(5)
        .LinkedStyle = "Нумерованный список 5" 'wdStyleListNumber5
        .NumberFormat = "%1.%2.%3.%4.%5." & Chr(160)
        .TextPosition = CentimetersToPoints(1.5)
        .ResetOnHigher = 4
    End With
   With .ListLevels(6)
       .LinkedStyle = "Нумерованный список 6" 'wdStyleListNumber6
       .NumberFormat = "%1.%2.%3.%4.%5.%6." & Chr(160)
       .TextPosition = CentimetersToPoints(1.5)
       .ResetOnHigher = 5
   End With
   With .ListLevels(7)
       .LinkedStyle = "Нумерованный список 7" 'wdStyleListNumber7
       .NumberFormat = "%1.%2.%3.%4.%5.%6.%7." & Chr(160)
       .TextPosition = CentimetersToPoints(1.5)
       .ResetOnHigher = 6
   End With
   With .ListLevels(8)
       .LinkedStyle = "Нумерованный список 8" 'wdStyleListNumber7
       .NumberFormat = "%1.%2.%3.%4.%5.%6.%7.%8." & Chr(160)
       .TextPosition = CentimetersToPoints(2)
       .ResetOnHigher = 7
   End With
    With .ListLevels(9)
       .LinkedStyle = "Нумерованный список 9" 'wdStyleListNumber7
       .NumberFormat = "%1.%2.%3.%4.%5.%6.%7.%8.%9." & Chr(160)
       .TextPosition = CentimetersToPoints(2)
       .ResetOnHigher = 8
   End With
End With
    
ListGalleries(wdOutlineNumberGallery).ListTemplates(1).Name = "ListTemplateWombat"

Selection.Range.ListFormat.ApplyListTemplateWithLevel ListTemplate:= _
    ListGalleries(wdOutlineNumberGallery).ListTemplates(1), _
    ContinuePreviousList:=False, ApplyTo:=wdListApplyToWholeList, _
    DefaultListBehavior:=wdWord10ListBehavior

'Report
MsgBox "Outline Numbered Lists Gallery Done", vbInformation
End Sub

Sub SetNumberedListsGallery()
m = ListGalleries(wdNumberGallery).ListTemplates.Count
For i = 1 To m
 n = ListGalleries(wdNumberGallery).ListTemplates(i).ListLevels.Count
 For j = 1 To n
  With ListGalleries(wdNumberGallery).ListTemplates(i).ListLevels(j)
  .NumberStyle = wdListNumberStyleArabic
  .NumberPosition = 0
  .TextPosition = 0
  .NumberFormat = "%1."
  .TrailingCharacter = wdTrailingTab
  .TabPosition = MillimetersToPoints(10)
  End With
 Next j
Next i
'Report
MsgBox "Numbered Lists Gallery Done", vbInformation
End Sub

Sub SetBulletedListsGallery()

m = ListGalleries(wdBulletGallery).ListTemplates.Count
For i = 1 To m
 n = ListGalleries(wdBulletGallery).ListTemplates(i).ListLevels.Count
 For j = 1 To n
  With ListGalleries(wdBulletGallery).ListTemplates(i).ListLevels(j)
  .NumberPosition = 0
  .TextPosition = 0
  .TrailingCharacter = wdTrailingTab
  .TabPosition = MillimetersToPoints(10)
  End With
 Next j
Next i

'Report
MsgBox "Bulleted Lists Gallery Done", vbInformation
End Sub

Sub InsertTextAtEndOfDocument()
ActiveDocument.Content.InsertAfter Text:=Chr(13)
 
m = ListGalleries(wdOutlineNumberGallery).ListTemplates.Count
For i = 1 To m
 ActiveDocument.Content.InsertAfter Text:=Chr(13) & i & ListGalleries(wdOutlineNumberGallery).ListTemplates(i).Name
 n = ListGalleries(wdOutlineNumberGallery).ListTemplates(i).ListLevels.Count
 For j = 1 To n
  ActiveDocument.Content.InsertAfter Text:=Chr(13) & i & ListGalleries(wdOutlineNumberGallery).ListTemplates(i).ListLevels(j).LinkedStyle
 Next j
Next i
 
End Sub

