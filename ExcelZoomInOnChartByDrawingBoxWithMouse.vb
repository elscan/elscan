'last updated May 16/2021
'This is to be pasted into an excel visual basic module
'Open YourFile.xlsm and navigate to developer->visual basic
'A window should pop up with the label     Microsoft Visual Basic for Applications - YourFile.xlsm
'Within that window, navigate to VBAProject(YourFile.xlsm)->Microsoft Excel Objects->Sheet1(Sheet1)
'Clicking on Sheet1(Sheet1) should open

'PREREQUISITE: you need to know how to make an ActiveX form control, make it transparent, and place it over your plot area
'PREREQUISITE: you need to be able to go into developer mode, right-click on your ActiveX label, and then click properties. Toward the top-left of properties, you should see 'Label1' in bold. Otherwise, you need to change all instances of 'Label1' in this code to whatever you see in bold toward the top left of your ActiveX label properties

'BIG BUG warning: if you mouse-move over the ActiveX label before the chart becomes active, the program will crash and you will need to reset (see the stop-button in the toolbar at the top)
'                 if you have multiple charts and the active chart is not the one with the ActiveX label, nothing will make sense


'-v- First, we declare some variables.
'labelHeight and labelWidth have units of points
Dim labelHeight As Single, labelWidth As Single ' !The user needs to put this label in! via Developer->Insert->Label(ActiveX control) . Then, you need to make it transparent via right-click->Properties->BackStyle
                                                ' !You need to make sure that your ActiveX label is Label1 (right-click on the ActiveX label and then click Properties. You should see Label1 in bold toward top-left)
                                                ' !If your ActiveX label is not Label1 then you need to find and replace all Label1 with whatever you see in the AxtiveX label properties
                                                ' !Careful, the characters (l) and (1) look quite similar in most text editors

'chartX and clickX variables are NOT in points. They are (should be) in the same units of your chart's x-axis
'See Private Sub updateDimensions(X,Y). X and Y are in points, but get converted to your chart's units in this subroutine
Dim chartXmin As Single, chartYmin As Single, chartXmax As Single, chartYmax As Single
Dim click1X As Single, click1Y As Single, click2X As Single, click2Y As Single
Dim chartX As Single, chartY As Single

Dim sq As shape, shapeIter As shape 'NOTE: sq is a RECTANGLE.
Dim sqExists As Boolean             'NOTE: sq is a RECTANGLE.
'-^- First, we declare some variables.

Private Sub Label1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'X and Y are measured in points
  Call updateDimensions(X, Y)
  'Call upDateMouseStuff(X, Y)
  Call lookForsq
  If Not sqExists Then Call makeSq
  If sqExists Then Call updateSq(X, Y)
End Sub

Private Sub Label1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Label1.BackStyle = fmBackStyleTransparent 'I had to do all this style and visibility toggling so that the ActiveX label did not block the plot area
  Label1.Visible = False  'I had to do all this style and visibility toggling so that the ActiveX label did not block the plot area
  Label1.Visible = True   'I had to do all this style and visibility toggling so that the ActiveX label did not block the plot area
  
  '-v-******* Zooming with a rectangle
  'If you ctrl+left-click, the top left of the rectangle is defined at the click spot
  'If you left-click, the bottom right of the rectangle is defined at the click spot and the chart zooms (approximately) to the rectangle
  'If you right-click, the chart is auto-scaled
  'Note: if you left-click before ctrl+left-click, the chart zoom will be wierd and you have to right-click to auto-zoom again
  If Button = 1 Then
    If Shift = 2 Then
      sq.Left = X + Label1.Left
      sq.Top = Y + Label1.Top
      sq.Width = 0
      sq.Height = 0
      sq.Visible = True
      click1X = chartX
      click1Y = chartY
    Else
      click2X = chartX
      click2Y = chartY
      sq.Visible = False
      Call updateAxes  'This is the subroutine that adjusts the axes scale based on the top-left and bottom-right coordinates of the rectangle. This zooms in (approximately) to the rectangle
      Call updateLabel 'When you zoom on the chart, the plot area is changed. So the ActiveX label needs to be resized
    End If
  End If
  If Button = 2 Then
    ActiveChart.Axes(xlCategory).MinimumScaleIsAuto = True
    ActiveChart.Axes(xlCategory).MaximumScaleIsAuto = True
    ActiveChart.Axes(xlValue).MinimumScaleIsAuto = True
    ActiveChart.Axes(xlValue).MaximumScaleIsAuto = True
    Call updateLabel
  End If
'*-^-*******Zooming with a rectangle
  sq.ZOrder msoBringToFront 'I want to make sure that out ActiveX label and rectangle are in front of the active chart
  Label1.BringToFront       'I want to make sure that out ActiveX label and rectangle are in front of the active chart
End Sub
Private Sub updateAxes() 'This is the subroutine that adjusts the axes scale based on the top-left and bottom-right coordinates of the rectangle. This zooms in (approximately) to the rectangle
  ActiveChart.Axes(xlCategory).MinimumScale = click1X - 0.02 * Abs(click2X - click1X)
  ActiveChart.Axes(xlCategory).MaximumScale = click2X + 0.02 * Abs(click2X - click1X)
  ActiveChart.Axes(xlValue).MinimumScale = click2Y - 0.02 * Abs(click1Y - click2Y)
  ActiveChart.Axes(xlValue).MaximumScale = click1Y + 0.02 * Abs(click1Y - click2Y)
End Sub
Private Sub updateDimensions(X, Y) 'We need to make sure we know our most-up-to-date plot area dimensions so we can correctly determine our current x,y-coordinate, relative to the top-left corner of the plot area
  labelWidth = Label1.Width   'Possibility for bug: looks like we the Label must be updated before Dimensions are updated. If this is not the case, the program will bug
  labelHeight = Label1.Height
  
  chartXmin = CSng(ActiveChart.Axes(xlCategory).MinimumScale) 'Recall that CSng(anyObject) will generate a copy (anyObject) as a Single (if can't recall, find docs)
  chartXmax = CSng(ActiveChart.Axes(xlCategory).MaximumScale)
  chartYmin = CSng(ActiveChart.Axes(xlValue).MinimumScale)
  chartYmax = CSng(ActiveChart.Axes(xlValue).MaximumScale)
  
  chartX = chartXmin + X * (chartXmax - chartXmin) / labelWidth 'This is where the mouse coordinates (measured in points relative to the top-left of the ActiveX label) are converted to the coordinates on your chart
  chartY = chartYmin + (labelHeight - Y) * (chartYmax - chartYmin) / labelHeight
End Sub
Private Sub updateLabel() 'When you zoom on the chart, the PlotArea is changed. So the ActiveX label needs to be resized
'This method is actually inaccurate. If you go into Design Mode in excel, you can see that the center of the ActiveX label is up and to the left of where it should be
'All of these variables have units of points (if points is unfamiliar, find docs)
  Label1.Left = ActiveChart.ChartArea.Left + ActiveChart.PlotArea.InsideLeft
  Label1.Top = ActiveChart.ChartArea.Top + ActiveChart.PlotArea.InsideTop
  Label1.Width = ActiveChart.PlotArea.InsideWidth
  Label1.Height = ActiveChart.PlotArea.InsideHeight
End Sub
Private Sub lookForsq()
  sqExists = False
  For Each shapeIter In ActiveSheet.Shapes 'I could have just said [For Each shape_ In ActiveSheet.Shapes]. The language is nicer, but I like to reclaa that we are making an index, or an iterator that goes through the list of shapes in the sheet
    If shapeIter.Name = "mysq" Then sqExists = True    'I am sure that this line and the next could be combined into one line. Apologies
    If shapeIter.Name = "mysq" Then Set sq = shapeIter
  Next shapeIter
End Sub
Private Sub makeSq()
  Set sq = ActiveSheet.Shapes.AddShape(1, 10, 10, 10, 10) 'I just put the rectandle in an arbitrary location. It updates right away when you start drawing the box on the chart
  sq.Name = "mysq"
  sq.Fill.Transparency = 1#
End Sub
Private Sub updateSq(X, Y)
 sq.Width = Abs(X + Label1.Left - sq.Left)
 sq.Height = Abs(Y + Label1.Top - sq.Top)
End Sub
Private Sub upDateMouseStuff(X, Y)
  Range("R3").Value = "Mouse X,Y"
  Range("S3").Value = X
  Range("T3").Value = Y
  Range("R4").Value = "Width,Height"
  Range("S4").Value = labelWidth
  Range("T4").Value = labelHeight
  Range("R5").Value = "xMin,xMax"
  Range("S5").Value = chartXmin
  Range("T5").Value = chartXmax
  Range("R6").Value = "yMin,yMax"
  Range("S6").Value = chartYmin
  Range("T6").Value = chartYmax
  Range("R7").Value = "Chartx,y"
  Range("S7").Value = chartX
  Range("T7").Value = chartY
End Sub
