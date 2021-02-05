VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub resize_charts()
'resizes charts to the same size/setout, change as required
    Dim objchart_M As Object
    Dim objchart_V As Object
    Dim objchart_D As Object
    Set objchart_M = Sheet1.ChartObjects("moment_chart")
    Set objchart_V = Sheet1.ChartObjects("shear_chart")
    Set objchart_D = Sheet1.ChartObjects("defl_chart")

    Dim left As Long
    Dim top As Long
    Dim height As Long
    Dim width As Long
    Dim label_left As Long
    Dim label_mid As Long

    left = 31
    top = 7
    height = 150
    width = 350
    label_left = -2
    label_mid = 82

    With objchart_M.Chart.PlotArea
        .InsideLeft = left
        .InsideTop = top
        .InsideHeight = height
        .InsideWidth = width
    End With

    With objchart_M.Chart.Axes(xlValue).AxisTitle
        .left = label_left
        .top = label_mid - .height / 2
    End With

    With objchart_V.Chart.PlotArea
        .InsideLeft = left
        .InsideTop = top
        .InsideHeight = height
        .InsideWidth = width
    End With

    With objchart_V.Chart.Axes(xlValue).AxisTitle
        .left = label_left
        .top = label_mid - .height / 2
    End With

    With objchart_D.Chart.PlotArea
        .InsideLeft = left
        .InsideTop = top
        .InsideHeight = height
        .InsideWidth = width
    End With

    With objchart_D.Chart.Axes(xlValue).AxisTitle
        .left = label_left
        .top = label_mid - .height / 2
    End With

    'cleanup
    Set objchart_M = Nothing
    Set objchart_V = Nothing
    Set objchart_D = Nothing

End Sub
