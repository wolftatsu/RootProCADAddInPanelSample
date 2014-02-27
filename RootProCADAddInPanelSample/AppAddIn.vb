Imports System
Imports System.Collections.Generic
Imports System.Windows.Forms
Imports RootPro.RootProCAD
Imports RootPro.RootProCAD.Command
Imports RootPro.RootProCAD.Geometry
Imports RootPro.RootProCAD.UI
<System.AddIn.AddIn("AppAddIn", Version:="1.0", Publisher:="", Description:="")> _
Partial Class AppAddIn
    ' panelサイズ
    Const width As Double = 1660.0
    Const height As Double = 4150.0
    Private Sub AppAddIn_Startup(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Startup
        CommandManager.AddMacroCommand("太陽光パネル作成", AddressOf Me.MacroCommand)
        CommandManager.AddMacroCommand("パネル枠線作成", AddressOf Me.MacroCommand2)
    End Sub

    Private Sub AppAddIn_Shutdown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shutdown
        CommandManager.RemoveMacroCommand(AddressOf Me.MacroCommand)
        CommandManager.RemoveMacroCommand(AddressOf Me.MacroCommand2)
    End Sub
    ' パネル作成
    Private Sub MacroCommand()
        On Error Resume Next
        Dim doc As Document = ActiveDocument

        Dim drawing As Drawing = doc.CurrentDrawing
        Dim points(1) As Point2d

        ' "format(x, y)"
        Dim startPoint As String = InputBox("開始座標")
        Dim pointsStr As String = Split(startPoint, ",")(0)
        Dim panelCount As Integer = CInt(InputBox("パネル列数"))

        Dim startX As Double = Double.Parse(Split(startPoint, ",")(0))
        Dim startY As Double = Double.Parse(Split(startPoint, ",")(1))

        doc.UndoManager.BeginUndoUnit()

        Dim currentX As Double = startX
        Dim currentY As Double = startY

        Dim panelCounter As Integer = panelCount
        Dim creator As PanelCreator = New PanelCreator(drawing, Geometry, panelCounter, currentX, currentY, doc.SelectionManager)

        creator.run(20)
        creator.run(10)
        creator.run(5)
        creator.run(1)

        doc.UndoManager.EndUndoUnit()

    End Sub
    ' パネル枠線作成
    Private Sub MacroCommand2()
        On Error Resume Next
        Dim doc As Document = ActiveDocument

        Dim drawing As Drawing = doc.CurrentDrawing
        Dim points(1) As Point2d
        Dim creator As PanelCreator = New PanelCreator(drawing, Geometry, 0, 0, 0, doc.SelectionManager)
        creator.writePanelLine()
    End Sub
    ' 補助線に沿ってめいいっぱいパネル配置
    Private Sub MacroCommand3()
        On Error Resume Next
        Dim doc As Document = ActiveDocument

        Dim drawing As Drawing = doc.CurrentDrawing
        Dim points(1) As Point2d
        Dim creator As PanelCreator = New PanelCreator(drawing, Geometry, 0, 0, 0, doc.SelectionManager)
        creator.createPanelsFully()

    End Sub



    Private Class PanelCreator
        Dim currentX As Double
        Dim currentY As Double
        Dim drawing As Drawing
        Dim Geometry As Geometry
        Dim selectinoManager As SelectionManager
        Public panelCounter As Integer
        Public angle As Integer = 0
        ' panelサイズ
        Const width As Double = 1660.0
        Const height As Double = 4150.0
        Const charaH As Double = 2500
        Public Sub New(ByVal drawing As Drawing, ByVal geometry As Geometry, ByVal panelCounter As Integer, ByVal currentX As Double, ByVal currentY As Double, ByVal selectionManager As SelectionManager)
            Me.currentX = currentX
            Me.currentY = currentY
            Me.panelCounter = panelCounter
            Me.drawing = drawing
            Me.Geometry = geometry
            Me.selectinoManager = selectionManager

        End Sub
        ' パネル配置
        Public Sub run(ByVal panelGroupCount As Integer)
            Dim groupCount As Integer = panelCounter \ panelGroupCount
            Dim backupX As Double = currentX
            Dim textShape As TextShape
            Dim panel As PolylineShape

            If groupCount > 0 Then
                For i = 1 To groupCount
                    panel = Drawing.Shapes.AddPolyline(createAPanel(currentX, currentY, panelGroupCount))
                    panel.ColorNumber = 6
                    panel.LinewidthNumber = 3
                    panel.LinetypeNumber = 1
                    'panel.Rotate(Geometry.CreatePoint(currentX, currentY), angle)
                    ' パネル選択
                    Me.selectinoManager.SelectShape(panel, SelectionCombineMode.Add)

                    If panelGroupCount > 4 Then
                        Dim tmpX As Double = currentX + ((width * panelGroupCount) / 2)
                        Dim textPoint As Point2d = Geometry.CreatePoint(currentX, currentY - 1000)
                        textShape = Drawing.Shapes.AddText(CStr(panelGroupCount) + "列", textPoint, 0)
                        textShape.FontHeight = charaH
                        'textShape.Rotate(textPoint, angle)
                        ' 文字選択
                        Me.selectinoManager.SelectShape(textShape, SelectionCombineMode.Add)
                    End If
                    currentX = currentX + (width * panelGroupCount)
                Next
                panelCounter = panelCounter - (panelGroupCount * groupCount)

            End If

        End Sub
        ' パネル配置座標を返す（連続線分の座標配列）
        Public Function createAPanel(ByVal x As Double, ByVal y As Double, ByVal panelCount As Integer) As Point2d()

            Dim panelPoints(4) As Point2d

            panelPoints(0) = Geometry.CreatePoint(x, y)
            panelPoints(1) = Geometry.CreatePoint(x + (width * panelCount), y)
            panelPoints(2) = Geometry.CreatePoint(x + (width * panelCount), y - height)
            panelPoints(3) = Geometry.CreatePoint(x, y - height)
            panelPoints(4) = Geometry.CreatePoint(x, y)
            createAPanel = panelPoints

        End Function
        ' パネル配置ガイド線の上部50xmの所に境界の線を引く
        Public Sub writePanelLine()
            Dim shape As SelectedShape = Me.selectinoManager.SelectedShapes.Item(0)
            Dim points() As Point2d = shape.Shape.GetBoundingPoints()


            Dim firstPoint As Point2d = points.GetValue(0)
            Dim panelLine As PolylineShape

            Dim linePoints(1) As Point2d
            Dim i As Integer = 0
            Dim lastPoint As Point2d = firstPoint

            firstPoint = getFirstPoint(points)
            lastPoint = getlastPoint(points)

            linePoints(0) = Geometry.CreatePoint(firstPoint.X, firstPoint.Y + 500)
            linePoints(1) = Geometry.CreatePoint(lastPoint.X, lastPoint.Y + 500)

            panelLine = Drawing.Shapes.AddPolyline(linePoints)
            panelLine.LinetypeNumber = 1
            panelLine.ColorNumber = 4
            panelLine.LinewidthNumber = 3

        End Sub
        Public Sub createPanelsFully()




        End Sub
        Private Function getFirstPoint(ByVal points() As Point2d) As Point2d
            Dim firstPoint As Point2d = points.GetValue(0)
            For Each p In points
                If firstPoint.X > p.X Or (firstPoint.X = p.X And firstPoint.Y < p.Y) Then
                    firstPoint = p
                End If
            Next
            getFirstPoint = firstPoint
        End Function
        Private Function getlastPoint(ByVal points() As Point2d) As Point2d
            Dim lastPoint As Point2d = points.GetValue(0)
            For Each p In points
                If lastPoint.X < p.X Or (lastPoint.X = p.X And lastPoint.Y < p.Y) Then
                    lastPoint = p
                End If
            Next
            getlastPoint = lastPoint
        End Function



    End Class

End Class

