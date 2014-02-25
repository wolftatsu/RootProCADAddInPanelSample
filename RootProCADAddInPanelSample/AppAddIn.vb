Imports System
Imports System.Collections.Generic
Imports System.Windows.Forms
Imports RootPro.RootProCAD
Imports RootPro.RootProCAD.Command
Imports RootPro.RootProCAD.Geometry
Imports RootPro.RootProCAD.UI
<System.AddIn.AddIn("AppAddIn", Version:="1.0", Publisher:="", Description:="")> _
Partial Class AppAddIn
    ' panel�T�C�Y
    Const width As Double = 1660.0
    Const height As Double = 4150.0
    Private Sub AppAddIn_Startup(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Startup
        CommandManager.AddMacroCommand("���z���p�l���쐬", AddressOf Me.MacroCommand)
        CommandManager.AddMacroCommand("�p�l���g���쐬", AddressOf Me.MacroCommand2)
    End Sub

    Private Sub AppAddIn_Shutdown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shutdown
        CommandManager.RemoveMacroCommand(AddressOf Me.MacroCommand)
        CommandManager.RemoveMacroCommand(AddressOf Me.MacroCommand2)
    End Sub
    ' �p�l���쐬
    Private Sub MacroCommand()
        On Error Resume Next
        Dim doc As Document = ActiveDocument

        Dim drawing As Drawing = doc.CurrentDrawing
        Dim points(1) As Point2d

        ' "format(x, y)"
        Dim startPoint As String = InputBox("�J�n���W")
        Dim pointsStr As String = Split(startPoint, ",")(0)
        Dim panelCount As Integer = CInt(InputBox("�p�l����"))

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
    ' �p�l���g���쐬
    Private Sub MacroCommand2()
        On Error Resume Next
        Dim doc As Document = ActiveDocument

        Dim drawing As Drawing = doc.CurrentDrawing
        Dim points(1) As Point2d
        Dim creator As PanelCreator = New PanelCreator(drawing, Geometry, 0, 0, 0, doc.SelectionManager)
        creator.writePanelLine()
    End Sub

    Private Class PanelCreator
        Dim currentX As Double
        Dim currentY As Double
        Dim drawing As Drawing
        Dim Geometry As Geometry
        Dim selectinoManager As SelectionManager
        Public panelCounter As Integer
        Public angle As Integer = 0
        ' panel�T�C�Y
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
        Public Sub run(ByVal panelGroupCount As Integer)
            Dim groupCount As Integer = panelCounter \ panelGroupCount
            Dim backupX As Double = currentX
            Dim textShape As TextShape
            Dim panel As PolylineShape

            If groupCount > 0 Then
                For i = 1 To groupCount
                    panel = Drawing.Shapes.AddPolyline(createAPanel(currentX, currentY, panelGroupCount))
                    panel.ColorNumber = 6
                    'panel.Rotate(Geometry.CreatePoint(currentX, currentY), angle)
                    ' �p�l���I��
                    Me.selectinoManager.SelectShape(panel, SelectionCombineMode.Add)

                    If panelGroupCount > 4 Then
                        Dim tmpX As Double = currentX + ((width * panelGroupCount) / 2)
                        'drawing.Shapes.AddText(CStr(panelGroupCount) + "��", Geometry.CreatePoint(backupX + (width * panelGroupCount) / 2, currentY), 0)
                        Dim textPoint As Point2d = Geometry.CreatePoint(currentX, currentY - 1000)
                        textShape = Drawing.Shapes.AddText(CStr(panelGroupCount) + "��", textPoint, 0)
                        textShape.FontHeight = charaH
                        'textShape.Rotate(textPoint, angle)
                        ' �����I��
                        Me.selectinoManager.SelectShape(textShape, SelectionCombineMode.Add)
                    End If
                    currentX = currentX + (width * panelGroupCount)
                Next
                panelCounter = panelCounter - (panelGroupCount * groupCount)

            End If

        End Sub
        Public Function createAPanel(ByVal x As Double, ByVal y As Double, ByVal panelCount As Integer) As Point2d()

            Dim panelPoints(4) As Point2d

            panelPoints(0) = Geometry.CreatePoint(x, y)
            panelPoints(1) = Geometry.CreatePoint(x + (width * panelCount), y)
            panelPoints(2) = Geometry.CreatePoint(x + (width * panelCount), y - height)
            panelPoints(3) = Geometry.CreatePoint(x, y - height)
            panelPoints(4) = Geometry.CreatePoint(x, y)
            createAPanel = panelPoints

        End Function
        Public Sub writePanelLine()
            Dim shape As SelectedShape = Me.selectinoManager.SelectedShapes.Item(0)
            Dim firstPoint As Point2d = shape.Shape.GetBoundingPoints().GetValue(0)
            Dim panelLine As PolylineShape

            Dim linePoints(1) As Point2d
            Dim i As Integer = 0

            Dim ys(4) As Double
            For i = 0 To UBound(shape.Shape.GetBoundingPoints())
                ys(i) = shape.Shape.GetBoundingPoints().GetValue(i).Y
            Next

            Array.Sort(ys)

            Dim baseY As Double = ys.GetValue(ys.Length - 1)

            i = 0
            For Each s In shape.Shape.GetBoundingPoints()
                If s.Y = baseY Then
                    linePoints(i) = Geometry.CreatePoint(s.X, s.Y + 500)
                    i = i + 1
                    If i > 1 Then
                        Exit For
                    End If
                End If
            Next

            panelLine = drawing.Shapes.AddPolyline(linePoints)
            panelLine.LinetypeNumber = 2
            panelLine.ColorNumber = 4
            panelLine.LinewidthNumber = 3

        End Sub

    End Class

End Class

