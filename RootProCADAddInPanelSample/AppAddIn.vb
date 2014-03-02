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
    'Const width As Double = 1660.0
    'Const height As Double = 4150.0
    Private Sub AppAddIn_Startup(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Startup
        CommandManager.AddMacroCommand("���z���p�l���쐬", AddressOf Me.MacroCommand)
        CommandManager.AddMacroCommand("�p�l���g���쐬", AddressOf Me.MacroCommand2)
        CommandManager.AddMacroCommand("�p�l���ݒu�V�~�����[�V����", AddressOf Me.MacroCommand3)
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
    Private Sub MacroCommand3()
        On Error Resume Next
        Dim doc As Document = ActiveDocument

        Dim drawing As Drawing = doc.CurrentDrawing
        Dim points(1) As Point2d
        Dim simulator As PanelSimulater = New PanelSimulater(drawing, Geometry, 0, 0, 0, doc.SelectionManager)
        MsgBox(doc.SelectionManager.SelectedShapes.Count)

        simulator.simulate()


    End Sub

    Public Class PanelCreator
        Protected currentX As Double
        Protected currentY As Double
        Protected drawing As Drawing
        Protected Geometry As Geometry
        Protected selectinoManager As SelectionManager
        ' �u�������񐔁iSimulator�̏ꍇ�́A���̗񐔁j
        Public panelCounter As Integer
        Public angle As Integer = 0
        ' panel�T�C�Y
        Public Const width As Double = 1660.0
        Public Const height As Double = 4150.0
        Public Const charaH As Double = 2500
        Public Sub New(ByVal drawing As Drawing, ByVal geometry As Geometry, ByVal panelCounter As Integer, ByVal currentX As Double, ByVal currentY As Double, ByVal selectionManager As SelectionManager)
            Me.currentX = currentX
            Me.currentY = currentY
            Me.panelCounter = panelCounter
            Me.drawing = drawing
            Me.Geometry = geometry
            Me.selectinoManager = selectionManager
        End Sub
        ' �u�������������p�l����u���i�ÓI���\�b�h�Łj
        Public Shared Sub putPanel(ByVal drawing As Drawing, ByVal geometry As Geometry, ByVal selectionManager As SelectionManager, ByVal theNumberWantToPut As Integer, ByVal startPoint As Point2d)
            Dim currentX As Double = startPoint.X
            Dim currentY As Double = startPoint.Y
            Dim panelCounter As Integer = theNumberWantToPut
            Dim creator As PanelCreator = New PanelCreator(drawing, geometry, panelCounter, currentX, currentY, selectionManager)
            creator.run(20)
            creator.run(10)
            creator.run(5)
            creator.run(1)
        End Sub
        ' �u�������������p�l����u���i���I���\�b�h�Łj
        Public Sub run(ByVal panelGroupCount As Integer)
            Dim groupCount As Integer = panelCounter \ panelGroupCount
            Dim backupX As Double = currentX
            Dim textShape As TextShape
            Dim panel As PolylineShape

            If groupCount > 0 Then
                For i = 1 To groupCount
                    panel = drawing.Shapes.AddPolyline(createAPanel(currentX, currentY, panelGroupCount))
                    panel.ColorNumber = 6
                    'panel.Rotate(Geometry.CreatePoint(currentX, currentY), angle)
                    ' �p�l���I��
                    Me.selectinoManager.SelectShape(panel, SelectionCombineMode.Add)

                    If panelGroupCount > 4 Then
                        Dim tmpX As Double = currentX + ((width * panelGroupCount) / 2)
                        'drawing.Shapes.AddText(CStr(panelGroupCount) + "��", Geometry.CreatePoint(backupX + (width * panelGroupCount) / 2, currentY), 0)
                        Dim textPoint As Point2d = Geometry.CreatePoint(currentX, currentY - 1000)
                        textShape = drawing.Shapes.AddText(CStr(panelGroupCount) + "��", textPoint, 0)
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
            Dim panelLine As PolylineShape

            Dim linePoints(1) As Point2d
            Dim firstPoint As Point2d = getFirstPoint(shape.Shape)
            Dim endPoint As Point2d = getEndPoint(shape.Shape)
            linePoints(0) = Geometry.CreatePoint(firstPoint.X, firstPoint.Y + 500)
            linePoints(1) = Geometry.CreatePoint(endPoint.X, endPoint.Y + 500)

            panelLine = drawing.Shapes.AddPolyline(linePoints)
            panelLine.LinetypeNumber = 2
            panelLine.ColorNumber = 4
            panelLine.LinewidthNumber = 3

        End Sub
        ' ----- utils -----
        ' �}�`�̎n�_���擾����
        Public Shared Function getFirstPoint(ByVal shape As Shape) As Point2d
            getFirstPoint = shape.GetBoundingPoints(0)
            For Each p In shape.GetBoundingPoints
                If getFirstPoint.X >= p.X And getFirstPoint.Y <= p.Y Then
                    getFirstPoint = p
                End If
            Next

        End Function
        ' �}�`�̏I�_���擾����
        Protected Shared Function getEndPoint(ByVal shape As Shape) As Point2d
            getEndPoint = shape.GetBoundingPoints(0)
            For Each p In shape.GetBoundingPoints
                If getEndPoint.X <= p.X And getEndPoint.Y <= p.Y Then
                    getEndPoint = p
                End If
            Next
        End Function
        ' �}�`�̒������擾����
        Public Shared Function getLength(ByVal shape As Shape) As Double
            Dim s As Point2d = getFirstPoint(shape)
            Dim e As Point2d = getEndPoint(shape)
            Return e.Y - s.Y
        End Function
    End Class
    ' �p�l���z�u�V�~�����[�^�[
    Public Class PanelSimulater
        Inherits PanelCreator
        Public Sub New(ByVal drawing As Drawing, ByVal geometry As Geometry, ByVal panelCounter As Integer, ByVal currentX As Double, ByVal currentY As Double, ByVal selectionManager As SelectionManager)
            MyBase.New(drawing, geometry, panelCounter, currentX, currentY, selectionManager)
        End Sub

        Public Sub selectionSample()

            'For i = 0 To selectinoManager.SelectedShapes.Count
            '    Dim s As Shape = selectinoManager.SelectedShapes.Item(i).Shape
            '    MsgBox(getFirstPoint(s).Y)
            'Next

            Dim col As SelectedLineCollection = New SelectedLineCollection(selectinoManager.SelectedShapes)
            MsgBox(PanelCreator.getFirstPoint(col.getNext()).Y)
            MsgBox(PanelCreator.getFirstPoint(col.getNext()).Y)
            MsgBox(PanelCreator.getFirstPoint(col.getNext()).Y)

        End Sub
        ' �p�l���z�u�V�~�����[�V����
        ' �u���邾���u��
        Public Sub simulate()

            Dim lines As SelectedLineCollection = New SelectedLineCollection(selectinoManager.SelectedShapes)
            Dim basePanelCount As Integer = Me.panelCounter

            Dim param As PutPanelVo = New PutPanelVo(PanelCreator.getLength(lines.getCurrent()), panelCounter, PanelCreator.getFirstPoint(lines.getCurrent()), Nothing)

            Do While True

                Dim ret As PutPanelVo = putPanelFully(param)
                If ret.isFetch = True Then
                    ' �u���Ȃ�
                    If lines.hasNext = True Then
                        Dim nextLine As Shape = lines.getNext()
                        param = New PutPanelVo(PanelCreator.getLength(nextLine), ret.theNumberWantToPut, PanelCreator.getFirstPoint(nextLine), Nothing)
                    Else
                        ' �����I��!
                        Exit Do
                    End If
                Else
                    '�u����
                    param = New PutPanelVo(ret.length, panelCounter, ret.startPoint, Nothing)

                End If
            Loop

            MsgBox("����!!")

        End Sub
        ' �p�l�����߂������ς��z�u
        ' Return: ����line��Fetch���邩�ǂ���
        Protected Function putPanelFully(ByVal param As PutPanelVo) As PutPanelVo

            Dim ret As PutPanelVo

            ' �����ɑ΂��ĉ������邩�H
            ' TODO �u�������2�������������̍l��
            Dim theNumberCanBePut As Integer = param.length / PanelCreator.width

            ' TODO �u����������2�������������̑Ή�

            ' �u�������������邩
            If param.theNumberWantToPut > theNumberCanBePut Then
                ' �u���Ȃ� -> �u���镪�����u��
                PanelCreator.putPanel(drawing, Geometry, selectinoManager, theNumberCanBePut, param.startPoint)
                ret = New PutPanelVo(0, param.theNumberWantToPut - theNumberCanBePut, Nothing, True)
            Else
                ' �u���� -> �u���������u��
                PanelCreator.putPanel(drawing, Geometry, selectinoManager, param.theNumberWantToPut, param.startPoint)
                ' Todo 2m������
                ' 
                Dim startPoint As Point2d
                startPoint = New Point2d(param.startPoint.X + (width * param.theNumberWantToPut) + width + 2000, param.startPoint.Y)
                ret = New PutPanelVo(0, panelCounter, startPoint, False)
            End If

            putPanelFully = ret

        End Function
    End Class
    ' �p�l���ݒu���\�b�h�̃p�����^�N���X
    Public Class PutPanelVo
        Public length As Double
        Public theNumberWantToPut As Integer
        Public startPoint As Point2d
        Public isFetch As Boolean
        Public Sub New(ByVal length As Double, ByVal theNumberWantToPut As Integer, ByVal startPoint As Point2d, ByVal isFetch As Boolean)
            Me.length = length
            Me.theNumberWantToPut = theNumberWantToPut
            Me.startPoint = startPoint
            Me.isFetch = isFetch
        End Sub

    End Class

    ' �I���ς݂̊�����Ǘ�����N���X
    Public Class SelectedLineCollection
        Protected lineArray() As Shape
        Protected currentIndex As Integer
        Public Sub New(ByVal lines As SelectedShapeCollection)
            For i = 0 To lines.Count - 1
                ReDim Preserve lineArray(i)
                lineArray(i) = lines.Item(i).Shape
            Next
            currentIndex = 0
        End Sub
        Public Function getCurrent() As Shape
            Return lineArray(currentIndex)
        End Function
        Public Function getNext() As Shape
            'Dim ret As Shape
            'Dim base As Shape = lineArray.GetValue(currentIndex)
            'Dim cindex As Integer

            'ret = lineArray.GetValue(currentIndex)
            'For i = 0 To UBound(lineArray)
            '    Dim s As Shape = lineArray.GetValue(i)
            '    If PanelCreator.getFirstPoint(base).Y > PanelCreator.getFirstPoint(s).Y And PanelCreator.getFirstPoint(ret).Y < PanelCreator.getFirstPoint(s).Y Then
            '        ret = s
            '        cindex = i
            '    End If
            'Next

            currentIndex = currentIndex + 1
            getNext = lineArray.GetValue(currentIndex)
        End Function
        Public Function hasNext() As Boolean
            Return currentIndex = lineArray.Length - 1
        End Function


    End Class


End Class