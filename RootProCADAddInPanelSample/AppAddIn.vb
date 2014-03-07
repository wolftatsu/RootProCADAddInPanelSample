Imports System
Imports System.Collections.Generic
Imports System.Windows.Forms
Imports RootPro.RootProCAD
Imports RootPro.RootProCAD.Command
Imports RootPro.RootProCAD.Geometry
Imports RootPro.RootProCAD.UI
<System.AddIn.AddIn("AppAddIn", Version:="1.0", Publisher:="", Description:="")> _
Partial Class AppAddIn
    Private Sub AppAddIn_Startup(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Startup
        CommandManager.AddMacroCommand("���z���p�l���쐬", AddressOf Me.MacroCommand)
        CommandManager.AddMacroCommand("���g���쐬", AddressOf Me.MacroCommand2)
        CommandManager.AddMacroCommand("���g���i�c�j�쐬", AddressOf Me.MacroCommand7)
        CommandManager.AddMacroCommand("30�Z���`���H�쐬", AddressOf Me.MacroCommand6)
        CommandManager.AddMacroCommand("�p�l���ݒu�V�~�����[�V�����i�ォ�珇�ɔz�u�j", AddressOf Me.MacroCommand3)
        CommandManager.AddMacroCommand("�p�l���ݒu�V�~�����[�V�����i�����珇�ɔz�u�j", AddressOf Me.MacroCommand4)
        CommandManager.AddMacroCommand("���ԍ��t��", AddressOf Me.MacroCommand5)
        ' CommandManager.AddMacroCommand("TestSelectionManager", AddressOf Me.TestSelectionManager)
    End Sub
    Private Sub AppAddIn_Shutdown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shutdown
        CommandManager.RemoveMacroCommand(AddressOf Me.MacroCommand)
        CommandManager.RemoveMacroCommand(AddressOf Me.MacroCommand2)
        CommandManager.RemoveMacroCommand(AddressOf Me.MacroCommand3)
        CommandManager.RemoveMacroCommand(AddressOf Me.MacroCommand4)
        CommandManager.RemoveMacroCommand(AddressOf Me.MacroCommand5)
        CommandManager.RemoveMacroCommand(AddressOf Me.MacroCommand6)
        CommandManager.RemoveMacroCommand(AddressOf Me.MacroCommand7)
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
        Dim creator As PanelCreator = New PanelCreator(drawing, Geometry, panelCounter, currentX, currentY, doc.SelectionManager, doc.LayerTable.RootLayer.ChildLayers)

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
        Dim creator As PanelCreator = New PanelCreator(drawing, Geometry, 0, 0, 0, doc.SelectionManager, doc.LayerTable.RootLayer.ChildLayers)
        doc.UndoManager.BeginUndoUnit()
        creator.writePanelLine(500)
        doc.UndoManager.EndUndoUnit()
    End Sub
    '   �p�l���ݒu�V�~�����[�V�������s
    Private Sub MacroCommand3()
        On Error Resume Next
        Dim doc As Document = ActiveDocument
        Dim panelCount As Integer = CInt(InputBox("�p�l���񐔂���͂��Ă�������"))
        Dim drawing As Drawing = doc.CurrentDrawing
        Dim simulator As PanelSimulater = New PanelSimulater(drawing, Geometry, panelCount, 0, 0, doc.SelectionManager, doc.LayerTable.RootLayer.ChildLayers)
        '�g���@�\on
        simulator.isTheSandboxAvailable = True
        simulator.orderByAsc = True
        ' MsgBox(doc.SelectionManager.SelectedShapes.Count)

        doc.UndoManager.BeginUndoUnit()
        simulator.simulate()
        doc.UndoManager.EndUndoUnit()

    End Sub
    '   �p�l���ݒu�V�~�����[�V�������s�i������z�u�j
    Private Sub MacroCommand4()
        On Error Resume Next
        Dim doc As Document = ActiveDocument
        Dim panelCount As Integer = CInt(InputBox("�p�l���񐔂���͂��Ă�������"))
        Dim drawing As Drawing = doc.CurrentDrawing
        Dim simulator As PanelSimulater = New PanelSimulater(drawing, Geometry, panelCount, 0, 0, doc.SelectionManager, doc.LayerTable.RootLayer.ChildLayers)
        '�g���@�\on
        simulator.isTheSandboxAvailable = True
        simulator.orderByAsc = False
        ' MsgBox(doc.SelectionManager.SelectedShapes.Count)

        doc.UndoManager.BeginUndoUnit()
        simulator.simulate()
        doc.UndoManager.EndUndoUnit()

    End Sub
    ' ���ԍ��t��
    Private Sub MacroCommand5()
        On Error Resume Next
        Dim doc As Document = ActiveDocument
        Dim drawing As Drawing = doc.CurrentDrawing

        Dim names As SelectedLineCollection = New SelectedLineCollection(doc.SelectionManager.SelectedShapes, False)
        Dim putPanelCounter As Integer = CInt(InputBox("�J�n�ԍ�")) - 1
        If putPanelCounter < 1 Then
            putPanelCounter = 0
        End If
        Dim simulator As PanelSimulater = New PanelSimulater(drawing, Geometry, 52, 0, 0, doc.SelectionManager, doc.LayerTable.RootLayer.ChildLayers)
        '�g���@�\on
        simulator.isTheSandboxAvailable = True
        simulator.orderByAsc = False


        Dim shape As Shape = names.getCurrent()

        doc.UndoManager.BeginUndoUnit()

        Do While names.hasNext

            putPanelCounter = putPanelCounter + 1

            Dim linePoints(1) As Point2d
            linePoints(0) = PanelCreator.getFirstPoint(shape)
            linePoints(1) = Geometry.CreatePoint(linePoints(0).X + 2000, linePoints(0).Y + 2000)
            Dim areaNumber As LeadShape = drawing.Shapes.AddLead("���" + CStr(putPanelCounter), linePoints)

            areaNumber.FontHeight = 2500
            areaNumber.Layer = simulator.getPanelLayer()

            shape.Delete()

            shape = names.getNext()

            ' areaNumber.Layer = Me.getPanelLayer()

        Loop

        doc.UndoManager.EndUndoUnit()

    End Sub
    ' 30cm���H�쐬
    Private Sub MacroCommand6()
        On Error Resume Next
        Dim doc As Document = ActiveDocument

        Dim drawing As Drawing = doc.CurrentDrawing
        Dim points(1) As Point2d
        Dim creator As PanelCreator = New PanelCreator(drawing, Geometry, 0, 0, 0, doc.SelectionManager, doc.LayerTable.RootLayer.ChildLayers)
        doc.UndoManager.BeginUndoUnit()
        creator.writePanelLine(300)
        doc.UndoManager.EndUndoUnit()

    End Sub
    ' �����i�c�j�쐬
    Private Sub MacroCommand7()
        On Error Resume Next
        Dim doc As Document = ActiveDocument

        Dim drawing As Drawing = doc.CurrentDrawing
        Dim points(1) As Point2d
        Dim creator As PanelCreator = New PanelCreator(drawing, Geometry, 0, 0, 0, doc.SelectionManager, doc.LayerTable.RootLayer.ChildLayers)
        doc.UndoManager.BeginUndoUnit()
        creator.writePanelLineVartical()
        doc.UndoManager.EndUndoUnit()
    End Sub
    Private Sub TestSelectionManager()
        On Error Resume Next
        Dim doc As Document = ActiveDocument
        Dim panelCount As Integer = CInt(InputBox("�p�l���񐔂���͂��Ă�������"))
        Dim drawing As Drawing = doc.CurrentDrawing
        Dim simulator As PanelSimulater = New PanelSimulater(drawing, Geometry, panelCount, 0, 0, doc.SelectionManager, Nothing)
        ' MsgBox(doc.SelectionManager.SelectedShapes.Count)
        Dim rootLayer As Layer = doc.LayerTable.RootLayer

        Dim l1 As Layer = rootLayer.ChildLayers(0)
        Dim l2 As Layer = rootLayer.ChildLayers(1)
        MsgBox(l2.Name)
        'For Each l As Layer In rootLayer.ChildLayers
        '    Dim str2 = l.Name()

        'Next

        simulator.selectionSample()


    End Sub
    ' �p�l���쐬�N���X
    Public Class PanelCreator
        Protected currentX As Double
        Protected currentY As Double
        Protected drawing As Drawing
        Protected Geometry As Geometry
        Protected selectinoManager As SelectionManager
        Protected layers As LayerCollection
        ' �u�������񐔁iSimulator�̏ꍇ�́A���̗񐔁j
        Public panelCounter As Integer
        Public angle As Integer = 0
        ' panel�T�C�Y
        Public Const width As Double = 1660.0
        Public Const height As Double = 4150.0
        Public Const charaH As Double = 2500
        Public Sub New(ByVal drawing As Drawing, ByVal geometry As Geometry, ByVal panelCounter As Integer, ByVal currentX As Double, ByVal currentY As Double, ByVal selectionManager As SelectionManager, ByRef layers As LayerCollection)
            Me.currentX = currentX
            Me.currentY = currentY
            Me.panelCounter = panelCounter
            Me.drawing = drawing
            Me.Geometry = geometry
            Me.selectinoManager = selectionManager
            Me.layers = layers
        End Sub
        Public Sub New(ByVal drawing As Drawing, ByVal geometry As Geometry, ByVal panelCounter As Integer, ByVal currentX As Double, ByVal currentY As Double, ByVal selectionManager As SelectionManager)
            MyClass.New(drawing, geometry, panelCounter, currentX, currentY, selectionManager, Nothing)
        End Sub
        ' �p�l���z�u
        ' �u�������������p�l����u���i�ÓI���\�b�h�Łj
        Public Shared Sub putPanel(ByVal drawing As Drawing, ByVal geometry As Geometry, ByVal selectionManager As SelectionManager, ByVal theNumberWantToPut As Integer, ByVal startPoint As Point2d, ByRef layers As LayerCollection)
            Dim currentX As Double = startPoint.X
            Dim currentY As Double = startPoint.Y
            Dim panelCounter As Integer = theNumberWantToPut
            Dim creator As PanelCreator = New PanelCreator(drawing, geometry, panelCounter, currentX, currentY, selectionManager, layers)
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
                    panel.LinewidthNumber = 3
                    panel.LinetypeNumber = 1
                    panel.Layer = Me.getPanelLayer()

                    ' �p�l���I��
                    Me.selectinoManager.SelectShape(panel, SelectionCombineMode.Add)

                    If panelGroupCount > 4 Then
                        Dim tmpX As Double = currentX + ((width * panelGroupCount) / 2)
                        Dim textPoint As Point2d = Geometry.CreatePoint(currentX, currentY - 1000)
                        textShape = drawing.Shapes.AddText(CStr(panelGroupCount) + "��", textPoint, 0)
                        textShape.FontHeight = charaH
                        textShape.Layer = Me.getPanelLayer()
                        textShape.ColorNumber = 1


                        ' �����I��
                        Me.selectinoManager.SelectShape(textShape, SelectionCombineMode.Add)
                    End If
                    currentX = currentX + (width * panelGroupCount)
                Next
                panelCounter = panelCounter - (panelGroupCount * groupCount)

            End If

        End Sub
        ' �p�l���z�u���W��Ԃ��i�A�������̍��W�z��j
        Public Function createAPanel(ByVal x As Double, ByVal y As Double, ByVal panelCount As Integer) As Point2d()

            Dim panelPoints(4) As Point2d

            panelPoints(0) = Geometry.CreatePoint(x, y)
            panelPoints(1) = Geometry.CreatePoint(x + (width * panelCount), y)
            panelPoints(2) = Geometry.CreatePoint(x + (width * panelCount), y - height)
            panelPoints(3) = Geometry.CreatePoint(x, y - height)
            panelPoints(4) = Geometry.CreatePoint(x, y)
            createAPanel = panelPoints

        End Function
        ' �p�l���z�u�K�C�h���̏㕔xcm�̏��ɋ��E�̐�������
        Public Sub writePanelLine(ByVal top As Integer)
            Dim shape As SelectedShape = Me.selectinoManager.SelectedShapes.Item(0)
            Dim linePoints(1) As Point2d
            Dim firstPoint As Point2d = Geometry.CreatePoint(getFirstPoint(shape.Shape).X, getFirstPoint(shape.Shape).Y + top)
            Dim endPoint As Point2d = Geometry.CreatePoint(getEndPoint(shape.Shape).X + 1000, getEndPoint(shape.Shape).Y + top)
            writePanelLine(firstPoint, endPoint)
        End Sub
        Public Sub writePanelLineVartical()
            Dim shape As SelectedShape = Me.selectinoManager.SelectedShapes.Item(0)
            Dim linePoints(1) As Point2d
            Dim firstPoint As Point2d = Geometry.CreatePoint(getEndPoint(shape.Shape).X + 1000, getEndPoint(shape.Shape).Y + 500)
            Dim endPoint As Point2d = Geometry.CreatePoint(getEndPoint(shape.Shape).X + 1000, getEndPoint(shape.Shape).Y - 7000)
            writePanelLine(firstPoint, endPoint)

        End Sub

        Public Sub writePanelLine(ByVal firstPoint As Point2d, ByVal endPoint As Point2d)
            Dim panelLine As PolylineShape
            Dim linePoints(1) As Point2d
            linePoints(0) = Geometry.CreatePoint(firstPoint.X, firstPoint.Y)
            linePoints(1) = Geometry.CreatePoint(endPoint.X, endPoint.Y)

            panelLine = drawing.Shapes.AddPolyline(linePoints)
            panelLine.LinetypeNumber = 1
            panelLine.ColorNumber = 4
            panelLine.LinewidthNumber = 2
            panelLine.Layer = Me.getPanelLayer()

        End Sub
        ' ----- utils -----
        ' �p�l�����C���擾
        Public Function getPanelLayer() As Layer
            For i As Integer = 0 To Me.layers.Count - 1
                If Me.layers(i).Name = "�p�l��" Then
                    Return layers(i)
                End If
            Next
            Return layers(0)
        End Function
        ' ----- static-utils -----
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
            getLength = e.X - s.X
        End Function

    End Class
    ' �p�l���z�u�V�~�����[�^�[
    Public Class PanelSimulater
        Inherits PanelCreator
        Public isTheSandboxAvailable As Boolean
        Protected putPanelCounter As Integer
        Public orderByAsc As Boolean
        Public Sub New(ByVal drawing As Drawing, ByVal geometry As Geometry, ByVal panelCounter As Integer, ByVal currentX As Double, ByVal currentY As Double, ByVal selectionManager As SelectionManager, ByRef layers As LayerCollection)
            MyBase.New(drawing, geometry, panelCounter, currentX, currentY, selectionManager, layers)
            putPanelCounter = 0
        End Sub
        ' SelectedObject�̋����m�F
        Public Sub selectionSample()

            'For i = 0 To selectinoManager.SelectedShapes.Count
            '    Dim s As Shape = selectinoManager.SelectedShapes.Item(i).Shape
            '    MsgBox(getFirstPoint(s).Y)
            'Next

            Dim col As SelectedLineCollection = New SelectedLineCollection(selectinoManager.SelectedShapes, True)
            MsgBox(PanelCreator.getFirstPoint(col.getCurrent()).Y)
            MsgBox(PanelCreator.getFirstPoint(col.getNext()).Y)
            MsgBox(PanelCreator.getFirstPoint(col.getNext()).Y)

        End Sub
        ' �p�l���z�u�V�~�����[�V����
        ' �I�����ꂽ�p�l���Ԋu���ɉ����āA���͂��ꂽ�񐔕��̃p�l����z�u����
        Public Sub simulate()

            Dim lines As SelectedLineCollection = New SelectedLineCollection(selectinoManager.SelectedShapes, Me.orderByAsc)
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
            Dim putCount As Integer

            ' �����ɑ΂��ĉ������邩�H
            Dim theNumberCanBePut As Integer = param.length / PanelCreator.width
            ' �u�������2�������������̍l��
            If theNumberCanBePut < 2 Then
                Return New PutPanelVo(0, param.theNumberWantToPut, Nothing, True)
            End If

            ' �u�������������邩
            If param.theNumberWantToPut > theNumberCanBePut Then
                ' �u���Ȃ� -> �u���镪�����u��
                putCount = theNumberCanBePut
                PanelCreator.putPanel(drawing, Geometry, selectinoManager, theNumberCanBePut, param.startPoint, layers)
                ret = New PutPanelVo(0, param.theNumberWantToPut - theNumberCanBePut, Nothing, True)
            Else
                ' �u���� -> �u���������u��
                putCount = param.theNumberWantToPut
                PanelCreator.putPanel(drawing, Geometry, selectinoManager, param.theNumberWantToPut, param.startPoint, layers)
                ' Todo 2m������
                Dim startPoint As Point2d
                startPoint = New Point2d(param.startPoint.X + (width * param.theNumberWantToPut) + 2000, param.startPoint.Y)
                ' �u����������1�̏ꍇ�́A��C���ׁ̈A1�p�l�����J���Ă���
                If param.theNumberWantToPut < 2 Then
                    startPoint = New Point2d(startPoint.X + width, startPoint.Y)
                End If
                Dim length = param.length - (startPoint.X - param.startPoint.X)
                ret = New PutPanelVo(length, panelCounter, startPoint, False)
            End If

            ' ���g���`��
            If Me.isTheSandboxAvailable = True Then
                If param.theNumberWantToPut = panelCounter Then
                    ' ���J�n��
                    Dim lineFirstPoint As Point2d = Geometry.CreatePoint(param.startPoint.X - 1000, param.startPoint.Y + IIf(Me.orderByAsc, 500, -6500))
                    Dim lineEndPoint As Point2d = Geometry.CreatePoint(param.startPoint.X + (width * putCount), param.startPoint.Y + IIf(Me.orderByAsc, 500, -6500))
                    writePanelLine(lineFirstPoint, lineEndPoint)
                End If
                If param.theNumberWantToPut <= theNumberCanBePut Then
                    ' ���I����
                    ' �p�l������
                    ' FIXME ���Ԋu��7m�Œ�
                    Dim lineFirstPoint As Point2d = Geometry.CreatePoint(param.startPoint.X, param.startPoint.Y + IIf(Me.orderByAsc, -6500, 500))
                    Dim lineEndPoint As Point2d = Geometry.CreatePoint(param.startPoint.X + (width * putCount) + 1000, param.startPoint.Y + IIf(Me.orderByAsc, -6500, 500))
                    writePanelLine(lineFirstPoint, lineEndPoint)
                    ' ���Ԃ̐�
                    lineFirstPoint = Geometry.CreatePoint(lineEndPoint.X, lineEndPoint.Y + IIf(Me.orderByAsc, 7000, 0))
                    lineEndPoint = Geometry.CreatePoint(lineEndPoint.X, lineEndPoint.Y + IIf(Me.orderByAsc, 0, -7000))
                    writePanelLine(lineFirstPoint, lineEndPoint)
                End If
            End If

            ' ���ԍ��쐬
            If Me.isTheSandboxAvailable = True And param.theNumberWantToPut = panelCounter Then
                putPanelCounter = putPanelCounter + 1
                Dim textPointAry(1) As Point2d
                textPointAry(0) = param.startPoint
                textPointAry(1) = Geometry.CreatePoint(param.startPoint.X + 2000, param.startPoint.Y + 2000)
                Dim areaNumber As LeadShape = Me.drawing.Shapes.AddLead("���" + CStr(putPanelCounter), textPointAry)
                areaNumber.FontHeight = 2500
                areaNumber.Layer = Me.getPanelLayer()
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
        'Protected lineArray() As Shape
        Protected lineArray = New List(Of SortableShape)
        Protected currentIndex As Integer
        Public Sub New(ByVal lines As SelectedShapeCollection, ByVal orderByAsc As Boolean)
            For i = 0 To lines.Count - 1
                'ReDim Preserve lineArray(i)
                'lineArray(i) = lines.Item(i).Shape
                lineArray.add(New SortableShape(PanelCreator.getFirstPoint(lines.Item(i).Shape).Y, lines.Item(i).Shape, orderByAsc))
            Next
            lineArray.Sort()
            currentIndex = 0
        End Sub
        Public Function getCurrent() As Shape
            Return lineArray(currentIndex).shape
        End Function
        Public Function getNext() As Shape
            Dim al As New System.Collections.ArrayList()
            If hasNext() = True Then
                currentIndex = currentIndex + 1
                getNext = lineArray(currentIndex).shape
            Else
                getNext = getCurrent()
            End If
        End Function
        Public Function hasNext() As Boolean
            Return Not currentIndex = lineArray.count - 1
        End Function
    End Class
    ' shape��y���Ń\�[�g�\�Ƀ��b�v�����N���X
    Public Class SortableShape
        Implements IComparable(Of SortableShape)
        Public y As Double
        Public shape As Shape
        Public orderByAsc As Boolean
        Public Sub New(ByVal y As Double, ByVal shape As Shape, ByVal orderByAsc As Boolean)
            Me.y = y
            Me.shape = shape
            Me.orderByAsc = orderByAsc
        End Sub
        Public Function CompareTo(ByVal obj As SortableShape) As Integer Implements System.IComparable(Of SortableShape).CompareTo

            If Me.y.CompareTo(obj.y) < 0 Then
                Return IIf(Me.orderByAsc = True, 1, -1)
            ElseIf Me.y.CompareTo(obj.y) > 0 Then
                Return IIf(Me.orderByAsc = True, -1, 1)
            Else
                Return 0
            End If


        End Function
    End Class

End Class
