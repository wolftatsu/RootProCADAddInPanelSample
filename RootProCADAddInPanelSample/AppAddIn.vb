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
        CommandManager.AddMacroCommand("太陽光パネル作成", AddressOf Me.MacroCommand)
        CommandManager.AddMacroCommand("パネル枠線作成", AddressOf Me.MacroCommand2)
        CommandManager.AddMacroCommand("パネル設置シミュレーション", AddressOf Me.MacroCommand3)
        CommandManager.AddMacroCommand("TestSelectionManager", AddressOf Me.TestSelectionManager)
    End Sub

    Private Sub AppAddIn_Shutdown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shutdown
        CommandManager.RemoveMacroCommand(AddressOf Me.MacroCommand)
        CommandManager.RemoveMacroCommand(AddressOf Me.MacroCommand2)
        CommandManager.RemoveMacroCommand(AddressOf Me.MacroCommand3)
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
    '   パネル設置シミュレーション実行
    Private Sub MacroCommand3()
        On Error Resume Next
        Dim doc As Document = ActiveDocument
        Dim panelCount As Integer = CInt(InputBox("パネル列数を入力してください"))
        Dim drawing As Drawing = doc.CurrentDrawing
        Dim simulator As PanelSimulater = New PanelSimulater(drawing, Geometry, panelCount, 0, 0, doc.SelectionManager)
        ' MsgBox(doc.SelectionManager.SelectedShapes.Count)

        doc.UndoManager.BeginUndoUnit()
        simulator.simulate()
        doc.UndoManager.EndUndoUnit()

    End Sub
    Private Sub TestSelectionManager()
        On Error Resume Next
        Dim doc As Document = ActiveDocument
        Dim panelCount As Integer = CInt(InputBox("パネル列数を入力してください"))
        Dim drawing As Drawing = doc.CurrentDrawing
        Dim simulator As PanelSimulater = New PanelSimulater(drawing, Geometry, panelCount, 0, 0, doc.SelectionManager)
        ' MsgBox(doc.SelectionManager.SelectedShapes.Count)

        simulator.selectionSample()
    End Sub
    ' パネル作成クラス
    Public Class PanelCreator
        Protected currentX As Double
        Protected currentY As Double
        Protected drawing As Drawing
        Protected Geometry As Geometry
        Protected selectinoManager As SelectionManager
        ' 置きたい列数（Simulatorの場合は、基底の列数）
        Public panelCounter As Integer
        Public angle As Integer = 0
        ' panelサイズ
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
        ' パネル配置
        ' 置きたい個数だけパネルを置く（静的メソッド版）
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
        ' 置きたい個数だけパネルを置く（動的メソッド版）
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
                    ' パネル選択
                    Me.selectinoManager.SelectShape(panel, SelectionCombineMode.Add)

                    If panelGroupCount > 4 Then
                        Dim tmpX As Double = currentX + ((width * panelGroupCount) / 2)
                        Dim textPoint As Point2d = Geometry.CreatePoint(currentX, currentY - 1000)
                        textShape = drawing.Shapes.AddText(CStr(panelGroupCount) + "列", textPoint, 0)
                        textShape.FontHeight = charaH
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
        ' 図形の始点を取得する
        Public Shared Function getFirstPoint(ByVal shape As Shape) As Point2d
            getFirstPoint = shape.GetBoundingPoints(0)
            For Each p In shape.GetBoundingPoints
                If getFirstPoint.X >= p.X And getFirstPoint.Y <= p.Y Then
                    getFirstPoint = p
                End If
            Next
        End Function
        ' 図形の終点を取得する
        Protected Shared Function getEndPoint(ByVal shape As Shape) As Point2d
            getEndPoint = shape.GetBoundingPoints(0)
            For Each p In shape.GetBoundingPoints
                If getEndPoint.X <= p.X And getEndPoint.Y <= p.Y Then
                    getEndPoint = p
                End If
            Next
        End Function
        ' 図形の長さを取得する
        Public Shared Function getLength(ByVal shape As Shape) As Double
            Dim s As Point2d = getFirstPoint(shape)
            Dim e As Point2d = getEndPoint(shape)
            getLength = e.X - s.X
        End Function
    End Class
    ' パネル配置シミュレーター
    Public Class PanelSimulater
        Inherits PanelCreator
        Public Sub New(ByVal drawing As Drawing, ByVal geometry As Geometry, ByVal panelCounter As Integer, ByVal currentX As Double, ByVal currentY As Double, ByVal selectionManager As SelectionManager)
            MyBase.New(drawing, geometry, panelCounter, currentX, currentY, selectionManager)
        End Sub
        ' SelectedObjectの挙動確認
        Public Sub selectionSample()

            'For i = 0 To selectinoManager.SelectedShapes.Count
            '    Dim s As Shape = selectinoManager.SelectedShapes.Item(i).Shape
            '    MsgBox(getFirstPoint(s).Y)
            'Next

            Dim col As SelectedLineCollection = New SelectedLineCollection(selectinoManager.SelectedShapes)
            MsgBox(PanelCreator.getFirstPoint(col.getCurrent()).Y)
            MsgBox(PanelCreator.getFirstPoint(col.getNext()).Y)
            MsgBox(PanelCreator.getFirstPoint(col.getNext()).Y)

        End Sub
        ' パネル配置シミュレーション
        ' 選択されたパネル間隔線に沿って、入力された列数分のパネルを配置する
        Public Sub simulate()

            Dim lines As SelectedLineCollection = New SelectedLineCollection(selectinoManager.SelectedShapes)
            Dim basePanelCount As Integer = Me.panelCounter

            Dim param As PutPanelVo = New PutPanelVo(PanelCreator.getLength(lines.getCurrent()), panelCounter, PanelCreator.getFirstPoint(lines.getCurrent()), Nothing)

            Do While True

                Dim ret As PutPanelVo = putPanelFully(param)
                If ret.isFetch = True Then
                    ' 置けない
                    If lines.hasNext = True Then
                        Dim nextLine As Shape = lines.getNext()
                        param = New PutPanelVo(PanelCreator.getLength(nextLine), ret.theNumberWantToPut, PanelCreator.getFirstPoint(nextLine), Nothing)
                    Else
                        ' 処理終了!
                        Exit Do
                    End If
                Else
                    '置けた
                    param = New PutPanelVo(ret.length, panelCounter, ret.startPoint, Nothing)

                End If
            Loop

            MsgBox("完了!!")

        End Sub
        ' パネルをめいいっぱい配置
        ' Return: 次のlineにFetchするかどうか
        Protected Function putPanelFully(ByVal param As PutPanelVo) As PutPanelVo

            Dim ret As PutPanelVo

            ' 長さに対して何個おけるか？
            Dim theNumberCanBePut As Integer = param.length / PanelCreator.width
            ' 置ける個数が2個未満だった時の考慮
            If theNumberCanBePut < 2 Then
                Return New PutPanelVo(0, param.theNumberWantToPut, Nothing, True)
            End If

            ' 置きたい個数おけるか
            If param.theNumberWantToPut > theNumberCanBePut Then
                ' 置けない -> 置ける分だけ置く
                PanelCreator.putPanel(drawing, Geometry, selectinoManager, theNumberCanBePut, param.startPoint)
                ret = New PutPanelVo(0, param.theNumberWantToPut - theNumberCanBePut, Nothing, True)
            Else
                ' 置ける -> 置きたい数置く
                PanelCreator.putPanel(drawing, Geometry, selectinoManager, param.theNumberWantToPut, param.startPoint)
                ' Todo 2m幅おく
                Dim startPoint As Point2d
                startPoint = New Point2d(param.startPoint.X + (width * param.theNumberWantToPut) + 2000, param.startPoint.Y)
                ' 置きたい個数が1個の場合は、手修正の為、1パネル分開けておく
                If param.theNumberWantToPut < 2 Then
                    startPoint = New Point2d(startPoint.X + width, startPoint.Y)
                End If
                Dim length = param.length - (startPoint.X - param.startPoint.X)
                ret = New PutPanelVo(length, panelCounter, startPoint, False)
            End If

            putPanelFully = ret

        End Function
    End Class
    ' パネル設置メソッドのパラメタクラス
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
    ' 選択済みの基準線を管理するクラス
    Public Class SelectedLineCollection
        'Protected lineArray() As Shape
        Protected lineArray = New List(Of SortableShape)
        Protected currentIndex As Integer
        Public Sub New(ByVal lines As SelectedShapeCollection)
            For i = 0 To lines.Count - 1
                'ReDim Preserve lineArray(i)
                'lineArray(i) = lines.Item(i).Shape
                lineArray.add(New SortableShape(PanelCreator.getFirstPoint(lines.Item(i).Shape).Y, lines.Item(i).Shape))
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
    ' shapeをy軸でソート可能にラップしたクラス
    Public Class SortableShape
        Implements IComparable(Of SortableShape)
        Public y As Double
        Public shape As Shape
        Public Sub New(ByVal y As Double, ByVal shape As Shape)
            Me.y = y
            Me.shape = shape
        End Sub
        Public Function CompareTo(ByVal obj As SortableShape) As Integer Implements System.IComparable(Of SortableShape).CompareTo
            If Me.y.CompareTo(obj.y) < 0 Then
                Return 1
            ElseIf Me.y.CompareTo(obj.y) > 0 Then
                Return -2
            Else
                Return 0
            End If

        End Function
    End Class

End Class
