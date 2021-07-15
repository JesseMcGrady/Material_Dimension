Option Explicit On
Imports INFITF
Imports MECMOD
Imports PARTITF
Imports KnowledgewareTypeLib
Imports HybridShapeTypeLib
Imports SPATypeLib
Imports ProductStructureTypeLib
Imports System
Imports System.IO
Public Class CATIA_Property

    Dim mRootPart As MECMOD.Part
    Dim MyExcel As Object    'Application Excel
    Private _Documents As INFITF.Documents
    Property Documents As INFITF.Documents
        Get
            Return _Documents
        End Get
        Set(value As INFITF.Documents)
            If value Is Nothing Then
                Throw New ArgumentNullException("Documents", "Document can not be nothing")
            End If
            _Documents = value
        End Set
    End Property
    Private _ProductDocuments As INFITF.Document
    Property ProductDocument As INFITF.Document
        Get
            Return _ProductDocuments
        End Get
        Set(value As INFITF.Document)
            If value Is Nothing Then
                Throw New ArgumentNullException("Documents", "Document can not be nothing")
            End If
            _ProductDocuments = value
        End Set
    End Property
    Private _ShapeFactory As ShapeFactory
    Property ShapeFactory As ShapeFactory
        Get
            Return _ShapeFactory
        End Get
        Set(value As ShapeFactory)
            If value Is Nothing Then
                Throw New ArgumentNullException("ShapeFactory", "ShapeFactory can not be nothing")
            End If
            _ShapeFactory = value
        End Set
    End Property
    Private _HybridFactory As HybridShapeFactory
    Property HybridFactory As HybridShapeFactory
        Get
            Return _HybridFactory
        End Get
        Set(value As HybridShapeFactory)
            If value Is Nothing Then
                Throw New ArgumentNullException("HybridShapeFactory", "HybridShapeFactory can not be nothing")
            End If
            _HybridFactory = value
        End Set
    End Property
    Private _CATIA As INFITF.Application
    Property myCATIA As INFITF.Application
        Get
            Return _CATIA
        End Get
        Set(value As INFITF.Application)
            If value Is Nothing Then
                Throw New ArgumentNullException("CATIA", "CATIA can not be nothing")
            End If
            _CATIA = value
        End Set
    End Property
    Private _PartDocument As MECMOD.PartDocument
    Property PartDocument As MECMOD.PartDocument
        Get
            Return _PartDocument
        End Get
        Set(value As MECMOD.PartDocument)
            If value Is Nothing Then
                Throw New ArgumentNullException("PartDocuments", "PartDocuments can not be nothing")
            End If
            _PartDocument = value
        End Set
    End Property
    Private _Selection As INFITF.Selection
    Property Selection As INFITF.Selection
        Get
            Return _Selection
        End Get
        Set(value As INFITF.Selection)
            If value Is Nothing Then
                Throw New ArgumentNullException("Selection", "Selection can not be nothing")
            End If
            _Selection = value
        End Set
    End Property
    Public Shared Function SetInitialCATIA()
        'Dim XCATIA As CATIA_Property = New CATIA_Property
        Dim XCATIA As New CATIA_Property
        Try
            XCATIA.myCATIA = GetObject("", "CATIA.Application")
        Catch
            XCATIA.myCATIA = CreateObject("CATIA.Application")
        End Try

        XCATIA.myCATIA.Visible = True
        XCATIA.myCATIA.DisplayFileAlerts = True
        Try
            With XCATIA
                Try
                    .PartDocument = XCATIA.myCATIA.ActiveDocument
                    .Selection = .PartDocument.Selection
                Catch ex As Exception
                    .ProductDocument = XCATIA.myCATIA.ActiveDocument
                    .Selection = .ProductDocument.Selection
                End Try
                .Documents = XCATIA.myCATIA.Documents
            End With
            Return XCATIA
        Catch ex As Exception
            MessageBox.Show("Didn't find CATIA")
        End Try

        'myCATIA.DisplayFileAlerts = False
        Return Nothing
    End Function
    Public Shared Function c_Select(ByRef SelectionFrom As Selection, ByRef SelectType As Array, ByVal Comment As String) As String                 'Single Selection in CATIA
        Try
            Dim mStatus = SelectionFrom.SelectElement2(SelectType, Comment, False)        'Selection 單選指令
            Return mStatus
        Catch ex As Exception
            MessageBox.Show("Didn't Selected!!")
        End Try
    End Function
    Public Shared Function c_SelectdefaultAxis(ByRef PartDoc As PartDocument) As Selection                      'Select default Axis in Part
        Dim mPart As Part = PartDoc.Part
        Dim mSelection As Selection = PartDoc.Selection
        Dim sFilter(0)
        Dim Sel_Property As VisPropertySet
        Dim FirstResult As SelectedElement
        mSelection.Search("CatPrtSearch.AxisSystem,All")
        If mSelection.Count = 0 Then
            mSelection.Add(c_CreateAxis(PartDoc))
        Else
            mSelection.Item2(1)
            FirstResult = mSelection.Item(1)
            mSelection.Clear()
            mSelection.Add(FirstResult.Value)        'select element. value equal to object
        End If
        Return mSelection
    End Function
    Public Shared Function c_CreateAxis(ByRef PartDoc As PartDocument) As AxisSystem            'Create Axis in Part
        Dim mPart As Part = PartDoc.Part
        Dim mAxis As AxisSystems = mPart.AxisSystems
        Dim Axis1 As AxisSystem = mAxis.Add()

        Axis1.OriginType = CATAxisSystemOriginType.catAxisSystemOriginByCoordinates
        Dim ArrayOrigin(2)
        ArrayOrigin(0) = 0#
        ArrayOrigin(1) = 0#
        ArrayOrigin(2) = 0#
        Axis1.PutOrigin(ArrayOrigin)

        Axis1.XAxisType = CATAxisSystemAxisType.catAxisSystemAxisByCoordinates
        Dim ArrayXaxis(2)
        ArrayXaxis(0) = 1.0#
        ArrayXaxis(1) = 0#
        ArrayXaxis(2) = 0#
        Axis1.PutOrigin(ArrayXaxis)

        Axis1.YAxisType = CATAxisSystemAxisType.catAxisSystemAxisByCoordinates
        Dim ArrayYaxis(2)
        ArrayYaxis(0) = 0#
        ArrayYaxis(1) = 1.0#
        ArrayYaxis(2) = 0#
        Axis1.PutOrigin(ArrayYaxis)

        Axis1.ZAxisType = CATAxisSystemAxisType.catAxisSystemAxisByCoordinates
        Dim ArrayZaxis(2)
        ArrayZaxis(0) = 0#
        ArrayZaxis(1) = 0#
        ArrayZaxis(2) = 1.0#
        Axis1.PutOrigin(ArrayZaxis)

        mPart.UpdateObject(Axis1)
        Axis1.IsCurrent = True
        mPart.Update()
        Return Axis1
    End Function
    Public Shared Function ExceptSelectedItemHide(ByRef ProductDoc As ProductDocument, ByVal ItemCount As Integer)                          'Only the one which is selected show
        Dim PartCount As Integer = ProductDoc.Product.Products.Count
        Dim i As Integer
        Dim Component As Product
        For i = 1 To PartCount
            Component = ProductDoc.Product.Products.Item(i)
            If i = ItemCount Then
                Call ShowObjectFromProduct(ProductDoc, Component.Name)
            Else
                Call HideObjectFromProduct(ProductDoc, Component.Name)
            End If
        Next
    End Function
    Public Shared Function GetISO_ViewPoint(ByRef mProduct As Product, ByRef CATIA As INFITF.Application, ByVal FileLocation As String) As String       '輸出ISO視角圖片
        Dim Space As SpecsAndGeomWindow = CATIA.ActiveWindow
        Dim ActWin As Window = CATIA.ActiveWindow
        Dim ActView As Viewer3D = ActWin.ActiveViewer
        'CATIA.StartCommand("Compass")
        Space.Layout = CatSpecsAndGeomWindowLayout.catWindowGeomOnly
        ActView.FullScreen = True
        ActView.Reframe()
        ActView.Viewpoint3D = CATIA.ActiveDocument.Cameras.Item(1)
        ActView.ZoomIn()
        Dim Color(2)
        ActView.GetBackgroundColor(Color)
        Dim BlackArray(2)
        BlackArray(0) = 1
        BlackArray(1) = 1
        BlackArray(2) = 1
        ActView.PutBackgroundColor(BlackArray)
        Dim imageFilePath As String = FileLocation + "\" + mProduct.Name + ".JPG"
        ActView.CaptureToFile(CatCaptureFormat.catCaptureFormatJPEG, imageFilePath)
        ActView.PutBackgroundColor(Color)

        Space.Layout = CatSpecsAndGeomWindowLayout.catWindowSpecsAndGeom
        ActView.FullScreen = False
        'CATIA.StartCommand("Compass")
        Return imageFilePath
    End Function
    Public Shared Function MutipleBoundingBox(ByRef PartDoc As PartDocument, ByRef ProDoc As Document)                          'Create Bounding in Product
        Dim mPart As Part = PartDoc.Part           '抓取目前的CATPart檔案
        Dim Selection As Selection = PartDoc.Selection           '定義Selection
        Dim mHybridShapeFactory As HybridShapeFactory = mPart.HybridShapeFactory
        Dim AxisSyst As AxisSystem = c_SelectdefaultAxis(PartDoc).Item(1).Value             '抓取參考Axis
        AxisSyst.IsCurrent = 1
        AxisSyst.Name = "Refer_Axis"
        Dim Originpoint As HybridShapePointCoord
        Dim AxisCoord(2)

        Dim OriginCoord(2)
        AxisSyst.GetOrigin(OriginCoord)
        Originpoint = mHybridShapeFactory.AddNewPointCoord(OriginCoord(0), OriginCoord(1), OriginCoord(2))
        'Dim AxisRef = mPart.CreateReferenceFromObject(Originpoint)
        AxisSyst.GetXAxis(AxisCoord)
        Dim hybridShapeDX As HybridShapeDirection = mHybridShapeFactory.AddNewDirectionByCoord(AxisCoord(0), AxisCoord(1), AxisCoord(2))
        AxisSyst.GetYAxis(AxisCoord)
        Dim hybridShapeDY As HybridShapeDirection = mHybridShapeFactory.AddNewDirectionByCoord(AxisCoord(0), AxisCoord(1), AxisCoord(2))
        AxisSyst.GetZAxis(AxisCoord)
        Dim hybridShapeDZ As HybridShapeDirection = mHybridShapeFactory.AddNewDirectionByCoord(AxisCoord(0), AxisCoord(1), AxisCoord(2))

        Dim Plane_line_1 As Line = c_HybridShapeLinePtDir(mHybridShapeFactory, Originpoint, hybridShapeDX)
        Dim Plane_line_2 As Line = c_HybridShapeLinePtDir(mHybridShapeFactory, Originpoint, hybridShapeDY)

        Selection.Clear()

        Dim oBodies As Bodies = mPart.Bodies

        Dim j As Integer
        j = oBodies.Count
        Dim Body1 As Body = oBodies.Add()
        Body1.Name = "BoundingBox" & j
        Dim mHybridBodies As HybridBodies = Body1.HybridBodies
        Dim mHybridBody As HybridBody = mHybridBodies.Add()
        mHybridBody.Name = "definition_points"
        Dim Reference1 As Reference = CreateRefFromObj(mPart, oBodies.Item("PartBody"))

        Dim HybridShapeExtremum1 As HybridShapeExtremum = c_HybridShapeExtremum(mHybridShapeFactory, hybridShapeDX, Reference1, 1)
        Dim HybridShapeExtremum2 As HybridShapeExtremum = c_HybridShapeExtremum(mHybridShapeFactory, hybridShapeDX, Reference1, 0)
        Dim HybridShapeExtremum3 As HybridShapeExtremum = c_HybridShapeExtremum(mHybridShapeFactory, hybridShapeDY, Reference1, 1)
        Dim HybridShapeExtremum4 As HybridShapeExtremum = c_HybridShapeExtremum(mHybridShapeFactory, hybridShapeDY, Reference1, 0)
        Dim HybridShapeExtremum5 As HybridShapeExtremum = c_HybridShapeExtremum(mHybridShapeFactory, hybridShapeDZ, Reference1, 1)
        Dim HybridShapeExtremum6 As HybridShapeExtremum = c_HybridShapeExtremum(mHybridShapeFactory, hybridShapeDZ, Reference1, 0)
        mPart.Update()
        Dim DefinitionPoint As HybridBody = mHybridBodies.Item("definition_points")

        Call HideObject(PartDoc, "definition_points")                               '隱藏繪製過程


        DefinitionPoint.AppendHybridShape(HybridShapeExtremum1)
        HybridShapeExtremum1.Name = "max_X"
        DefinitionPoint.AppendHybridShape(HybridShapeExtremum2)
        HybridShapeExtremum2.Name = "min_X"
        DefinitionPoint.AppendHybridShape(HybridShapeExtremum3)
        HybridShapeExtremum3.Name = "max_Y"
        DefinitionPoint.AppendHybridShape(HybridShapeExtremum4)
        HybridShapeExtremum4.Name = "min_Y"
        DefinitionPoint.AppendHybridShape(HybridShapeExtremum5)
        HybridShapeExtremum5.Name = "max_Z"
        DefinitionPoint.AppendHybridShape(HybridShapeExtremum6)
        HybridShapeExtremum6.Name = "min_Z"

        mPart.Update()
        '建立端點
        Dim Ref1 As Reference = mPart.CreateReferenceFromObject(HybridShapeExtremum1)
        Dim Point1 As HybridShapePointCoord = mHybridShapeFactory.AddNewPointCoordWithReference(0, 0, 0, Ref1)
        DefinitionPoint.AppendHybridShape(Point1)
        Dim Point_Ref11 As Reference = mPart.CreateReferenceFromObject(Point1)
        Dim Ref2 As Reference = mPart.CreateReferenceFromObject(HybridShapeExtremum2)
        Dim Point2 As HybridShapePointCoord = mHybridShapeFactory.AddNewPointCoordWithReference(0, 0, 0, Ref2)
        DefinitionPoint.AppendHybridShape(Point2)
        Dim Point_Ref12 As Reference = mPart.CreateReferenceFromObject(Point2)
        Dim Ref3 As Reference = mPart.CreateReferenceFromObject(HybridShapeExtremum3)
        Dim Point3 As HybridShapePointCoord = mHybridShapeFactory.AddNewPointCoordWithReference(0, 0, 0, Ref3)
        DefinitionPoint.AppendHybridShape(Point3)
        Dim Point_Ref13 As Reference = mPart.CreateReferenceFromObject(Point3)
        Dim Ref4 As Reference = mPart.CreateReferenceFromObject(HybridShapeExtremum4)
        Dim Point4 As HybridShapePointCoord = mHybridShapeFactory.AddNewPointCoordWithReference(0, 0, 0, Ref4)
        DefinitionPoint.AppendHybridShape(Point4)
        Dim Point_Ref14 As Reference = mPart.CreateReferenceFromObject(Point4)
        Dim Ref5 As Reference = mPart.CreateReferenceFromObject(HybridShapeExtremum5)
        Dim Point5 As HybridShapePointCoord = mHybridShapeFactory.AddNewPointCoordWithReference(0, 0, 0, Ref5)
        DefinitionPoint.AppendHybridShape(Point5)
        Dim Point_Ref5 As Reference = mPart.CreateReferenceFromObject(Point5)
        Dim Ref6 As Reference = mPart.CreateReferenceFromObject(HybridShapeExtremum6)
        Dim Point6 As HybridShapePointCoord = mHybridShapeFactory.AddNewPointCoordWithReference(0, 0, 0, Ref6)
        DefinitionPoint.AppendHybridShape(Point6)
        Dim Point_Ref6 As Reference = mPart.CreateReferenceFromObject(Point6)
        mPart.Update()

        AxisSyst.IsCurrent = 1
        '建立邊界盒草圖
        Dim Sketches1 = DefinitionPoint.HybridSketches
        Dim Reference_axis_syst As Reference = mPart.CreateReferenceFromName("Selection_RSur:(Face:(Brp:(Refer_Axis;1);None:());Refer_Axis)")
        Dim Standard_Body_Sketch1 As Sketch = Sketches1.Add(Reference_axis_syst)
        Dim Factory2D1 As Factory2D = Standard_Body_Sketch1.OpenEdition()
        Dim GeometricElements1 = Standard_Body_Sketch1.GeometricElements
        Dim Axis2D1 = GeometricElements1.Item("AbsoluteAxis")
        Dim Line_HDirection As Line2D = Axis2D1.GetItem("HDirection")
        Line_HDirection.ReportName = 1
        Dim Line_VDirection As Line2D = Axis2D1.GetItem("VDirection")
        Line_VDirection.ReportName = 2

        Dim Parameter As Double = 20000
        Dim Point_ref_1 As Point2D = Factory2D1.CreatePoint(-Parameter, -Parameter)
        Point_ref_1.ReportName = 3
        Dim Point_ref_2 As Point2D = Factory2D1.CreatePoint(Parameter, -Parameter)
        Point_ref_2.ReportName = 4
        Dim Point_ref_3 As Point2D = Factory2D1.CreatePoint(Parameter, Parameter)
        Point_ref_3.ReportName = 5
        Dim Point_ref_4 As Point2D = Factory2D1.CreatePoint(-Parameter, Parameter)
        Point_ref_4.ReportName = 6

        Dim Line_ref_1_2 = Factory2D1.CreateLine(-Parameter, -Parameter, Parameter, -Parameter)
        Line_ref_1_2.ReportName = 7
        Line_ref_1_2.StartPoint = Point_ref_1
        Line_ref_1_2.EndPoint = Point_ref_2

        Dim Line_ref_2_3 = Factory2D1.CreateLine(Parameter, -Parameter, Parameter, Parameter)
        Line_ref_2_3.ReportName = 8
        Line_ref_2_3.StartPoint = Point_ref_2
        Line_ref_2_3.EndPoint = Point_ref_3

        Dim Line_ref_3_4 = Factory2D1.CreateLine(-Parameter, Parameter, Parameter, Parameter)
        Line_ref_3_4.ReportName = 9
        Line_ref_3_4.StartPoint = Point_ref_3
        Line_ref_3_4.EndPoint = Point_ref_4

        Dim Line_ref_4_1 = Factory2D1.CreateLine(-Parameter, -Parameter, -Parameter, Parameter)
        Line_ref_4_1.ReportName = 10
        Line_ref_4_1.StartPoint = Point_ref_4
        Line_ref_4_1.EndPoint = Point_ref_1

        Dim reference_Line_1_2 = mPart.CreateReferenceFromObject(Line_ref_1_2)
        Dim reference_Line_2_3 = mPart.CreateReferenceFromObject(Line_ref_2_3)
        Dim reference_Line_3_4 = mPart.CreateReferenceFromObject(Line_ref_3_4)
        Dim reference_Line_4_1 = mPart.CreateReferenceFromObject(Line_ref_4_1)
        Dim Electrode_constraints As Constraints = Standard_Body_Sketch1.Constraints


        Dim Constraint_toto_2 = Electrode_constraints.AddBiEltCst(CatConstraintType.catCstTypeDistance, Point_Ref11, reference_Line_2_3)
        Dim Constraint_toto_3 = Electrode_constraints.AddBiEltCst(CatConstraintType.catCstTypeDistance, Point_Ref13, reference_Line_3_4)
        Dim Constraint_toto_4 = Electrode_constraints.AddBiEltCst(CatConstraintType.catCstTypeDistance, reference_Line_4_1, Point_Ref12)
        Dim Constraint_toto_1 = Electrode_constraints.AddBiEltCst(CatConstraintType.catCstTypeDistance, reference_Line_1_2, Point_Ref14)

        Dim Length1 As Dimension = Constraint_toto_1.Dimension
        Length1.Value = 0
        Dim Length2 As Dimension = Constraint_toto_2.Dimension
        Length2.Value = 0
        Dim Length3 As Dimension = Constraint_toto_3.Dimension
        Length3.Value = 0
        Dim Length4 As Dimension = Constraint_toto_4.Dimension
        Length4.Value = 0

        Standard_Body_Sketch1.CloseEdition()
        mPart.Update()

        '建立Z軸兩端平面
        Dim Plan_inferieur As HybridShapePlaneOffsetPt
        Dim Plan_orgin As HybridShapePlane2Lines
        Dim Origin_line_1 As Reference = mPart.CreateReferenceFromObject(Line_HDirection)
        Dim Origin_line_2 As Reference = mPart.CreateReferenceFromObject(Line_VDirection)

        Plan_orgin = mHybridShapeFactory.AddNewPlane2Lines(Origin_line_1, Origin_line_2)
        Dim Ref_Plan_origin As Reference = mPart.CreateReferenceFromObject(Plan_orgin)
        Plan_inferieur = mHybridShapeFactory.AddNewPlaneOffsetPt(Ref_Plan_origin, Point_Ref6)
        DefinitionPoint.AppendHybridShape(Plan_inferieur)

        Dim Plan_superieur As HybridShapePlaneOffsetPt = mHybridShapeFactory.AddNewPlaneOffsetPt(Ref_Plan_origin, Point_Ref5)
        DefinitionPoint.AppendHybridShape(Plan_superieur)

        mPart.Update()

        Dim Point_inf As HybridShapePointCoord = mHybridShapeFactory.AddNewPointCoordWithReference(0, 0, 0, Point_Ref6)
        DefinitionPoint.AppendHybridShape(Point_inf)
        Dim ref_point_inf As Reference = mPart.CreateReferenceFromObject(Point_inf)

        Dim proj_pt_inf As HybridShapeProject = mHybridShapeFactory.AddNewProject(Point_Ref6, Plan_superieur)
        DefinitionPoint.AppendHybridShape(proj_pt_inf)

        Dim Point_sup As HybridShapePointCoord = mHybridShapeFactory.AddNewPointCoordWithReference(0, 0, 0, proj_pt_inf)
        DefinitionPoint.AppendHybridShape(Point_sup)
        Dim ref_Point_sup As Reference = mPart.CreateReferenceFromObject(Point_sup)

        Dim Line_guide As HybridShapeLinePtPt = mHybridShapeFactory.AddNewLinePtPt(ref_point_inf, ref_Point_sup)
        DefinitionPoint.AppendHybridShape(Line_guide)
        Dim ref_guideline = mPart.CreateReferenceFromObject(Line_guide)

        Dim oStart As Length = Line_guide.BeginOffset
        oStart.Value = 0
        Dim oEnd As Length = Line_guide.EndOffset
        oEnd.Value = 0

        Dim Constraints_for_Z = mPart.Constraints
        Dim Constraint_dz = Constraints_for_Z.AddMonoEltCst(CatConstraintType.catCstTypeLength, ref_guideline)
        Dim Length_dz = Constraint_dz.Dimension

        mPart.Update()

        Dim Sketch2 = mHybridBody.HybridSketches
        Dim Standard_body_sketch2 = Sketch2.Add(Plan_inferieur)

        Dim Factory2D2 As Factory2D = Standard_body_sketch2.OpenEdition
        Dim GeometricElements2 = Standard_body_sketch2.GeometricElements

        Dim pont As Double = 200000
        Dim point_ref1_1 = Factory2D2.CreatePoint(-pont, -pont)
        Dim point_ref1_2 = Factory2D2.CreatePoint(pont, -pont)
        Dim point_ref1_3 = Factory2D2.CreatePoint(pont, pont)
        Dim point_ref1_4 = Factory2D2.CreatePoint(-pont, pont)

        Dim Line_ref1_1_2 = Factory2D2.CreateLine(-pont, -pont, pont, -pont)
        Line_ref1_1_2.StartPoint = point_ref1_1
        Line_ref1_1_2.EndPoint = point_ref1_2

        Dim Line_ref1_2_3 = Factory2D2.CreateLine(pont, -pont, pont, pont)
        Line_ref1_2_3.StartPoint = point_ref1_2
        Line_ref1_2_3.EndPoint = point_ref1_3

        Dim Line_ref1_3_4 = Factory2D2.CreateLine(-pont, pont, pont, pont)
        Line_ref1_3_4.StartPoint = point_ref1_4
        Line_ref1_3_4.EndPoint = point_ref1_3

        Dim Line_ref1_4_1 = Factory2D2.CreateLine(-pont, -pont, -pont, pont)
        Line_ref1_4_1.StartPoint = point_ref1_1
        Line_ref1_4_1.EndPoint = point_ref1_4

        Dim Reference_line_ref1_1_2 = CreateRefFromObj(mPart, Line_ref1_1_2)
        Dim Reference_line_ref1_2_3 = CreateRefFromObj(mPart, Line_ref1_2_3)
        Dim Reference_line_ref1_3_4 = CreateRefFromObj(mPart, Line_ref1_3_4)
        Dim Reference_line_ref1_4_1 = CreateRefFromObj(mPart, Line_ref1_4_1)

        Dim Proj_1_2 As Geometry2D = Factory2D2.CreateProjection(reference_Line_1_2)
        Dim Proj_2_3 As Geometry2D = Factory2D2.CreateProjection(reference_Line_2_3)
        Dim Proj_3_4 As Geometry2D = Factory2D2.CreateProjection(reference_Line_3_4)
        Dim Proj_4_1 As Geometry2D = Factory2D2.CreateProjection(reference_Line_4_1)

        Dim Ref_Line_sk1_1_2 = CreateRefFromObj(mPart, Proj_1_2)
        Dim Ref_Line_sk1_2_3 = CreateRefFromObj(mPart, Proj_2_3)
        Dim Ref_Line_sk1_3_4 = CreateRefFromObj(mPart, Proj_3_4)
        Dim Ref_Line_sk1_4_1 = CreateRefFromObj(mPart, Proj_4_1)

        Electrode_constraints = Standard_body_sketch2.Constraints
        Dim constraint_toto_11 = Electrode_constraints.AddBiEltCst(CatConstraintType.catCstTypeDistance, Reference_line_ref1_1_2, Ref_Line_sk1_1_2)
        Dim constraint_toto_12 = Electrode_constraints.AddBiEltCst(CatConstraintType.catCstTypeDistance, Ref_Line_sk1_2_3, Reference_line_ref1_2_3)
        Dim constraint_toto_13 = Electrode_constraints.AddBiEltCst(CatConstraintType.catCstTypeDistance, Ref_Line_sk1_3_4, Reference_line_ref1_3_4)
        Dim constraint_toto_14 = Electrode_constraints.AddBiEltCst(CatConstraintType.catCstTypeDistance, Line_ref1_4_1, Ref_Line_sk1_4_1)

        Dim Length11 As Dimension = constraint_toto_11.Dimension
        Length11.Value = 0
        Dim Length12 As Dimension = constraint_toto_12.Dimension
        Length12.Value = 0
        Dim Length13 As Dimension = constraint_toto_13.Dimension
        Length13.Value = 0
        Dim Length14 As Dimension = constraint_toto_14.Dimension
        Length14.Value = 0

        Standard_body_sketch2.CloseEdition()
        mPart.Update()

        'create third sketch

        Dim Sketch3 = DefinitionPoint.HybridSketches
        Dim Standard_body_Sketch3 = Sketch3.Add(Plan_inferieur)
        Dim Factory2D3 = Standard_body_Sketch3.OpenEdition
        Dim GeometricElements3 = Standard_body_Sketch3.GeometricElements

        Dim proj1_1_2 = Factory2D3.CreateProjection(Reference_line_ref1_1_2)
        Dim proj1_2_3 = Factory2D3.CreateProjection(Reference_line_ref1_2_3)
        Dim proj1_3_4 = Factory2D3.CreateProjection(Reference_line_ref1_3_4)
        Dim proj1_4_1 = Factory2D3.CreateProjection(Reference_line_ref1_4_1)

        Dim Ref_proj1_1_2 = CreateRefFromObj(mPart, proj1_1_2)
        Dim Ref_proj1_2_3 = CreateRefFromObj(mPart, proj1_2_3)
        Dim Ref_proj1_3_4 = CreateRefFromObj(mPart, proj1_3_4)
        Dim Ref_proj1_4_1 = CreateRefFromObj(mPart, proj1_4_1)

        Dim Constraints_dim = Standard_body_Sketch3.Constraints
        Dim constraint_dx = Constraints_dim.AddMonoEltCst(CatConstraintType.catCstTypeLength, Ref_proj1_1_2)
        Dim constraint_dy = Constraints_dim.AddMonoEltCst(CatConstraintType.catCstTypeLength, Ref_proj1_2_3)
        Dim Length_dx = constraint_dx.Dimension
        Dim dx_value = Length_dx.Value
        Dim Length_dy = constraint_dy.Dimension
        Dim dy_value = Length_dy.Value

        Standard_body_Sketch3.CloseEdition()
        mPart.Update()


        'Create Formula
        Dim string_1 As String = "Offset_Bbox_Max_X." & j
        Dim string_2 As String = "Offset_Bbox_Min_X." & j
        Dim string_3 As String = "Offset_Bbox_Max_Y." & j
        Dim string_4 As String = "Offset_Bbox_Min_Y." & j
        Dim string_5 As String = "Offset_Bbox_Max_Z." & j
        Dim string_6 As String = "Offset_Bbox_Min_Z." & j
        Dim string_7 As String = "Bbox_dx." & j
        Dim string_8 As String = "Bbox_dy." & j
        Dim string_9 As String = "Bbox_dz." & j



        Dim Offset_Bbox_Max_X As RealParam = mPart.Parameters.CreateDimension(string_1, "Length", 0)
        Dim Offset_Bbox_Min_X As RealParam = mPart.Parameters.CreateDimension(string_2, "Length", 0)
        Dim Offset_Bbox_Max_Y As RealParam = mPart.Parameters.CreateDimension(string_3, "Length", 0)
        Dim Offset_Bbox_Min_Y As RealParam = mPart.Parameters.CreateDimension(string_4, "Length", 0)
        Dim Offset_Bbox_Max_Z As RealParam = mPart.Parameters.CreateDimension(string_5, "Length", 0)
        Dim Offset_Bbox_Min_Z As RealParam = mPart.Parameters.CreateDimension(string_6, "Length", 0)

        Dim Bbox_dx_hidden = mPart.Parameters.CreateDimension(string_7, "Length", dx_value)
        Dim Bbox_dy_hidden = mPart.Parameters.CreateDimension(string_8, "Length", dy_value)
        Dim Bbox_dz_hidden = mPart.Parameters.CreateDimension(string_9, "Length", Length_dz.Value)
        Bbox_dx_hidden.Hidden = True
        Bbox_dy_hidden.Hidden = True
        Bbox_dz_hidden.Hidden = True

        Dim Bbox_dx = mPart.Parameters.CreateDimension(string_7, "Length", 0)
        Dim Bbox_dy = mPart.Parameters.CreateDimension(string_8, "Length", 0)
        Dim Bbox_dz = mPart.Parameters.CreateDimension(string_9, "Length", 0)

        Dim Formula_1 As Formula = mPart.Relations.CreateFormula("formula_Bbox_1." & j, "", Length14, string_2)
        Dim Formula_2 As Formula = mPart.Relations.CreateFormula("formula_Bbox_2." & j, "", Length12, string_1)
        Dim Formula_3 As Formula = mPart.Relations.CreateFormula("formula_Bbox_3." & j, "", Length11, string_4)
        Dim Formula_4 As Formula = mPart.Relations.CreateFormula("formula_Bbox_4." & j, "", Length13, string_3)
        Dim Formula_5 As Formula = mPart.Relations.CreateFormula("formula_Bbox_5." & j, "", oEnd, string_6)
        Dim Formula_6 As Formula = mPart.Relations.CreateFormula("formula_Bbox_6." & j, "", oStart, string_5)

        Dim Formula_7 As Formula = mPart.Relations.CreateFormula("formula_Bbox_7." & j, "", Bbox_dx, "Bbox_dx." & j & "+Offset_Bbox_Min_X." & j & "+Offset_Bbox_Max_X." & j & "-0mm")
        Dim Formula_8 As Formula = mPart.Relations.CreateFormula("formula_Bbox_8." & j, "", Bbox_dy, "Bbox_dy." & j & "+Offset_Bbox_Min_Y." & j & "+Offset_Bbox_Max_Y." & j & "-0mm")
        Dim Formula_9 As Formula = mPart.Relations.CreateFormula("formula_Bbox_9." & j, "", Bbox_dz, "Bbox_dz." & j & "+Offset_Bbox_Min_Z." & j & "+Offset_Bbox_Max_Z." & j & "-0mm")

        mPart.Update()


        'Creation Sweep
        Dim SweepRef_1 = CreateRefFromObj(mPart, Standard_body_Sketch3)
        Dim Guide1 = CreateRefFromObj(mPart, Line_guide)
        Dim Sweep1 As HybridShapeSweepExplicit = mHybridShapeFactory.AddNewSweepExplicit(SweepRef_1, Guide1)
        DefinitionPoint.AppendHybridShape(Sweep1)
        mPart.Update()

        '將建立的面Join

        Dim mShapeFactory As ShapeFactory = mPart.ShapeFactory
        Dim HybridShapes1 As HybridShapes = DefinitionPoint.HybridShapes

        Dim Refer1 As Reference = CreateRefFromObj(mPart, Sweep1)

        Dim CloseSurface1 As CloseSurface = mShapeFactory.AddNewCloseSurface(Refer1)
        CloseSurface1.Name = "BoundingBox"
        mPart.Update()
        With PartDoc.Selection
            .Clear()
            .Add(Body1)
            .VisProperties.SetRealColor(255, 255, 128, 0)
            .VisProperties.SetRealOpacity(150, 1)
            .VisProperties.SetRealWidth(4, 1)
        End With

        PartDoc.Selection.Clear()
        mPart.Update()

    End Function
    Public Shared Function BoundingBox(ByRef PartDoc As PartDocument)                       'Create Bounding Box in Part
        'If (InStr(CATIA.PartDocuments.Name, ".CATPart")) <> 0 Then
        Dim mPart As Part = PartDoc.Part           '抓取目前的CATPart檔案
        Dim Selection As Selection = PartDoc.Selection           '定義Selection
        mPart.Update()
        Dim mHybridShapeFactory As HybridShapeFactory = mPart.HybridShapeFactory
        Dim sFilter(0)
        'MessageBox.Show("Select a local axis")
        sFilter(0) = "AxisSystem"
        Dim Instruction As String = "select a local axis"
        'Dim sStatus = Selection.SelectElement2(sFilter, "select a local axis", False)        'Selection 單選指令
        Dim sStatus = c_Select(PartDoc.Selection, sFilter, Instruction)
        Dim AxisCoord(2)
        Dim AxisSyst As AxisSystem
        AxisSyst = Selection.Item(1).Value
        Dim AxisRef = Selection.Item(1).Value
        Dim Ref_Name_systaxis = AxisSyst.Name
        AxisSyst.IsCurrent = 1
        AxisSyst.Name = "Refer_Axis"
        Dim Axname = AxisSyst.Name
        Dim Originpoint As HybridShapePointCoord

        Dim OriginCoord(2)
        AxisSyst.GetOrigin(OriginCoord)
        Originpoint = mHybridShapeFactory.AddNewPointCoord(OriginCoord(0), OriginCoord(1), OriginCoord(2))
        AxisRef = mPart.CreateReferenceFromObject(Originpoint)
        AxisSyst.GetXAxis(AxisCoord)
        Dim hybridShapeDX As HybridShapeDirection = mHybridShapeFactory.AddNewDirectionByCoord(AxisCoord(0), AxisCoord(1), AxisCoord(2))
        AxisSyst.GetYAxis(AxisCoord)
        Dim hybridShapeDY As HybridShapeDirection = mHybridShapeFactory.AddNewDirectionByCoord(AxisCoord(0), AxisCoord(1), AxisCoord(2))
        AxisSyst.GetZAxis(AxisCoord)
        Dim hybridShapeDZ As HybridShapeDirection = mHybridShapeFactory.AddNewDirectionByCoord(AxisCoord(0), AxisCoord(1), AxisCoord(2))

        Dim Plane_line_1 As Line = c_HybridShapeLinePtDir(mHybridShapeFactory, Originpoint, hybridShapeDX)
        Dim Plane_line_2 As Line = c_HybridShapeLinePtDir(mHybridShapeFactory, Originpoint, hybridShapeDY)

        Selection.Clear()

        Dim oBodies As Bodies = mPart.Bodies

        Dim j As Integer
        j = oBodies.Count
        Dim Body1 As Body = oBodies.Add()
        Body1.Name = "BoundingBox" & j
        Dim HybridBodies1 As HybridBodies = Body1.HybridBodies
        Dim HybridBody1 As HybridBody = HybridBodies1.Add()
        HybridBody1.Name = "definition_points"

        ReDim sFilter(0)
        'MessageBox.Show("Select a Face")
        sFilter(0) = "Face"
        sStatus = Selection.SelectElement2(sFilter, "select a face", False)
        If (sStatus = "Cancel") Then
            Exit Function
        End If

        Dim Reference1 As Reference = Selection.Item(1).Value
        Dim HybridShapeExtract1 As HybridShapeExtract = mHybridShapeFactory.AddNewExtract(Reference1)
        HybridShapeExtract1.PropagationType = 1
        HybridShapeExtract1.ComplementaryExtract = False
        HybridShapeExtract1.IsFederated = False
        Reference1 = HybridShapeExtract1

        '建立各個面的極端線
        Dim HybridShapeExtremum1 As HybridShapeExtremum = c_HybridShapeExtremum(mHybridShapeFactory, hybridShapeDX, Reference1, 1)
        Dim HybridShapeExtremum2 As HybridShapeExtremum = c_HybridShapeExtremum(mHybridShapeFactory, hybridShapeDX, Reference1, 0)
        Dim HybridShapeExtremum3 As HybridShapeExtremum = c_HybridShapeExtremum(mHybridShapeFactory, hybridShapeDY, Reference1, 1)
        Dim HybridShapeExtremum4 As HybridShapeExtremum = c_HybridShapeExtremum(mHybridShapeFactory, hybridShapeDY, Reference1, 0)
        Dim HybridShapeExtremum5 As HybridShapeExtremum = c_HybridShapeExtremum(mHybridShapeFactory, hybridShapeDZ, Reference1, 1)
        Dim HybridShapeExtremum6 As HybridShapeExtremum = c_HybridShapeExtremum(mHybridShapeFactory, hybridShapeDZ, Reference1, 0)
        mPart.Update()

        Dim HybridBody2 As HybridBody = HybridBodies1.Item("definition_points")

        Call HideObject(PartDoc, "definition_points")                               '隱藏繪製過程


        HybridBody2.AppendHybridShape(HybridShapeExtremum1)
        HybridShapeExtremum1.Name = "max_X"
        HybridBody2.AppendHybridShape(HybridShapeExtremum2)
        HybridShapeExtremum2.Name = "min_X"
        HybridBody2.AppendHybridShape(HybridShapeExtremum3)
        HybridShapeExtremum3.Name = "max_Y"
        HybridBody2.AppendHybridShape(HybridShapeExtremum4)
        HybridShapeExtremum4.Name = "min_Y"
        HybridBody2.AppendHybridShape(HybridShapeExtremum5)
        HybridShapeExtremum5.Name = "max_Z"
        HybridBody2.AppendHybridShape(HybridShapeExtremum6)
        HybridShapeExtremum6.Name = "min_Z"

        mPart.Update()
        '建立端點
        Dim Ref1 As Reference = mPart.CreateReferenceFromObject(HybridShapeExtremum1)
        Dim Point1 As HybridShapePointCoord = mHybridShapeFactory.AddNewPointCoordWithReference(0, 0, 0, Ref1)
        HybridBody2.AppendHybridShape(Point1)
        Dim Point_Ref11 As Reference = mPart.CreateReferenceFromObject(Point1)
        Dim Ref2 As Reference = mPart.CreateReferenceFromObject(HybridShapeExtremum2)
        Dim Point2 As HybridShapePointCoord = mHybridShapeFactory.AddNewPointCoordWithReference(0, 0, 0, Ref2)
        HybridBody2.AppendHybridShape(Point2)
        Dim Point_Ref12 As Reference = mPart.CreateReferenceFromObject(Point2)
        Dim Ref3 As Reference = mPart.CreateReferenceFromObject(HybridShapeExtremum3)
        Dim Point3 As HybridShapePointCoord = mHybridShapeFactory.AddNewPointCoordWithReference(0, 0, 0, Ref3)
        HybridBody2.AppendHybridShape(Point3)
        Dim Point_Ref13 As Reference = mPart.CreateReferenceFromObject(Point3)
        Dim Ref4 As Reference = mPart.CreateReferenceFromObject(HybridShapeExtremum4)
        Dim Point4 As HybridShapePointCoord = mHybridShapeFactory.AddNewPointCoordWithReference(0, 0, 0, Ref4)
        HybridBody2.AppendHybridShape(Point4)
        Dim Point_Ref14 As Reference = mPart.CreateReferenceFromObject(Point4)
        Dim Ref5 As Reference = mPart.CreateReferenceFromObject(HybridShapeExtremum5)
        Dim Point5 As HybridShapePointCoord = mHybridShapeFactory.AddNewPointCoordWithReference(0, 0, 0, Ref5)
        HybridBody2.AppendHybridShape(Point5)
        Dim Point_Ref5 As Reference = mPart.CreateReferenceFromObject(Point5)
        Dim Ref6 As Reference = mPart.CreateReferenceFromObject(HybridShapeExtremum6)
        Dim Point6 As HybridShapePointCoord = mHybridShapeFactory.AddNewPointCoordWithReference(0, 0, 0, Ref6)
        HybridBody2.AppendHybridShape(Point6)
        Dim Point_Ref6 As Reference = mPart.CreateReferenceFromObject(Point6)
        mPart.Update()

        AxisSyst.IsCurrent = 1
        '建立邊界盒草圖
        Dim Sketches1 = HybridBody1.HybridSketches
        Dim Reference_axis_syst As Reference = mPart.CreateReferenceFromName("Selection_RSur:(Face:(Brp:(Refer_Axis;1);None:());Refer_Axis)")
        Dim Standard_Body_Sketch1 As Sketch = Sketches1.Add(Reference_axis_syst)
        Dim Factory2D1 As Factory2D = Standard_Body_Sketch1.OpenEdition()
        Dim GeometricElements1 = Standard_Body_Sketch1.GeometricElements
        Dim Axis2D1 = GeometricElements1.Item("AbsoluteAxis")
        Dim Line_HDirection As Line2D = Axis2D1.GetItem("HDirection")
        Line_HDirection.ReportName = 1
        Dim Line_VDirection As Line2D = Axis2D1.GetItem("VDirection")
        Line_VDirection.ReportName = 2

        Dim Parameter As Double = 20000
        Dim Point_ref_1 As Point2D = Factory2D1.CreatePoint(-Parameter, -Parameter)
        Point_ref_1.ReportName = 3
        Dim Point_ref_2 As Point2D = Factory2D1.CreatePoint(Parameter, -Parameter)
        Point_ref_2.ReportName = 4
        Dim Point_ref_3 As Point2D = Factory2D1.CreatePoint(Parameter, Parameter)
        Point_ref_3.ReportName = 5
        Dim Point_ref_4 As Point2D = Factory2D1.CreatePoint(-Parameter, Parameter)
        Point_ref_4.ReportName = 6

        Dim Line_ref_1_2 = Factory2D1.CreateLine(-Parameter, -Parameter, Parameter, -Parameter)
        Line_ref_1_2.ReportName = 7
        Line_ref_1_2.StartPoint = Point_ref_1
        Line_ref_1_2.EndPoint = Point_ref_2

        Dim Line_ref_2_3 = Factory2D1.CreateLine(Parameter, -Parameter, Parameter, Parameter)
        Line_ref_2_3.ReportName = 8
        Line_ref_2_3.StartPoint = Point_ref_2
        Line_ref_2_3.EndPoint = Point_ref_3

        Dim Line_ref_3_4 = Factory2D1.CreateLine(-Parameter, Parameter, Parameter, Parameter)
        Line_ref_3_4.ReportName = 9
        Line_ref_3_4.StartPoint = Point_ref_3
        Line_ref_3_4.EndPoint = Point_ref_4

        Dim Line_ref_4_1 = Factory2D1.CreateLine(-Parameter, -Parameter, -Parameter, Parameter)
        Line_ref_4_1.ReportName = 10
        Line_ref_4_1.StartPoint = Point_ref_4
        Line_ref_4_1.EndPoint = Point_ref_1

        Dim reference_Line_1_2 = mPart.CreateReferenceFromObject(Line_ref_1_2)
        Dim reference_Line_2_3 = mPart.CreateReferenceFromObject(Line_ref_2_3)
        Dim reference_Line_3_4 = mPart.CreateReferenceFromObject(Line_ref_3_4)
        Dim reference_Line_4_1 = mPart.CreateReferenceFromObject(Line_ref_4_1)
        Dim Electrode_constraints As Constraints = Standard_Body_Sketch1.Constraints


        Dim Constraint_toto_2 = Electrode_constraints.AddBiEltCst(CatConstraintType.catCstTypeDistance, Point_Ref11, reference_Line_2_3)
        Dim Constraint_toto_3 = Electrode_constraints.AddBiEltCst(CatConstraintType.catCstTypeDistance, Point_Ref13, reference_Line_3_4)
        Dim Constraint_toto_4 = Electrode_constraints.AddBiEltCst(CatConstraintType.catCstTypeDistance, reference_Line_4_1, Point_Ref12)
        Dim Constraint_toto_1 = Electrode_constraints.AddBiEltCst(CatConstraintType.catCstTypeDistance, reference_Line_1_2, Point_Ref14)

        Dim Length1 As Dimension = Constraint_toto_1.Dimension
        Length1.Value = 0
        Dim Length2 As Dimension = Constraint_toto_2.Dimension
        Length2.Value = 0
        Dim Length3 As Dimension = Constraint_toto_3.Dimension
        Length3.Value = 0
        Dim Length4 As Dimension = Constraint_toto_4.Dimension
        Length4.Value = 0

        Standard_Body_Sketch1.CloseEdition()
        mPart.Update()

        '建立Z軸兩端平面
        Dim Plan_inferieur As HybridShapePlaneOffsetPt
        Dim Plan_orgin As HybridShapePlane2Lines
        Dim Origin_line_1 As Reference = mPart.CreateReferenceFromObject(Line_HDirection)
        Dim Origin_line_2 As Reference = mPart.CreateReferenceFromObject(Line_VDirection)

        Plan_orgin = mHybridShapeFactory.AddNewPlane2Lines(Origin_line_1, Origin_line_2)
        Dim Ref_Plan_origin As Reference = mPart.CreateReferenceFromObject(Plan_orgin)
        Plan_inferieur = mHybridShapeFactory.AddNewPlaneOffsetPt(Ref_Plan_origin, Point_Ref6)
        HybridBody2.AppendHybridShape(Plan_inferieur)

        Dim Plan_superieur As HybridShapePlaneOffsetPt = mHybridShapeFactory.AddNewPlaneOffsetPt(Ref_Plan_origin, Point_Ref5)
        HybridBody2.AppendHybridShape(Plan_superieur)

        mPart.Update()

        Dim Point_inf As HybridShapePointCoord = mHybridShapeFactory.AddNewPointCoordWithReference(0, 0, 0, Point_Ref6)
        HybridBody2.AppendHybridShape(Point_inf)
        Dim ref_point_inf As Reference = mPart.CreateReferenceFromObject(Point_inf)

        Dim proj_pt_inf As HybridShapeProject = mHybridShapeFactory.AddNewProject(Point_Ref6, Plan_superieur)
        HybridBody2.AppendHybridShape(proj_pt_inf)

        Dim Point_sup As HybridShapePointCoord = mHybridShapeFactory.AddNewPointCoordWithReference(0, 0, 0, proj_pt_inf)
        HybridBody2.AppendHybridShape(Point_sup)
        Dim ref_Point_sup As Reference = mPart.CreateReferenceFromObject(Point_sup)

        Dim Line_guide As HybridShapeLinePtPt = mHybridShapeFactory.AddNewLinePtPt(ref_point_inf, ref_Point_sup)
        HybridBody2.AppendHybridShape(Line_guide)
        Dim ref_guideline = mPart.CreateReferenceFromObject(Line_guide)

        Dim oStart As Length = Line_guide.BeginOffset
        oStart.Value = 0
        Dim oEnd As Length = Line_guide.EndOffset
        oEnd.Value = 0

        Dim Constraints_for_Z = mPart.Constraints
        Dim Constraint_dz = Constraints_for_Z.AddMonoEltCst(CatConstraintType.catCstTypeLength, ref_guideline)
        Dim Length_dz = Constraint_dz.Dimension

        mPart.Update()

        Dim Sketch2 = HybridBody1.HybridSketches
        Dim Standard_body_sketch2 = Sketch2.Add(Plan_inferieur)

        Dim Factory2D2 As Factory2D = Standard_body_sketch2.OpenEdition
        Dim GeometricElements2 = Standard_body_sketch2.GeometricElements

        Dim pont As Double = 200000
        Dim point_ref1_1 = Factory2D2.CreatePoint(-pont, -pont)
        Dim point_ref1_2 = Factory2D2.CreatePoint(pont, -pont)
        Dim point_ref1_3 = Factory2D2.CreatePoint(pont, pont)
        Dim point_ref1_4 = Factory2D2.CreatePoint(-pont, pont)

        Dim Line_ref1_1_2 = Factory2D2.CreateLine(-pont, -pont, pont, -pont)
        Line_ref1_1_2.StartPoint = point_ref1_1
        Line_ref1_1_2.EndPoint = point_ref1_2

        Dim Line_ref1_2_3 = Factory2D2.CreateLine(pont, -pont, pont, pont)
        Line_ref1_2_3.StartPoint = point_ref1_2
        Line_ref1_2_3.EndPoint = point_ref1_3

        Dim Line_ref1_3_4 = Factory2D2.CreateLine(-pont, pont, pont, pont)
        Line_ref1_3_4.StartPoint = point_ref1_4
        Line_ref1_3_4.EndPoint = point_ref1_3

        Dim Line_ref1_4_1 = Factory2D2.CreateLine(-pont, -pont, -pont, pont)
        Line_ref1_4_1.StartPoint = point_ref1_1
        Line_ref1_4_1.EndPoint = point_ref1_4

        Dim Reference_line_ref1_1_2 = CreateRefFromObj(mPart, Line_ref1_1_2)
        Dim Reference_line_ref1_2_3 = CreateRefFromObj(mPart, Line_ref1_2_3)
        Dim Reference_line_ref1_3_4 = CreateRefFromObj(mPart, Line_ref1_3_4)
        Dim Reference_line_ref1_4_1 = CreateRefFromObj(mPart, Line_ref1_4_1)

        Dim Proj_1_2 As Geometry2D = Factory2D2.CreateProjection(reference_Line_1_2)
        Dim Proj_2_3 As Geometry2D = Factory2D2.CreateProjection(reference_Line_2_3)
        Dim Proj_3_4 As Geometry2D = Factory2D2.CreateProjection(reference_Line_3_4)
        Dim Proj_4_1 As Geometry2D = Factory2D2.CreateProjection(reference_Line_4_1)

        Dim Ref_Line_sk1_1_2 = CreateRefFromObj(mPart, Proj_1_2)
        Dim Ref_Line_sk1_2_3 = CreateRefFromObj(mPart, Proj_2_3)
        Dim Ref_Line_sk1_3_4 = CreateRefFromObj(mPart, Proj_3_4)
        Dim Ref_Line_sk1_4_1 = CreateRefFromObj(mPart, Proj_4_1)

        Electrode_constraints = Standard_body_sketch2.Constraints
        Dim constraint_toto_11 = Electrode_constraints.AddBiEltCst(CatConstraintType.catCstTypeDistance, Reference_line_ref1_1_2, Ref_Line_sk1_1_2)
        Dim constraint_toto_12 = Electrode_constraints.AddBiEltCst(CatConstraintType.catCstTypeDistance, Ref_Line_sk1_2_3, Reference_line_ref1_2_3)
        Dim constraint_toto_13 = Electrode_constraints.AddBiEltCst(CatConstraintType.catCstTypeDistance, Ref_Line_sk1_3_4, Reference_line_ref1_3_4)
        Dim constraint_toto_14 = Electrode_constraints.AddBiEltCst(CatConstraintType.catCstTypeDistance, Line_ref1_4_1, Ref_Line_sk1_4_1)

        Dim Length11 As Dimension = constraint_toto_11.Dimension
        Length11.Value = 0
        Dim Length12 As Dimension = constraint_toto_12.Dimension
        Length12.Value = 0
        Dim Length13 As Dimension = constraint_toto_13.Dimension
        Length13.Value = 0
        Dim Length14 As Dimension = constraint_toto_14.Dimension
        Length14.Value = 0

        Standard_body_sketch2.CloseEdition()
        mPart.Update()

        'create third sketch

        Dim Sketch3 = HybridBody1.HybridSketches
        Dim Standard_body_Sketch3 = Sketch3.Add(Plan_inferieur)
        Dim Factory2D3 = Standard_body_Sketch3.OpenEdition
        Dim GeometricElements3 = Standard_body_Sketch3.GeometricElements

        Dim proj1_1_2 = Factory2D3.CreateProjection(Reference_line_ref1_1_2)
        Dim proj1_2_3 = Factory2D3.CreateProjection(Reference_line_ref1_2_3)
        Dim proj1_3_4 = Factory2D3.CreateProjection(Reference_line_ref1_3_4)
        Dim proj1_4_1 = Factory2D3.CreateProjection(Reference_line_ref1_4_1)

        Dim Ref_proj1_1_2 = CreateRefFromObj(mPart, proj1_1_2)
        Dim Ref_proj1_2_3 = CreateRefFromObj(mPart, proj1_2_3)
        Dim Ref_proj1_3_4 = CreateRefFromObj(mPart, proj1_3_4)
        Dim Ref_proj1_4_1 = CreateRefFromObj(mPart, proj1_4_1)

        Dim Constraints_dim = Standard_body_Sketch3.Constraints
        Dim constraint_dx = Constraints_dim.AddMonoEltCst(CatConstraintType.catCstTypeLength, Ref_proj1_1_2)
        Dim constraint_dy = Constraints_dim.AddMonoEltCst(CatConstraintType.catCstTypeLength, Ref_proj1_2_3)
        Dim Length_dx = constraint_dx.Dimension
        Dim dx_value = Length_dx.Value
        Dim Length_dy = constraint_dy.Dimension
        Dim dy_value = Length_dy.Value

        Standard_body_Sketch3.CloseEdition()
        mPart.Update()


        'Create Formula
        Dim string_1 As String = "Offset_Bbox_Max_X." & j
        Dim string_2 As String = "Offset_Bbox_Min_X." & j
        Dim string_3 As String = "Offset_Bbox_Max_Y." & j
        Dim string_4 As String = "Offset_Bbox_Min_Y." & j
        Dim string_5 As String = "Offset_Bbox_Max_Z." & j
        Dim string_6 As String = "Offset_Bbox_Min_Z." & j
        Dim string_7 As String = "Bbox_dx." & j
        Dim string_8 As String = "Bbox_dy." & j
        Dim string_9 As String = "Bbox_dz." & j



        Dim Offset_Bbox_Max_X As RealParam = mPart.Parameters.CreateDimension(string_1, "Length", 0)
        Dim Offset_Bbox_Min_X As RealParam = mPart.Parameters.CreateDimension(string_2, "Length", 0)
        Dim Offset_Bbox_Max_Y As RealParam = mPart.Parameters.CreateDimension(string_3, "Length", 0)
        Dim Offset_Bbox_Min_Y As RealParam = mPart.Parameters.CreateDimension(string_4, "Length", 0)
        Dim Offset_Bbox_Max_Z As RealParam = mPart.Parameters.CreateDimension(string_5, "Length", 0)
        Dim Offset_Bbox_Min_Z As RealParam = mPart.Parameters.CreateDimension(string_6, "Length", 0)

        Dim Bbox_dx_hidden = mPart.Parameters.CreateDimension(string_7, "Length", dx_value)
        Dim Bbox_dy_hidden = mPart.Parameters.CreateDimension(string_8, "Length", dy_value)
        Dim Bbox_dz_hidden = mPart.Parameters.CreateDimension(string_9, "Length", Length_dz.Value)
        Bbox_dx_hidden.Hidden = True
        Bbox_dy_hidden.Hidden = True
        Bbox_dz_hidden.Hidden = True

        Dim Bbox_dx = mPart.Parameters.CreateDimension(string_7, "Length", 0)
        Dim Bbox_dy = mPart.Parameters.CreateDimension(string_8, "Length", 0)
        Dim Bbox_dz = mPart.Parameters.CreateDimension(string_9, "Length", 0)

        Dim Formula_1 As Formula = mPart.Relations.CreateFormula("formula_Bbox_1." & j, "", Length14, string_2)
        Dim Formula_2 As Formula = mPart.Relations.CreateFormula("formula_Bbox_2." & j, "", Length12, string_1)
        Dim Formula_3 As Formula = mPart.Relations.CreateFormula("formula_Bbox_3." & j, "", Length11, string_4)
        Dim Formula_4 As Formula = mPart.Relations.CreateFormula("formula_Bbox_4." & j, "", Length13, string_3)
        Dim Formula_5 As Formula = mPart.Relations.CreateFormula("formula_Bbox_5." & j, "", oEnd, string_6)
        Dim Formula_6 As Formula = mPart.Relations.CreateFormula("formula_Bbox_6." & j, "", oStart, string_5)

        Dim Formula_7 As Formula = mPart.Relations.CreateFormula("formula_Bbox_7." & j, "", Bbox_dx, "Bbox_dx." & j & "+Offset_Bbox_Min_X." & j & "+Offset_Bbox_Max_X." & j & "-0mm")
        Dim Formula_8 As Formula = mPart.Relations.CreateFormula("formula_Bbox_8." & j, "", Bbox_dy, "Bbox_dy." & j & "+Offset_Bbox_Min_Y." & j & "+Offset_Bbox_Max_Y." & j & "-0mm")
        Dim Formula_9 As Formula = mPart.Relations.CreateFormula("formula_Bbox_9." & j, "", Bbox_dz, "Bbox_dz." & j & "+Offset_Bbox_Min_Z." & j & "+Offset_Bbox_Max_Z." & j & "-0mm")

        mPart.Update()


        'Creation Sweep
        Dim SweepRef_1 = CreateRefFromObj(mPart, Standard_body_Sketch3)
        Dim Guide1 = CreateRefFromObj(mPart, Line_guide)
        Dim Sweep1 As HybridShapeSweepExplicit = mHybridShapeFactory.AddNewSweepExplicit(SweepRef_1, Guide1)
        HybridBody2.AppendHybridShape(Sweep1)
        mPart.Update()

        '將建立的面Join

        Dim mShapeFactory As ShapeFactory = mPart.ShapeFactory
        Dim HybridShapes1 As HybridShapes = HybridBody1.HybridShapes

        Dim Refer1 As Reference = CreateRefFromObj(mPart, Sweep1)

        Dim CloseSurface1 As CloseSurface = mShapeFactory.AddNewCloseSurface(Refer1)
        CloseSurface1.Name = "BoundingBox"
        mPart.Update()
        With PartDoc.Selection
            .Clear()
            .Add(Body1)
            .VisProperties.SetRealColor(255, 255, 128, 0)
            .VisProperties.SetRealOpacity(150, 1)
            .VisProperties.SetRealWidth(4, 1)
        End With

        PartDoc.Selection.Clear()
        mPart.Update()
        'Else msgbox("The active document must be a CATPart")
        'End If


    End Function
    Public Shared Sub HideObject(ByRef PartDoc As PartDocument, ByRef GeometricalSetName As Object)                        '隱藏Part物件
        Dim visPropertySet1 As VisPropertySet
        Dim mSelection As Selection = PartDoc.Selection
        mSelection.Clear()
        PartDoc.Selection.Search("Name:*" & GeometricalSetName & "*,All")
        visPropertySet1 = mSelection.VisProperties
        visPropertySet1.SetShow(1)
        mSelection.Clear()
    End Sub
    Public Shared Sub HideObjectFromProduct(ByRef ProDoc As ProductStructureTypeLib.ProductDocument, ByRef GeometricalSetName As Object)                        '隱藏Product物件
        Dim visPropertySet1 As VisPropertySet
        Dim mSelection As Selection = ProDoc.Selection
        mSelection.Clear()
        ProDoc.Selection.Search("Name:*" & GeometricalSetName & "*,All")
        visPropertySet1 = mSelection.VisProperties
        visPropertySet1.SetShow(1)
        mSelection.Clear()
    End Sub
    Public Shared Sub ShowObjectFromProduct(ByRef ProDoc As ProductStructureTypeLib.ProductDocument, ByRef GeometricalSetName As Object)                        '顯示Product物件
        Dim visPropertySet1 As VisPropertySet
        Dim mSelection As Selection = ProDoc.Selection
        mSelection.Clear()
        ProDoc.Selection.Search("Name:*" & GeometricalSetName & "*,All")
        visPropertySet1 = mSelection.VisProperties
        visPropertySet1.SetShow(0)
        mSelection.Clear()
    End Sub
    Public Shared Function CreateRefFromObj(ByRef mPart As Part, ByRef mObject As Object) As Reference                     '抓取Reference
        Dim oObject As Reference
        Try
            oObject = mPart.CreateReferenceFromObject(mObject)
            Return oObject
        Catch ex As Exception
            MsgBox("Can't Get Reference!" & vbCrLf & ex.Message)
            'Throw New SystemException("Can't Get Reference!")
            Return Nothing
        Finally

        End Try
    End Function

    Public Shared Function c_HybridShapeExtremum(ByRef mHybridShapeFactory As HybridShapeFactory, hybridShapeDX As HybridShapeDirection, Reference1 As Reference, ByVal D_MinMax As Integer)              '建立極端點
        Dim HybridShapeExtremum1 As HybridShapeExtremum = mHybridShapeFactory.AddNewExtremum(Reference1, hybridShapeDX, D_MinMax)
        Return HybridShapeExtremum1
    End Function

    Public Shared Function c_HybridShapeLinePtDir(mHybridShapeFactory As HybridShapeFactory, Originpoint As HybridShapePointCoord, hybridShapeDX As HybridShapeDirection)          '建立線(點和方向)
        Dim Plane_line_1 As HybridShapeLinePtDir = mHybridShapeFactory.AddNewLinePtDir(Originpoint, hybridShapeDX, 0, 0, False)
        Return Plane_line_1
    End Function
    Public Shared Function OpenToNewWindow(ByRef CatiaFactory As CATIA_Property, ByRef mProduct As Product) '打開子件檔案
        Dim StiEngine As CATSmarTeamInteg.StiEngine               'Smart Team Libery
        StiEngine = CatiaFactory.myCATIA.GetItem("CAIEngine")
        Dim StiDBItem As CATSmarTeamInteg.StiDBItem = StiEngine.GetStiDBItemFromAnyObject(mProduct.ReferenceProduct.Parent)
        Dim FileFullName As String = StiDBItem.GetDocumentFullPath
        CatiaFactory.Documents.Open(FileFullName)

    End Function
End Class