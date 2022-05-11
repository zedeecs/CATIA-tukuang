Dim CATIA As Object
Dim DrwDocument   As DrawingDocument
Dim DrwSheets     As DrawingSheets
Dim DrwSheet      As DrawingSheet
Dim DrwViews      As DrawingViews  'make background view active
Dim DrwView       As DrawingView
Dim DrwTexts      As DrawingTexts
Dim Text          As DrawingText
Dim Fact          As Factory2D
Dim Point         As Point2D
Dim Line          As Line2D
Dim Cicle         As Circle2D
Dim Selection     As Selection
Dim GeomElems     As GeometricElements

Sub CATMain()

    MainForms.Show 0


End Sub


Sub FTB_Init()     '��ʼ��

    Set CATIA = GetObject(, "CATIA.Application")
    Set DrwDocument = CATIA.ActiveDocument
    Set DrwSheets = DrwDocument.Sheets
    Set DrwSheet = DrwSheets.ActiveSheet
    Set DrwViews = DrwSheet.Views

End Sub


' ���ڱ�����ͼ��sheet.views.Item(i)��1 = mainview��������ͼ��, 2 = backgroundView(������ͼ)��3 = �����ĵ�һ�������ͼ
' CATIA.StartCommand ("ͼֽ����") �Ͳ������������������л���ͼ��
Sub EnterBackViews()
    Call FTB_Init
    DrwViews.Item(2).Activate
End Sub

Sub ExitBackViews()
    Call FTB_Init
    DrwViews.Item(1).Activate
End Sub

Function Col(idx As Integer) As Variant
    Col = Array(-180, -170, -160, -139, -129, -112, -98, -90, -82, -78, -76, -70, -64, -58, -43, 0, -52, -37, 0, 44, 52)(idx - 1)
    '           1      2     4     4     5     6     7    8    9   10   11   12   13   14   15      17   18     20  21
End Function

Function Row(idx As Integer) As Variant
    Row = Array(7, 14, 21, 28, 35, 42, 45, 49, 56, 0, -7, -14, -21, -28, 0, -13)(idx - 1)
    '          1   2   3   4   5   6   7   8   9     11   12   13   14      16
End Function

Function GetMacroID() As String
    GetMacroID = "Drawing_Titleblock_Sample1"
End Function


Function GetRevRowHeight() As Double
    GetRevRowHeight = 10#
End Function

Function GetDisplayFormat() As String
    GetDisplayFormat = Array("Letter", "Legal", "A0", "A1", "A2", "A3", "A4", "A", "B", "C", "D", "E", "F", "User")(Sheet.PaperSize)
End Function

Function GetOffset(idx As Integer) As Variant
    If Sheet.PaperSize = catPaperA0 Or Sheet.PaperSize = catPaperA1 Or (Sheet.PaperSize = catPaperUser And (GetWidth() > 594# Or GetHeight() > 594#)) Then
        GetOffset = Array(10, 25)(idx - 0)
    Else
        GetOffset = Array(5, 25)(idx - 0)
    End If
End Function

' ��ȡͼֽ���
Function GetWidth() As Double
    Select Case TypeName(Sheet)
     Case "DrawingSheet": GetWidth = Sheet.GetPaperWidth
     Case "Layout2DSheet": GetWidth = Sheet.PaperWidth
    End Select
End Function

' ��ȡͼֽ�߶�
Function GetHeight() As Double
    Select Case TypeName(Sheet)
     Case "DrawingSheet": GetHeight = Sheet.GetPaperHeight
     Case "Layout2DSheet": GetHeight = Sheet.PaperHeight
    End Select
End Function


' ��ȡͼ��ԭ��X�����½ǣ�
Function GetOH() As Double
    GetOH = GetWidth() - GetOffset(0)
End Function


' ��ȡͼ��ԭ��Y�����½ǣ�
Function GetOV() As Double
    GetOV = GetOffset(0)
End Function

Function GetRevLetter(index As Integer)
    GetRevLetter = Chr(Asc("A") + index - 1)
End Function

Function CreateLine(iX1 As Double, iY1 As Double, iX2 As Double, iY2 As Double, iName As String) As Curve2D
    '-------------------------------------------------------------------------------
    ' Creates a sketcher lines thanks to the current 2D factory set to the global variable Fact
    ' The created line is reneamed to the given iName
    ' Start point  and End point are created and renamed iName&"_start", iName&"_end"
    ' ���崴���߶εĺ����������������ʼ��������꣬�߶����ƣ�
    '-------------------------------------------------------------------------------
    Set CreateLine = Fact.CreateLine(iX1, iY1, iX2, iY2)
    CreateLine.Name = iName
    Set Point = CreateLine.StartPoint 'Create the start point
    Point.Name = iName & "_start"
    Set Point = CreateLine.EndPoint 'Create the start point
    Point.Name = iName & "_end"
End Function

Function CreateText(iCaption As String, iX As Double, iY As Double, iName As String) As DrawingText
    '-------------------------------------------------------------------------------
    'How to create a text
    '-------------------------------------------------------------------------------
    Set CreateText = Texts.Add(iCaption, iX, iY)
    CreateText.Name = iName
    CreateText.AnchorPosition = catMiddleCenter
End Function

Function CreateTextAF(iCaption As String, iX As Double, iY As Double, iName As String, iAnchorPosition As CatTextAnchorPosition, iFontSize As Double) As DrawingText
    '-------------------------------------------------------------------------------
    'How to create a text
    ' example: CreateTextAF Text_01,GetOH() + Col(1) + 11.       ,GetOV() + .5*Row(1)  ,"TitleBlock_Text_Rights" , catMiddleLeft   ,1.5
    ' ����������壺��������, X�������, Y����+0.5�и�, ����Ԫ�ص�����, ê��, �ָ�
    '-------------------------------------------------------------------------------
    Set CreateTextAF = Texts.Add(iCaption, iX, iY)
    CreateTextAF.Name = iName
    CreateTextAF.AnchorPosition = iAnchorPosition
    CreateTextAF.SetFontSize 0, 0, iFontSize
End Function

Sub SelectAll(iQuery As String)
    Selection.Clear
    Selection.Add (View)
    'MsgBox iQuery
    Selection.Search iQuery & ",sel"
End Sub

Sub DeleteAll(iQuery As String)
    '-------------------------------------------------------------------------------
    'Delete all elements  matching the query string iQuery
    'Pay attention no to provide a localized query string.
    '-------------------------------------------------------------------------------
    Selection.Clear
    Selection.Add (View)
    'MsgBox iQuery
    Selection.Search iQuery & ",sel"
    ' Avoid Delete failure in case of an empty query result
    If Selection.Count2 <> 0 Then Selection.Delete
End Sub


'Sub CATDrw_Creation(targetSheet As CATIABase)
Sub CATDrw_Creation()
    '-------------------------------------------------------------------------------
    'How to create the FTB
    '-------------------------------------------------------------------------------
    'If Not CATInit(targetSheet) Then Exit Sub
        If CATCheckRef(1) Then Exit Sub 'To check whether a FTB exists already in the sheet
            CATCreateReference          'To place on the drawing a reference point
            CATFrame      'To draw the frame
            CATCreateTitleBlockFrame    'To draw the geometry
            ' CATCreateTitleBlockStandard 'To draw the standard representation ��˾ûҪ��������ţ�ע�͵�
            CATTitleBlockText     'To fill in the title block ���ֿ�
            CATColorGeometry 'To change the geometry color
            CATExit targetSheet      'To save the sketch edition
End Sub

Sub CATDrw_Deletion()
    '-------------------------------------------------------------------------------
    'How to delete the FTB
    '-------------------------------------------------------------------------------
    If Not CATInit(targetSheet) Then Exit Sub
        If CATCheckRef(0) Then Exit Sub
            DeleteAll "..Name=Frame_*"
            DeleteAll "..Name=TitleBlock_*"
            DeleteAll "..Name=RevisionBlock_*"
            DeleteAll "..Name=Reference_*"
            DeleteAll "..Name=SecTitleBlock_*"
            DeleteAll "..Name=PartNumberBlock_*"
            CATExit targetSheet
End Sub

Sub CATDrw_Resizing()
    '-------------------------------------------------------------------------------
    'How to resize the FTB
    '-------------------------------------------------------------------------------
    If Not CATInit(targetSheet) Then Exit Sub
        If CATCheckRef(0) Then Exit Sub
            Dim TbTranslation(2)
            ComputeTitleBlockTranslation TbTranslation
            Dim RbTranslation(2)
            ComputeRevisionBlockTranslation RbTranslation
            If TbTranslation(0) <> 0 Or TbTranslation(1) <> 0 Then
                ' Redraw Sheet Frame
                DeleteAll "CATDrwSearch.DrwText.Name=Frame_Text_*"
                DeleteAll "CATDrwSearch.2DGeometry.Name=Frame_*"
                CATFrame
                ' Redraw Standard Pictorgram
                CATDeleteTitleBlockStandard
                CATCreateTitleBlockStandard
                ' Redraw Title Block Frame
                CATDeleteTitleBlockFrame
                CATCreateTitleBlockFrame
                CATMoveTitleBlockText TbTranslation
                ' Redraw revision block
                CATDeleteRevisionBlockFrame
                CATCreateRevisionBlockFrame
                CATMoveRevisionBlockText RbTranslation

                ' Move the views
                CATColorGeometry
                CATMoveViews TbTranslation
                CATLinks

            End If
            CATExit targetSheet
End Sub

Sub CATDrw_Update()
    '-------------------------------------------------------------------------------
    'How to update the FTB
    '-------------------------------------------------------------------------------
    If Not CATInit(targetSheet) Then Exit Sub
        If CATCheckRef(0) Then Exit Sub
            CATDeleteTitleBlockStandard
            ' CATCreateTitleBlockStandard
            CATLinks
            CATColorGeometry
            CATExit targetSheet

End Sub

Function GetContext()
    ' Find execution context
    Select Case TypeName(Sheet)
     Case "DrawingSheet"
        Select Case TypeName(ActiveDoc)
         Case "DrawingDocument": GetContext = "DRW"
         Case "ProductDocument": GetContext = "SCH"
         Case Else: GetContext = "Unexpected"
        End Select

     Case "Layout2DSheet": GetContext = "LAY"
     Case Else: GetContext = "Unexpected"
    End Select
End Function

Sub CATDrw_CheckedBy()
    '-------------------------------------------------------------------------------
    'How to update a bit more the FTB
    '-------------------------------------------------------------------------------
    If Not CATInit(targetSheet) Then Exit Sub
        If CATCheckRef(0) Then Exit Sub
            CATFillField "TitleBlock_Text_Controller_1", "TitleBlock_Text_CDate_1", "checked"
            CATExit targetSheet
End Sub

Sub CATDrw_AddRevisionBlock()
    '-------------------------------------------------------------------------------
    'How to create or modify a revison block
    '-------------------------------------------------------------------------------
    If Not CATInit(targetSheet) Then Exit Sub
        If CATCheckRef(0) Then Exit Sub

            CATAddRevisionBlockText 'To fill in the title block
            CATDeleteRevisionBlockFrame
            CATCreateRevisionBlockFrame 'To draw the geometry

            CATColorGeometry
            CATExit targetSheet
End Sub

Function CATInit()
    '-------------------------------------------------------------------------------
    'How to init the dialog and create main objects
    '-------------------------------------------------------------------------------
    Set Selection = CATIA.ActiveDocument.Selection
    Set Sheet = targetSheet
    Set Sheets = Sheet.Parent
    Set ActiveDoc = Sheets.Parent
    Set Views = Sheet.Views
    Set View = Views.Item(2)        'Get the background view
    Set Texts = View.Texts
    Set Fact = View.Factory2D
    Set GeomElems = View.GeometricElements

    If GetContext() = "Unexpected" Then
        Msg = "The macro runs in an inappropriate environment." & Chr(13) & "The script will terminate wihtout finishing the current action."
        Title = "Unexpected environement error"
        MsgBox Msg, 16, Title
        CATInit = False 'Exit with error
        Exit Function
    End If

    Selection.Clear
    CATIA.HSOSynchronized = False

    CATInit = True 'Exit without error
End Function

Sub CATExit()
    '-------------------------------------------------------------------------------
    'How to restore the document working mode
    '-------------------------------------------------------------------------------
    Selection.Clear
    CATIA.HSOSynchronized = True

    View.SaveEdition
End Sub


Sub CATCreateReference()
    '-------------------------------------------------------------------------------
    'How to create a reference text
    ' ��ͼֽ�����½Ǵ���һ�������֣���Ǵ����������԰�����ʶ��
    '-------------------------------------------------------------------------------
    Set Text = Texts.Add("", GetWidth() - 25, 5)
    ' Set Text = Texts.Add("", GetWidth() - GetOffset(1), GetOffset(1))
    Text.Name = "Reference_" + GetMacroID
End Sub

Function CATCheckRef(Mode As Integer) As Integer
    '-------------------------------------------------------------------------------
    'How to check that the called macro is the right one
    '-------------------------------------------------------------------------------
    nbTexts = Texts.Count
    i = 0
    notFound = 0
    While (notFound = 0 And i < nbTexts)
        i = i + 1
        Set Text = Texts.Item(i)
        WholeName = Text.Name
        leftText = Left(WholeName, 10)
        If (leftText = "Reference_") Then
            notFound = 1
            refText = "Reference_" + GetMacroID()
            If (Mode = 1) Then
                MsgBox "Frame and Titleblock already created!"
                CATCheckRef = 1
                Exit Function
            ElseIf (Text.Name <> refText) Then
                MsgBox "Frame and Titleblock created using another style:" + Chr(10) + "        " + GetMacroID()
                CATCheckRef = 1
                Exit Function
            Else
                CATCheckRef = 0
                Exit Function
            End If
        End If
    Wend

    If Mode = 1 Then
        CATCheckRef = 0
    Else
        MsgBox "No Frame and Titleblock!"
        CATCheckRef = 1
    End If

End Function

Function CATCheckRev() As Integer
    '-------------------------------------------------------------------------------
    'How to check that a revision block alredy exists
    '-------------------------------------------------------------------------------
    SelectAll "CATDrwSearch.DrwText.Name=RevisionBlock_Text_Rev_*"
    CATCheckRev = Selection.Count2
End Function

Sub CATFrame()
    '-------------------------------------------------------------------------------
    'How to create the Frame
    '-------------------------------------------------------------------------------
    Dim Cst_1   As Double  'Length (in cm) between 2 horinzontal marks
    Dim Cst_2   As Double  'Length (in cm) between 2 vertical marks
    Dim Nb_CM_H As Integer 'Number/2 of horizontal centring marks
    Dim Nb_CM_V As Integer 'Number/2 of vertical centring marks
    Dim Ruler   As Integer 'Ruler length (in cm)

    ' CATFrameStandard     Nb_CM_H, Nb_CM_V, Ruler, Cst_1, Cst_2  ' ������ͼ�������Ǻͱ������
    CATFrameBorder
    ' CATFrameCentringMark Nb_CM_H, Nb_CM_V, Ruler, Cst_1, Cst_2  ' �ر� Centring Marks
    ' CATFrameText         Nb_CM_H, Nb_CM_V, Ruler, Cst_1, Cst_2  ' �ر� Frame Text
    ' CATFrameRuler        Ruler, Cst_1  ' �ر� Frame Ruler
End Sub

Sub CATFrameStandard(Nb_CM_H As Integer, Nb_CM_V As Integer, Ruler As Integer, Cst_1 As Double, Cst_2 As Double)
    '-------------------------------------------------------------------------------
    'How to compute standard values
    '-------------------------------------------------------------------------------
    Cst_1 = 74.2 '297, 594, 1189 are multiples of 74.2
    Cst_2 = 52.5 '210, 420, 841  are multiples of 52.2

    If Sheet.Orientation = catPaperPortrait And _
        (Sheet.PaperSize = catPaperA0 Or _
        Sheet.PaperSize = catPaperA2 Or _
        Sheet.PaperSize = catPaperA4) Or _
        Sheet.Orientation = catPaperLandscape And _
        (Sheet.PaperSize = catPaperA1 Or _
        Sheet.PaperSize = catPaperA3) Then
        Cst_1 = 52.5
        Cst_2 = 74.2
    End If

    Nb_CM_H = CInt(0.5 * GetWidth() / Cst_1)
    Nb_CM_V = CInt(0.5 * GetHeight() / Cst_2)
    Ruler = CInt((Nb_CM_H - 1) * Cst_1 / 50) * 100   'here is computed the maximum ruler length
    If GetRulerLength() < Ruler Then Ruler = GetRulerLength()
End Sub

Sub CATFrameBorder()
    '-------------------------------------------------------------------------------
    'How to draw the frame border
    ' ��ƫ��25������3��5����10
    '-------------------------------------------------------------------------------
    On Error Resume Next
    CreateLine GetOH(), GetOV(), GetOffset(1), GetOV(), "Frame_Border_Bottom"
    CreateLine GetOffset(1), GetOV(), GetOffset(1), GetHeight() - GetOffset(0), "Frame_Border_Left"
    CreateLine GetOffset(1), GetHeight() - GetOffset(0), GetOH(), GetHeight() - GetOffset(0), "Frame_Border_Top"
    CreateLine GetOH(), GetHeight() - GetOffset(0), GetOH(), GetOV(), "Frame_Border_Right"
    If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
End Sub

Sub CATFrameCentringMark(Nb_CM_H As Integer, Nb_CM_V As Integer, Ruler As Integer, Cst_1 As Double, Cst_2 As Double)
    '-------------------------------------------------------------------------------
    'How to draw the centring marks
    '-------------------------------------------------------------------------------
    On Error Resume Next
    CreateLine 0.5 * GetWidth(), GetHeight() - GetOffset(), 0.5 * GetWidth(), GetHeight(), "Frame_CentringMark_Top"
    CreateLine 0.5 * GetWidth(), GetOV(), 0.5 * GetWidth(), 0#, "Frame_CentringMark_Bottom"
    CreateLine GetOV(), 0.5 * GetHeight(), 0#, 0.5 * GetHeight(), "Frame_CentringMark_Left"
    CreateLine GetWidth() - GetOffset(), 0.5 * GetHeight(), GetWidth(), 0.5 * GetHeight(), "Frame_CentringMark_Right"
    For i = Nb_CM_H To Ruler / 2 / Cst_1 Step -1
        If (i * Cst_1 < 0.5 * GetWidth() - 1#) Then
            X = 0.5 * GetWidth() + i * Cst_1
            CreateLine X, GetOV(), X, 0.25 * GetOffset(), "Frame_CentringMark_Bottom_" & Int(X)
            X = 0.5 * GetWidth() - i * Cst_1
            CreateLine X, GetOV(), X, 0.25 * GetOffset(), "Frame_CentringMark_Bottom_" & Int(X)
        End If
        Next
        For i = 1 To Nb_CM_H
            If (i * Cst_1 < 0.5 * GetWidth() - 1#) Then
                X = 0.5 * GetWidth() + i * Cst_1
                CreateLine X, GetHeight() - GetOffset(), X, GetHeight() - 0.25 * GetOffset(), "Frame_CentringMark_Top_" & Int(X)
                X = 0.5 * GetWidth() - i * Cst_1
                CreateLine X, GetHeight() - GetOffset(), X, GetHeight() - 0.25 * GetOffset(), "Frame_CentringMark_Top_" & Int(X)
            End If
            Next

            For i = 1 To Nb_CM_V
                If (i * Cst_2 < 0.5 * GetHeight() - 1#) Then
                    Y = 0.5 * GetHeight() + i * Cst_2
                    CreateLine GetOV(), Y, 0.25 * GetOffset(), Y, "Frame_CentringMark_Left_" & Int(Y)
                    CreateLine GetOH(), Y, GetWidth() - 0.25 * GetOffset(), Y, "Frame_CentringMark_Right_" & Int(Y)
                    Y = 0.5 * GetHeight() - i * Cst_2
                    CreateLine GetOV(), Y, 0.25 * GetOffset(), Y, "Frame_CentringMark_Left_" & Int(Y)
                    CreateLine GetOH(), Y, GetWidth() - 0.25 * GetOffset(), Y, "Frame_CentringMark_Right_" & Int(Y)
                End If
                Next
                If Err.Number <> 0 Then Err.Clear
                    On Error GoTo 0
End Sub

Sub CATFrameText(Nb_CM_H As Integer, Nb_CM_V As Integer, Ruler As Integer, Cst_1 As Double, Cst_2 As Double)
    '-------------------------------------------------------------------------------
    'How to create coordinates
    '-------------------------------------------------------------------------------
    On Error Resume Next

    For i = Nb_CM_H To (Ruler / 2 / Cst_1 + 1) Step -1
        CreateText Chr(65 + Nb_CM_H - i), 0.5 * GetWidth() + (i - 0.5) * Cst_1, 0.5 * GetOffset(), "Frame_Text_Bottom_1_" & Chr(65 + Nb_CM_H - i)
        CreateText Chr(64 + Nb_CM_H + i), 0.5 * GetWidth() - (i - 0.5) * Cst_1, 0.5 * GetOffset(), "Frame_Text_Bottom_2_" & Chr(65 + Nb_CM_H + i)
        Next

        For i = 1 To Nb_CM_H
            t = Chr(65 + Nb_CM_H - i)
            CreateText(t, 0.5 * GetWidth() + (i - 0.5) * Cst_1, GetHeight() - 0.5 * GetOffset(), "Frame_Text_Top_1_" & t).Angle = -90
            t = Chr(64 + Nb_CM_H + i)
            CreateText(t, 0.5 * GetWidth() - (i - 0.5) * Cst_1, GetHeight() - 0.5 * GetOffset(), "Frame_Text_Top_2_" & t).Angle = -90
            Next

            For i = 1 To Nb_CM_V
                t = CStr(Nb_CM_V + i)
                CreateText t, GetWidth() - 0.5 * GetOffset(), 0.5 * GetHeight() + (i - 0.5) * Cst_2, "Frame_Text_Right_1_" & t
                CreateText(t, 0.5 * GetOffset(), 0.5 * GetHeight() + (i - 0.5) * Cst_2, "Frame_Text_Left_1_" & t).Angle = -90

                t = CStr(Nb_CM_V - i + 1)
                CreateText t, GetWidth() - 0.5 * GetOffset(), 0.5 * GetHeight() - (i - 0.5) * Cst_2, "Frame_Text_Right_1_" & t
                CreateText(t, 0.5 * GetOffset(), 0.5 * GetHeight() - (i - 0.5) * Cst_2, "Frame_Text_Left_2" & t).Angle = -90
                Next

                If Err.Number <> 0 Then Err.Clear
                    On Error GoTo 0
End Sub

Sub CATFrameRuler(Ruler As Integer, Cst_1 As Single)
    '-------------------------------------------------------------------------------
    'How to create a ruler
    '-------------------------------------------------------------------------------
    'Frame_Ruler_Guide -----------------------------------------------
    'Frame_Ruler_1cm   | | | | | | | | | | | | | | | | | | | | | | | |
    'Frame_Ruler_5cm   |         |         |         |         |

    On Error Resume Next
    CreateLine 0.5 * GetWidth() - Ruler / 2, 0.75 * GetOffset(), 0.5 * GetWidth() + Ruler / 2, 0.75 * GetOffset(), "Frame_Ruler_Guide"

    For i = 1 To Ruler / 100
        CreateLine 0.5 * GetWidth() - 50 * i, GetOV(), 0.5 * GetWidth() - 50 * i, 0.5 * GetOffset(), "Frame_Ruler_1_" & i
        CreateLine 0.5 * GetWidth() + 50 * i, GetOV(), 0.5 * GetWidth() + 50 * i, 0.5 * GetOffset(), "Frame_Ruler_2_" & i
        For j = 1 To 4
            CreateLine .5 * GetWidth() - 50 * i + 10 * j,  GetOV(),  .5 * GetWidth() - 50 * i + 10 * j,  .75 * GetOffset(), "Frame_Ruler_3"&i&"_"&j
            CreateLine .5 * GetWidth() + 50 * i - 10 * j,  GetOV(),  .5 * GetWidth() + 50 * i - 10 * j,  .75 * GetOffset(), "Frame_Ruler_4"&i&"_"&j
            Next
            Next

            If Err.Number <> 0 Then Err.Clear
                On Error GoTo 0
End Sub
Sub CATDeleteTitleBlockFrame()
    DeleteAll "CATDrwSearch.2DGeometry.Name=TitleBlock_Line_*"
End Sub
Sub CATCreateTitleBlockFrame()
    '-------------------------------------------------------------------------------
    'How to draw the title block geometry
    '-------------------------------------------------------------------------------
    CreateLine GetOH(), GetOV(), GetOH() + Col(1), GetOV(), "TitleBlock_Line_Bottom"
    CreateLine GetOH() + Col(1), GetOV(), GetOH() + Col(1), GetOV() + Row(9), "TitleBlock_Line_Left"
    CreateLine GetOH() + Col(1), GetOV() + Row(9), GetOH(), GetOV() + Row(9), "TitleBlock_Line_Top"
    CreateLine GetOH(), GetOV() + Row(9), GetOH(), GetOV(), "TitleBlock_Line_Right"
    CreateLine GetOH() + Col(1), GetOV() + Row(1), GetOH(), GetOV() + Row(1), "TitleBlock_Line_Row_1"
    CreateLine GetOH() + Col(1), GetOV() + Row(2), GetOH() + Col(14), GetOV() + Row(2), "TitleBlock_Line_Row_2"
    CreateLine GetOH() + Col(1), GetOV() + Row(3), GetOH(), GetOV() + Row(3), "TitleBlock_Line_Row_3"
    CreateLine GetOH() + Col(1), GetOV() + Row(4), GetOH() + Col(7), GetOV() + Row(4), "TitleBlock_Line_Row_4"
    CreateLine GetOH() + Col(1), GetOV() + Row(5), GetOH(), GetOV() + Row(5), "TitleBlock_Line_Row_5"
    CreateLine GetOH() + Col(1), GetOV() + Row(6), GetOH() + Col(7), GetOV() + Row(6), "TitleBlock_Line_Row_6"
    CreateLine GetOH() + Col(14), GetOV() + Row(7), GetOH(), GetOV() + Row(7), "TitleBlock_Line_Row_7"
    CreateLine GetOH() + Col(1), GetOV() + Row(8), GetOH() + Col(7), GetOV() + Row(8), "TitleBlock_Line_Row_8"
    ' For i = 1 To GetNbOfRevision()-1
    '   CreateLine GetOH() + Col(5),  GetOV()+Row(5)/GetNbOfRevision()*i,  GetOH(),  GetOV()+Row(5)/GetNbOfRevision()*i, "TitleBlock_Line_Row_5"&i
    ' Next
    CreateLine GetOH() + Col(2), GetOV(), GetOH() + Col(2), GetOV() + Row(9), "TitleBlock_Line_Column_1"
    CreateLine GetOH() + Col(3), GetOV() + Row(5), GetOH() + Col(3), GetOV() + Row(9), "TitleBlock_Line_Column_2"
    CreateLine GetOH() + Col(4), GetOV(), GetOH() + Col(4), GetOV() + Row(5), "TitleBlock_Line_Column_3"
    CreateLine GetOH() + Col(5), GetOV(), GetOH() + Col(5), GetOV() + Row(9), "TitleBlock_Line_Column_4"
    CreateLine GetOH() + Col(6), GetOV() + Row(5), GetOH() + Col(6), GetOV() + Row(9), "TitleBlock_Line_Column_6"
    CreateLine GetOH() + Col(7), GetOV(), GetOH() + Col(7), GetOV() + Row(9), "TitleBlock_Line_Column_7"
    CreateLine GetOH() + Col(8), GetOV() + Row(2), GetOH() + Col(8), GetOV() + Row(9), "TitleBlock_Line_Column_8"
    CreateLine GetOH() + Col(9), GetOV() + Row(1), GetOH() + Col(9), GetOV() + Row(2), "TitleBlock_Line_Column_9"
    CreateLine GetOH() + Col(10), GetOV() + Row(3), GetOH() + Col(10), GetOV() + Row(5), "TitleBlock_Line_Column_10"
    CreateLine GetOH() + Col(11), GetOV() + Row(1), GetOH() + Col(11), GetOV() + Row(2), "TitleBlock_Line_Column_11"
    CreateLine GetOH() + Col(12), GetOV() + Row(3), GetOH() + Col(12), GetOV() + Row(5), "TitleBlock_Line_Column_12"
    CreateLine GetOH() + Col(12), GetOV() + Row(1), GetOH() + Col(12), GetOV() + Row(2), "TitleBlock_Line_Column_13"
    CreateLine GetOH() + Col(13), GetOV() + Row(1), GetOH() + Col(13), GetOV() + Row(2), "TitleBlock_Line_Column_14"
    CreateLine GetOH() + Col(14), GetOV(), GetOH() + Col(14), GetOV() + Row(9), "TitleBlock_Line_Column_15"
    CreateLine GetOH() + Col(15), GetOV(), GetOH() + Col(15), GetOV() + Row(9), "TitleBlock_Line_Column_16"

    ' ���� ���ϽǸ������
    CreateLine GetOH(), GetHeight() - GetOffset(0) + Row(14), GetOH() + Col(17), GetHeight() - GetOffset(0) + Row(14), "SecTitleBlock_Line_Bottom"
    CreateLine GetOH() + Col(17), GetHeight() - GetOffset(0) + Row(14), GetOH() + Col(17), GetHeight() - GetOffset(0), "SecTitleBlock_Line_Left"
    CreateLine GetOH() + Col(18), GetHeight() - GetOffset(0) + Row(14), GetOH() + Col(18), GetHeight() - GetOffset(0), "SecTitleBlock_Line_Column_1"
    CreateLine GetOH(), GetHeight() - GetOffset(0) + Row(13), GetOH() + Col(17), GetHeight() - GetOffset(0) + Row(13), "SecTitleBlock_Line_Row_3"
    CreateLine GetOH(), GetHeight() - GetOffset(0) + Row(12), GetOH() + Col(17), GetHeight() - GetOffset(0) + Row(12), "SecTitleBlock_Line_Row_2"
    CreateLine GetOH(), GetHeight() - GetOffset(0) + Row(11), GetOH() + Col(17), GetHeight() - GetOffset(0) + Row(11), "SecTitleBlock_Line_Row_1"

    ' ���� ���Ͻ� ��תͼ�ſ�
    CreateLine GetOffset(1), GetHeight() - GetOffset(0) + Row(16), GetOffset(1) + Col(21), GetHeight() - GetOffset(0) + Row(16), "PartNumberBlock_Line_Bottom"
    CreateLine GetOffset(1) + Col(21), GetHeight() - GetOffset(0), GetOffset(1) + Col(21), GetHeight() - GetOffset(0) + Row(16), "PartNumberBlock_Line_Right"
    CreateLine GetOffset(1) + Col(20), GetHeight() - GetOffset(0), GetOffset(1) + Col(20), GetHeight() - GetOffset(0) + Row(16), "PartNumberBlock_Line_Row_1"

End Sub

Sub CATCreateTitleBlockStandard()
    '-------------------------------------------------------------------------------
    'How to create the standard representation
    ' ����ͶӰ����
    '-------------------------------------------------------------------------------
    Dim R1   As Double
    Dim R2   As Double
    Dim X(5) As Double
    Dim Y(7) As Double

    R1 = 2#
    R2 = 4#
    X(1) = GetOH() + Col(2) + 2#
    X(2) = X(1) + 1.5
    X(3) = X(1) + 9.5
    X(4) = X(1) + 15.5
    X(5) = X(1) + 21#
    Y(1) = GetOV() + (Row(2) + Row(3)) / 2#
    Y(2) = Y(1) + R1
    Y(3) = Y(1) + R2
    Y(4) = Y(1) + 5.5
    Y(5) = Y(1) - R1
    Y(6) = Y(1) - R2
    Y(7) = 2 * Y(1) - Y(4)

    If Sheet.ProjectionMethod <> catFirstAngle Then
        Xtmp = X(2)
        X(2) = X(1) + X(5) - X(3)
        X(3) = X(1) + X(5) - Xtmp
        X(4) = X(1) + X(5) - X(4)
    End If

    On Error Resume Next
    CreateLine X(1), Y(1), X(5), Y(1), "TitleBlock_Standard_Line_Axis_1"
    CreateLine X(4), Y(7), X(4), Y(4), "TitleBlock_Standard_Line_Axis_2"
    CreateLine X(2), Y(5), X(2), Y(2), "TitleBlock_Standard_Line_1"
    CreateLine X(2), Y(2), X(3), Y(3), "TitleBlock_Standard_Line_2"
    CreateLine X(3), Y(3), X(3), Y(6), "TitleBlock_Standard_Line_3"
    CreateLine X(3), Y(6), X(2), Y(5), "TitleBlock_Standard_Line_4"
    Set circle = Fact.CreateClosedCircle(X(4), Y(1), R1)
    circle.Name = "TitleBlock_Standard_Circle_1"
    Set circle = Fact.CreateClosedCircle(X(4), Y(1), R2)
    circle.Name = "TitleBlock_Standard_Circle_2"
    If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0

End Sub

Sub CATTitleBlockText()
    '-------------------------------------------------------------------------------
    'How to fill in the title block
    '-------------------------------------------------------------------------------

    Text_01 = "���"
    Text_02 = "����"
    Text_03 = "�����ļ���"
    Text_04 = "ǩ  ��"
    Text_05 = "�� ��"
    Text_06 = "�� ��"
    Text_07 = "�� λ"
    Text_08 = "��(��)��"
    Text_09 = "���"
    Text_10 = "У��"
    Text_11 = "����"
    Text_12 = "����"
    Text_13 = "����װ��ͼ��"
    Text_14 = "���"
    Text_15 = "��׼��"
    Text_16 = "����"
    Text_17 = "�׶α��"
    Text_18 = "�� ��"
    Text_19 = "����"
    Text_20 = "��׼"
    Text_21 = "��    ��  ��    ��"
    Text_23 = "ͼ ��"
    Text_24 = "��  ��"
    Text_25 = "����ӹ�"
    Text_26 = "�ȴ���"
    Text_27 = "���洦��"
    ' Text_15 = CATIA.SystemService.Environ("LOGNAME")
    ' If Text_15 = "" Then Text_15 = CATIA.SystemService.Environ("USERNAME")

    CreateTextAF Text_01, GetOH() + Col(1) + 5#, GetOV() + Row(5) + 3.5, "TitleBlock_Text_Mark", catMiddleCenter, 3
    CreateTextAF Text_02, GetOH() + Col(2) + 5#, GetOV() + Row(5) + 3.5, "TitleBlock_Text_MarkNum", catMiddleCenter, 3
    CreateTextAF Text_03, GetOH() + Col(3) + 15.5, GetOV() + Row(5) + 3.5, "TitleBlock_Text_ModifNum", catMiddleCenter, 3
    Texts.GetItem("TitleBlock_Text_ModifNum").SetParameterOnSubString catCharSpacing, 0, 0, 50
    CreateTextAF Text_04, GetOH() + Col(5) + 8.5, GetOV() + Row(5) + 3.5, "TitleBlock_Text_Sign", catMiddleCenter, 3
    CreateTextAF Text_05, GetOH() + Col(6) + 7#, GetOV() + Row(5) + 3.5, "TitleBlock_Text_Date", catMiddleCenter, 3
    CreateTextAF Text_06, GetOH() + Col(7) + 4#, GetOV() + Row(7), "TitleBlock_Text_Material", catMiddleCenter, 3
    Texts.GetItem("TitleBlock_Text_Material").WrappingWidth = 4
    CreateTextAF Text_06, GetOH() + Col(10) + 4#, GetOV() + Row(7), "TitleBlock_Text_Material_1", catMiddleCenter, 3

    CreateTextAF Text_11, GetOH() + Col(7) + 4#, GetOV() + Row(4), "TitleBlock_Text_Weight", catMiddleCenter, 3      '����
    CreateTextAF Text_12, GetOH() + Col(10) + 4#, GetOV() + Row(4), "TitleBlock_Text_PCS", catMiddleCenter, 3      '����
    CreateTextAF Text_12, GetOH() + Col(12) + 6#, GetOV() + Row(4), "TitleBlock_Text_PCS_1", catMiddleCenter, 3      '����
    CreateTextAF Text_09, GetOH() + Col(1) + 5#, GetOV() + Row(4) + 3.5, "TitleBlock_Text_Designer", catMiddleCenter, 3
    CreateTextAF Text_10, GetOH() + Col(1) + 5#, GetOV() + Row(3) + 3.5, "TitleBlock_Text_Designer_1", catMiddleCenter, 3     'У��
    CreateTextAF Text_15, GetOH() + Col(4) + 5#, GetOV() + Row(2) + 3.5, "TitleBlock_Text_ISO", catMiddleCenter, 3
    Texts.GetItem("TitleBlock_Text_ISO").SetParameterOnSubString catCharSpacing, 0, 0, -25
    CreateTextAF Text_16, GetOH() + Col(7) + 4#, GetOV() + Row(2) + 3.5, "TitleBlock_Text_Scale", catMiddleCenter, 3      '����
    CreateTextAF Text_14, GetOH() + Col(1) + 5#, GetOV() + Row(1) + 3.5, "TitleBlock_Text_Audit", catMiddleCenter, 3      '���
    CreateTextAF Text_17, GetOH() + Col(7) + 8#, GetOV() + Row(1) + 3.5, "TitleBlock_Text_stage", catMiddleCenter, 3      '�׶α��
    Texts.GetItem("TitleBlock_Text_stage").SetParameterOnSubString catCharSpacing, 0, 0, -25
    CreateTextAF Text_19, GetOH() + Col(1) + 5#, GetOV() + 3.5, "TitleBlock_Text_Craft", catMiddleCenter, 3      '����
    CreateTextAF Text_20, GetOH() + Col(4) + 5#, GetOV() + 3.5, "TitleBlock_Text_Approve", catMiddleCenter, 3      '��׼
    CreateTextAF Text_21, GetOH() + Col(7) + 20#, GetOV() + 3.5, "TitleBlock_Text_Page", catMiddleCenter, 3      '������
    CreateTextAF Text_23, GetOH() + Col(14) + 7.5, GetOV() + 3.5, "TitleBlock_Text_PartNum", catMiddleCenter, 3      'ͼ��
    CreateTextAF Text_23, GetOffset(1) + Col(20) + 0.5 * (Col(21) - Col(20)), GetHeight() - GetOffset(0) - 6.5, "TitleBlock_Text_PartNum_1", catMiddleCenter, 3  'ͼ��
    Texts.GetItem("TitleBlock_Text_PartNum_1").TextProperties.Mirror = catTextHorizontalAndVerticalFlip
    Texts.GetItem("TitleBlock_Text_PartNum_1").WrappingWidth = 4
    CreateTextAF Text_18, GetOH() + Col(14) + 7.5, GetOV() + Row(2), "TitleBlock_Text_Name", catMiddleCenter, 3     '����
    CreateTextAF Text_18, GetOH() - 21.5, GetOV() + Row(2), "TitleBlock_Text_Name_1", catMiddleCenter, 3              '����
    CreateTextAF Text_13, GetOH() + Col(14) + 7.5, GetOV() + Row(4), "TitleBlock_Text_asm", catMiddleCenter, 3     'װ������
    Texts.GetItem("TitleBlock_Text_asm").WrappingWidth = 14
    CreateTextAF Text_08, GetOH() + Col(14) + 7.5, GetOV() + Row(5) + 5, "TitleBlock_Text_Type", catMiddleCenter, 3    '�ͺ�
    Texts.GetItem("TitleBlock_Text_Type").SetParameterOnSubString catCharSpacing, 0, 0, -25
    CreateTextAF Text_07, GetOH() + Col(14) + 7.5, GetOV() + Row(7) + 5.5, "TitleBlock_Text_Company", catMiddleCenter, 3    '��λ
    CreateTextAF Text_07, GetOH() + Col(15) + 21.5, GetOV() + Row(7) + 5.5, "TitleBlock_Text_Company_1", catMiddleCenter, 3    '��λ
    CreateTextAF Text_24, GetOH() + Col(17) + 7.5, GetHeight() - GetOffset(0) + Row(11) + 3.5, "TitleBlock_Text_Secret", catMiddleCenter, 3     '�ܼ�
    CreateTextAF Text_25, GetOH() + Col(17) + 7.5, GetHeight() - GetOffset(0) + Row(12) + 3.5, "TitleBlock_Text_PolishingProcessing", catMiddleCenter, 3     '����ӹ�
    Texts.GetItem("TitleBlock_Text_PolishingProcessing").SetParameterOnSubString catCharSpacing, 0, 0, -25
    CreateTextAF Text_26, GetOH() + Col(17) + 7.5, GetHeight() - GetOffset(0) + Row(13) + 3.5, "TitleBlock_Text_HeatTreatment", catMiddleCenter, 3     '�ȴ���
    CreateTextAF Text_27, GetOH() + Col(17) + 7.5, GetHeight() - GetOffset(0) + Row(14) + 3.5, "TitleBlock_Text_SurfProcessing", catMiddleCenter, 3     '���洦��
    Texts.GetItem("TitleBlock_Text_SurfProcessing").SetParameterOnSubString catCharSpacing, 0, 0, -25




    ' Insert Text Attribute link on sheet's scale
    Set Text = CreateTextAF("", GetOH() + Col(8) + 15#, GetOV() + Row(2) + 3.5, "TitleBlock_Text_Scale_1", catMiddleCenter, 3) '����

    Select Case GetContext():
     Case "LAY": Text.InsertVariable 0, 0, ActiveDoc.Part.GetItem("CATLayoutRoot").Parameters.Item(ActiveDoc.Part.GetItem("CATLayoutRoot").Name + "\" + Sheet.Name + "\ViewMakeUp2DL.1\Scale")
     Case "DRW": Text.InsertVariable 0, 0, ActiveDoc.DrawingRoot.Parameters.Item("Drawing\" + Sheet.Name + "\ViewMakeUp.1\Scale")
     Case Else: Text.Text = "XX"
    End Select

    CreateTextAF Text_23, GetOH() - 21.5, GetOV() + 3.5, "TitleBlock_Text_EnoviaV5_Effectivity", catMiddleCenter, 3
    CreateTextAF Text_23, GetOffset(1) + 22, GetHeight() - GetOffset(0) - 6.5, "TitleBlock_Text_EnoviaV5_Effectivity_1", catMiddleCenter, 3
    Texts.GetItem("TitleBlock_Text_EnoviaV5_Effectivity_1").TextProperties.Mirror = catTextHorizontalAndVerticalFlip
    ' CreateTextAF(Text_23,GetOffset(1) + 22,       GetHeight() - GetOffset(0) - 6.5 ,  "TitleBlock_Text_EnoviaV5_Effectivity_1",     catMiddleCenter,      3).Propertie.Mirror = catTextHorizontalAndVerticalFlip
    CreateTextAF Text_11, GetOH() + Col(8) + 6#, GetOV() + Row(4), "TitleBlock_Text_Weight_1", catMiddleCenter, 3
    CreateTextAF "", GetOH() + Col(7) + 20#, GetOV() + 3.5, "TitleBlock_Text_Sheet_1", catMiddleCenter, 3
    ' CreateTextAF Text_04,GetOH() + Col(2) + 1.       ,GetOV() + Row(2)            ,"TitleBlock_Text_Weight"      ,catTopLeft     ,1.5
    ' CreateTextAF Text_06,GetOH() + Col(3) + 1.       ,GetOV() + Row(2)            ,"TitleBlock_Text_Number"      ,catTopLeft     ,1.5
    ' CreateTextAF Text_07,GetOH() + Col(4) + 1.       ,GetOV() + Row(2)            ,"TitleBlock_Text_Sheet"       ,catTopLeft     ,1.5
    ' CreateTextAF Text_08,GetOH() + Col(1) + 1.       ,GetOV() + Row(3)            ,"TitleBlock_Text_Size"        ,catTopLeft     ,1.5

    ' ��ʾͼ����С
    ' If (Sheet.PaperSize  = 13) Then
    '   CreateTextAF Text_09, GetOH() + .5*(Col(1)+Col(2)),   GetOV() + Row(2) + 2    ,"TitleBlock_Text_Size_1"      ,catBottomCenter, 5
    ' Else
    '   CreateTextAF Text_10, GetOH() + .5*(Col(1)+Col(2)),   GetOV() + Row(2) + 2    ,"TitleBlock_Text_Size_1"      ,catBottomCenter, 5
    ' End If

    ' CreateTextAF Text_12,GetOH() + Col(1) + 1.       ,GetOV() + Row(4)            ,"TitleBlock_Text_Controller"  ,catTopLeft     ,1.5
    ' CreateTextAF Text_05,GetOH() + Col(2) + 2.5      ,GetOV() + .5*(Row(3)+Row(4)),"TitleBlock_Text_Controller_1",catBottomCenter,3
    ' CreateTextAF Text_13,GetOH() + Col(1) + 1.       ,GetOV() + .5*(Row(3)+Row(4)),"TitleBlock_Text_CDate"       ,catTopLeft     ,1.5
    ' CreateTextAF Text_05,GetOH() + Col(2) + 2.5      ,GetOV() + Row(3)            ,"TitleBlock_Text_CDate_1"     ,catBottomCenter,3
    ' ' CreateTextAF Text_14,GetOH() + Col(1) + 1.       ,GetOV() + Row(5)            ,"TitleBlock_Text_Designer"    ,catTopLeft     ,1.5
    ' CreateTextAF Text_15,GetOH() + Col(2) + 2.5      ,GetOV() + .5*(Row(4)+Row(5)),"TitleBlock_Text_Designer_1"  ,catBottomCenter,3
    ' CreateTextAF Text_13,GetOH() + Col(1) + 1.       ,GetOV() + .5*(Row(4)+Row(5)),"TitleBlock_Text_DDate"       ,catTopLeft     ,1.5
    ' CreateTextAF ""&Date,GetOH() + Col(2) + 2.5      ,GetOV() + Row(4)            ,"TitleBlock_Text_DDate_1"     ,catBottomCenter,3
    ' CreateTextAF Text_05,GetOH() + .5*(Col(3)+Col(5)),GetOV() + Row(4)            ,"TitleBlock_Text_Title_1"     ,catMiddleCenter,7

    ' ���ı��
    ' For ii = 1 To GetNbOfRevision()
    '   iY=GetOV() + (ii-.5) * Row(5)/GetNbOfRevision()
    '   CreateTextAF GetRevLetter(ii),GetOH() + .5*(Col(5)+Col(6)),iY,"TitleBlock_Text_Modif_" + GetRevLetter(ii),catMiddleCenter,2.5
    '   CreateTextAF "_"             ,GetOH() + .5*Col(6)         ,iY,"TitleBlock_Text_MDate_" + GetRevLetter(ii),catMiddleCenter,2.
    ' Next

    CATLinks
End Sub

    Sub CATDeleteRevisionBlockFrame()
    DeleteAll "CATDrwSearch.2DGeometry.Name=RevisionBlock_Line_*"
End Sub

Sub CATCreateRevisionBlockFrame()
'-------------------------------------------------------------------------------
'How to draw the revision block geometry
'-------------------------------------------------------------------------------

Revision = CATCheckRev()
If Revision = 0 Then Exit Sub
    For ii = 0 To Revision
        iX = GetOH()
        iY1 = GetHeight() - GetOV() - GetRevRowHeight() * ii
        iY2 = GetHeight() - GetOV() - GetRevRowHeight() * (ii + 1)
        CreateLine iX + GetColRev(1), iY1, iX + GetColRev(1), iY2, "RevisionBlock_Line_Column_" + GetRevLetter(ii) + "_1"
        CreateLine iX + GetColRev(2), iY1, iX + GetColRev(2), iY2, "RevisionBlock_Line_Column_" + GetRevLetter(ii) + "_2"
        CreateLine iX + GetColRev(3), iY1, iX + GetColRev(3), iY2, "RevisionBlock_Line_Column_" + GetRevLetter(ii) + "_3"
        CreateLine iX + GetColRev(4), iY1, iX + GetColRev(4), iY2, "RevisionBlock_Line_Column_" + GetRevLetter(ii) + "_4"
        CreateLine iX + GetColRev(1), iY2, iX, iY2, "RevisionBlock_Line_Row_" + GetRevLetter(ii)
        Next
End Sub

Sub CATAddRevisionBlockText()
'-------------------------------------------------------------------------------
'How to fill in the revision block
'-------------------------------------------------------------------------------
Revision = CATCheckRev() + 1
X = GetOH()
Y = GetHeight() - GetOV() - GetRevRowHeight() * (Revision - 0.5)

Init = InputBox("This review has been done by:", "Reviewer's name", "XXX")
Description = InputBox("Comment to be inserted:", "Description", "None")

If Revision = 1 Then
    CreateTextAF "REV", X + GetColRev(1) + 1#, Y, "RevisionBlock_Text_Rev", catMiddleLeft, 5
    CreateTextAF "DATE", X + GetColRev(2) + 1#, Y, "RevisionBlock_Text_Date", catMiddleLeft, 5
    CreateTextAF "DESCRIPTION", X + GetColRev(3) + 1#, Y, "RevisionBlock_Text_Description", catMiddleLeft, 5
    CreateTextAF "INIT", X + GetColRev(4) + 1#, Y, "RevisionBlock_Text_Init", catMiddleLeft, 5
End If

CreateTextAF GetRevLetter(Revision), X + 0.5 * (GetColRev(1) + GetColRev(2)), Y - GetRevRowHeight(), "RevisionBlock_Text_Rev_" + GetRevLetter(Revision), catMiddleCenter, 5
CreateTextAF "" & Date, X + 0.5 * (GetColRev(2) + GetColRev(3)), Y - GetRevRowHeight(), "RevisionBlock_Text_Date_" + GetRevLetter(Revision), catMiddleCenter, 3.5
CreateTextAF Description, X + GetColRev(3) + 1#, Y - GetRevRowHeight(), "RevisionBlock_Text_Description_" + GetRevLetter(Revision), catMiddleLeft, 2.5
CreateTextAF Init, X + 0.5 * GetColRev(4), Y - GetRevRowHeight(), "RevisionBlock_Text_Init_" + GetRevLetter(Revision), catMiddleCenter, 5

On Error Resume Next
Texts.GetItem("TitleBlock_Text_MDate_" + GetRevLetter(Revision)).Text = "" & Date
If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Sub

Sub ComputeTitleBlockTranslation(TranslationTab As Variant)
    TranslationTab(0) = 0#
    TranslationTab(1) = 0#

    On Error Resume Next
    Set Text = Texts.GetItem("Reference_" + GetMacroID()) 'Get the reference text
    If Err.Number <> 0 Then
        Err.Clear
    Else
        TranslationTab(0) = GetWidth() - GetOffset() - Text.X
        TranslationTab(1) = GetOffset() - Text.Y
        Text.X = Text.X + TranslationTab(0)
        Text.Y = Text.Y + TranslationTab(1)
    End If
    On Error GoTo 0
End Sub

Sub ComputeRevisionBlockTranslation(TranslationTab As Variant)
    TranslationTab(0) = 0#
    TranslationTab(1) = 0#

    On Error Resume Next
    Set Text = Texts.GetItem("RevisionBlock_Text_Init") 'Get the reference text
    If Err.Number <> 0 Then
        Err.Clear
    Else
        TranslationTab(0) = GetWidth() - GetOffset() + GetColRev(4) - Text.X
        TranslationTab(1) = GetHeight() - GetOffset() - 0.5 * GetRevRowHeight() - Text.Y
    End If
    On Error GoTo 0
End Sub



Sub CATRemoveFrame()
    '-------------------------------------------------------------------------------
    'How to remove the whole frame
    '-------------------------------------------------------------------------------
    DeleteAll "CATDrwSearch.DrwText.Name=Frame_Text_*"
    DeleteAll "CATDrwSearch.2DGeometry.Name=Frame_*"
    DeleteAll "CATDrwSearch.2DPoint.Name=TitleBlock_Line_*"
    DeleteAll "CATDrwSearch.2DPoint.Name=SecTitleBlock_Line_*"
    DeleteAll "CATDrwSearch.2DPoint.Name=PartNumberBlock_Line_*"
End Sub

Sub CATDeleteTitleBlockStandard()
    '-------------------------------------------------------------------------------
    'How to remove the standard representation
    '-------------------------------------------------------------------------------
    DeleteAll "CATDrwSearch.2DGeometry.Name=TitleBlock_Standard*"
End Sub

Sub CATMoveTitleBlockText(Translation As Variant)
    '-------------------------------------------------------------------------------
    'How to translate the whole title block after changing the page setup
    '-------------------------------------------------------------------------------
    SelectAll "CATDrwSearch.DrwText.Name=TitleBlock_Text_*"
    Count = Selection.Count2
    For ii = 1 To Count
        Set Text = Selection.Item2(ii).Value
        Text.X = Text.X + Translation(0)
        Text.Y = Text.Y + Translation(1)
    Next
End Sub

Sub CATMoveViews(Translation As Variant)
    '-------------------------------------------------------------------------------
    'How to translate the views after changing the page setup
    '-------------------------------------------------------------------------------
    For i = 3 To Views.Count
        Views.Item(i).UnAlignedWithReferenceView
        Next
        For i = 3 To Views.Count
            Set View = Views.Item(i)
            View.X = View.X + Translation(0)
            View.Y = View.Y + Translation(1)
            View.AlignedWithReferenceView
            Next
End Sub

Sub CATMoveRevisionBlockText(Translation As Varient)
    '-------------------------------------------------------------------------------
    'How to translate the whole revision block after changing the page setup
    '-------------------------------------------------------------------------------
    SelectAll "CATDrwSearch.DrwText.Name=RevisionBlock_Text_*"
    Count = Selection.Count2
    For ii = 1 To Count
        Set Text = Selection.Item2(ii).Value
        Text.X = Text.X + Translation(0)
        Text.Y = Text.Y + Translation(1)
        Next
End Sub

Sub CATLinks()
    '-------------------------------------------------------------------------------
    'How to fill in texts with data of the part/product linked with current sheet
    '-------------------------------------------------------------------------------
    On Error Resume Next
    Dim ViewDocument

    Select Case GetContext():
     Case "LAY": Set ViewDocument = CATIA.ActiveDocument.Product
     Case "DRW":
        If Views.Count >= 3 Then
            Set ViewDocument = Views.Item(3).GenerativeBehavior.Document
        Else
            Set ViewDocument = Nothing
        End If
     Case Else: Set ViewDocument = Nothing
    End Select

    'Find the product document

    Dim ProductDrawn
    Set ProductDrawn = Nothing
    For i = 1 To 8
        If TypeName(ViewDocument) = "PartDocument" Then
            Set ProductDrawn = ViewDocument.Product
            Exit For
        End If
        If TypeName(ViewDocument) = "Product" Then
            Set ProductDrawn = ViewDocument
            Exit For
        End If
        Set ViewDocument = ViewDocument.Parent
        Next

        If ProductDrawn <> Nothing Then
            Texts.GetItem("TitleBlock_Text_EnoviaV5_Effectivity").Text = ProductDrawn.PartNumber + ProductDrawn.Revision
            Texts.GetItem("TitleBlock_Text_EnoviaV5_Effectivity_1").Text = ProductDrawn.PartNumber + ProductDrawn.Revision
            ' Texts.GetItem("TitleBlock_Text_EnoviaV5_Effectivity_1").SetFontSize =12
            Texts.GetItem("TitleBlock_Text_EnoviaV5_Effectivity_1").Text.Mirror = catTextHorizontalFlip
            Texts.GetItem("TitleBlock_Text_Name_1").Text = ProductDrawn.Definition
            ' Texts.GetItem("TitleBlock_Text_Title_1").Text  = ProductDrawn.Definition
            Dim ProductAnalysis As Analyze
            Set ProductAnalysis = ProductDrawn.Analyze
            Texts.GetItem("TitleBlock_Text_Weight_1").Text = FormatNumber(ProductAnalysis.Mass, 2)

            ' �� Parameters ��ȡ������Ϣ���������Ű취���������ã�Ҳ�����������û��Զ������
            Dim userMaterial As String
            userMaterial = ProductDrawn.Parameters.Item("����").ValueAsString
            If (userMaterial <> "") Then
                Texts.GetItem("TitleBlock_Text_Material_1").Text = userMaterial
            Else
                manual_Material = InputBox("û�в�����Ϣ���������ֶ�����" & Chr(13) & _
                " " & Chr(13) & _
                "�������㲿�����ϲ������ٸ��¹���ͼ", "���Ϲ�������", "�������������")
                Texts.GetItem("TitleBlock_Text_Material_1").Text = manual_Material
            End If

            ' ��ʱ�����ֶ��������
            Texts.GetItem("TitleBlock_Text_PCS_1").Text = InputBox("�������㲿���ӹ�����" & Chr(13) & _
            " " & Chr(13) & _
            "Ĭ��Ϊ1��", "�ӹ�����", "1")
        End If

        '-------------------------------------------------------------------------------
        'Display sheet format
        '-------------------------------------------------------------------------------
        Dim textFormat As DrawingText
        Set textFormat = Texts.GetItem("TitleBlock_Text_Size_1")
        textFormat.Text = GetDisplayFormat()
        If Len(GetDisplayFormat()) > 4 Then
            textFormat.SetFontSize 0, 0, 3.5
        Else
            textFormat.SetFontSize 0, 0, 5#
        End If

        '-------------------------------------------------------------------------------
        'Display sheet numbering
        '-------------------------------------------------------------------------------
        Dim nbSheet  As Integer
        Dim curSheet As Integer
        If Not DrwSheet.IsDetail Then
            For Each itSheet In Sheets
                If Not itSheet.IsDetail Then nbSheet = nbSheet + 1
                    Next
                    For Each itSheet In Sheets
                        If Not itSheet.IsDetail Then
                            curSheet = curSheet + 1
                            itSheet.Views.Item(2).Texts.GetItem("TitleBlock_Text_Sheet_1").Text = CStr(curSheet) & "         " & CStr(nbSheet)
                        End If
                        Next
                    End If
                    On Error GoTo 0
End Sub

Sub CATFillField(string1 As String, string2 As String, string3 As String)
    '-------------------------------------------------------------------------------
    'How to call a dialog to fill in manually a given text
    '-------------------------------------------------------------------------------
    Dim TextToFill_1 As DrawingText
    Dim TextToFill_2 As DrawingText
    Dim Person As String

    Set TextToFill_1 = Texts.GetItem(string1)
    Set TextToFill_2 = Texts.GetItem(string2)

    Person = TextToFill_1.Text
    If Person = "XXX" Then Person = "John Smith"

        Person = InputBox("This Document has been " + string3 + " by:", "Controller's name", Person)
        If Person = "" Then Person = "XXX"

            TextToFill_1.Text = Person
            TextToFill_2.Text = "" & Date
End Sub


Sub CATColorGeometry()
    '-------------------------------------------------------------------------------
    'How to color all geometric elements of the active view
    '-------------------------------------------------------------------------------

    ' Uncomment the following sections if needed
    Select Case GetContext():
        'Case "DRW":
        '    SelectAll "CATDrwSearch.2DGeometry"
        '    Selection.VisProperties.SetRealColor 0,0,0,0
        '    Selection.Clear
     Case "LAY":
        SelectAll "CATDrwSearch.2DGeometry"
        Selection.VisProperties.SetRealColor 255, 255, 255, 0
        Selection.Clear
        'Case "SCH":
        '    SelectAll "CATDrwSearch.2DGeometry"
        '    Selection.VisProperties.SetRealColor 0,0,0,0
        '    Selection.Clear

    End Select

End Sub

