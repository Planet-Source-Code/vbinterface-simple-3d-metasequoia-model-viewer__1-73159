Attribute VB_Name = "modMQO"
Option Explicit
    Private Type TriangleStruct
        V0 As Long
        V1 As Long
        V2 As Long
        OBJID As Integer
        Color As Long
        Normal As D3DVECTOR4
        tNormal As D3DVECTOR4
    End Type
    Public Type POINTAPI
        X As Long
        Y As Long
    End Type
    Private Type LOGBRUSH
        lbStyle As Long
        lbColor As Long
        lbHatch As Long
    End Type
    '---------------------------------------------------
    Public Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As Any, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
    Private Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
    Private Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
    Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
    Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
    Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
    Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
    Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
    Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
    '----------------------------------------------------
    Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
    Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
    '----------------------------------------------------
    Private Const WINDING = 2
    Dim tBrush As Long
    Dim tRgn As Long
    Dim CB As LOGBRUSH
    '----------------------------------------------------
    Dim RotMAT As D3DMATRIX
    Public ScaleMAT As D3DMATRIX
    Dim DXMAT As D3DMATRIX
    Public LightVect As D3DVECTOR4, tLightVect As D3DVECTOR4, NormLight As D3DVECTOR4
    Public ViewVect As D3DVECTOR4, tViewVect As D3DVECTOR4, NormView As D3DVECTOR4
    '----------------------------------------------------
    Dim Triangle() As TriangleStruct
    Dim Vertex() As D3DVECTOR4
    Dim tVertex() As D3DVECTOR4
    Dim T As Long
    Public TriangleCount As Long, VertexCount As Long
    Dim VisibleTrilist() As Double
    Dim VisibleTriCount As Long
    Public CenX As Long, CenY As Long
    Public sW As Long, sH As Long
    Public Const PIBY180 = 0.017453292519943
Sub LoadLights()
    LightVect.X = 0: LightVect.Y = 0: LightVect.z = -300
    D3DXVec4Normalize NormLight, LightVect
End Sub
Function LoadMQO(ByVal ObjectPath As String, ByVal Layer As Object, ByVal Reload As Boolean) As String
    On Error GoTo ErrHandler
    Dim fNum As Integer
    Dim TextLine As String, TempLine As String
    Dim VertexBlock As String, FaceBlock As String, ColorBlock As String
    Dim Location() As String, ObjectName As String
    Dim ObjectVertices As Long, ObjectTriangles As Long
    Dim i As Integer, L As Integer
    Dim StartObject As Boolean, StartVertices As Boolean, StartFaces As Boolean
    Dim vCount As Long
    Dim PreviousVertexCount As Long, PreviousTriangleCount As Long
    Dim PrevStr As String
    Dim V0 As Long, V1 As Long, V2 As Long
    Dim TN0 As D3DVECTOR, TN1 As D3DVECTOR, TN2 As D3DVECTOR
    Dim TN As D3DVECTOR, RTN0 As D3DVECTOR, RTN1 As D3DVECTOR
    Dim tNorm As D3DVECTOR
    Dim TriColor As Long
    Dim RV As Single, GV As Single, BV As Single
    VertexCount = 0: TriangleCount = 0: vCount = 0
    ReDim Vertex(0): ReDim Triangle(0)
    fNum = FreeFile
    Open ObjectPath For Input As fNum
        Do While Not EOF(fNum)
            Line Input #fNum, TempLine
            L = Len(TempLine)
            If L > 0 Then
                If StartVertices = True And Right(TempLine, 1) = "}" Then StartVertices = False
                If StartFaces = True And Right(TempLine, 1) = "}" Then StartFaces = False
                If StartObject = True Then
                    StartObject = False
                ElseIf StartVertices = True Then
                    VertexBlock = Mid(TempLine, 3, L - 2)
                    Location = Split(VertexBlock)
                    Vertex(VertexCount).X = (Location(0))
                    Vertex(VertexCount).Y = (Location(1))
                    Vertex(VertexCount).z = (Location(2))
                    Vertex(VertexCount).w = 1
                    VertexCount = VertexCount + 1
                ElseIf StartFaces = True Then
                    FaceBlock = Mid(TempLine, 3, L - 2)
                    If Val(Left(FaceBlock, 1)) > 3 Then
                        MsgBox "Triangulate Mesh. Face contains more than 3 vertices"
                        LoadMQO = -1
                        ClearObject
                        Exit Function
                    End If
                    Location = Split(FaceBlock)
                    V0 = Right(Location(1), Val(Len(Location(1)) - 2)) + vCount
                    V1 = Val(Location(2)) + vCount
                    If Right(Location(2), 1) = ")" Then
                        V2 = V1
                    Else
                        V2 = Left(Location(3), Val(Len(Location(3)) - 1)) + vCount
                    End If
                    Triangle(TriangleCount).V0 = V0: Triangle(TriangleCount).V1 = V1: Triangle(TriangleCount).V2 = V2
                    Triangle(TriangleCount).OBJID = 1
                    Triangle(TriangleCount).Color = 3
                    '---------calculate face normal------------------------
                    TN0.X = Vertex(V0).X: TN0.Y = Vertex(V0).Y: TN0.z = Vertex(V0).z
                    TN1.X = Vertex(V1).X: TN1.Y = Vertex(V1).Y: TN1.z = Vertex(V1).z
                    TN2.X = Vertex(V2).X: TN2.Y = Vertex(V2).Y: TN2.z = Vertex(V2).z
                    D3DXVec3Subtract RTN0, TN1, TN0
                    D3DXVec3Subtract RTN1, TN2, TN0
                    D3DXVec3Cross TN, RTN0, RTN1
                    D3DXVec3Normalize tNorm, TN
                    Triangle(TriangleCount).Normal.X = tNorm.X
                    Triangle(TriangleCount).Normal.Y = tNorm.Y
                    Triangle(TriangleCount).Normal.z = tNorm.z
                    '------------------------------------------------------
                    TriangleCount = TriangleCount + 1
                End If
                If Left(TempLine, 6) = "Object" Then   'use ucase
                    StartObject = True
                    StartVertices = False
                    StartFaces = False
                ElseIf Asc(Mid(TempLine, 1, 1)) = 9 And Mid(TempLine, 2, 6) = "vertex" Then   'use ucase
                    ObjectVertices = Val(Mid(TempLine, 8, L - 7))
                    If ObjectVertices > 0 Then
                        PreviousVertexCount = VertexCount - 1
                        ReDim Preserve Vertex(PreviousVertexCount + ObjectVertices)
                        StartVertices = True
                        StartFaces = False
                        vCount = VertexCount
                    End If
                ElseIf Asc(Mid(TempLine, 1, 1)) = 9 And Mid(TempLine, 2, 5) = "face " Then   'use ucase
                    ObjectTriangles = (Val(Mid(TempLine, 6, L - 5))) * 3
                    If ObjectTriangles > 0 Then
                        ReDim Preserve Triangle(TriangleCount + ObjectTriangles)
                        StartFaces = True
                        StartVertices = False
                    End If
                End If
            End If
        Loop
    Close fNum
    If VertexCount > 0 Then VertexCount = VertexCount - 1
    If TriangleCount > 0 Then TriangleCount = TriangleCount - 1
    ReDim VisibleTrilist(TriangleCount, 1)
    ReDim tVertex(VertexCount)
    Erase Location
    LoadMQO = VertexCount + 1 & " vertices, " & TriangleCount + 1 & " triangles"
ErrHandler:
    If Err.Number <> 0 Then
        MsgBox "Unable to Load Model"
        ClearObject
        Exit Function
    End If
End Function
Sub TransformWorld(ByVal XR As Single, YR As Single, ZR As Single, TransX As Single, TransY As Single, TransZ As Single)
    On Error GoTo ErrHandler
    Dim V0 As Long, V1 As Long, V2 As Long
    Dim X0 As Single, Y0 As Single, X1 As Single, Y1 As Single, X2 As Single, Y2 As Single
    Dim ZNormal As Single
    Dim TriZ As Double
    Dim TriNorm As D3DVECTOR4, tempVec As D3DVECTOR4
    D3DXMatrixIdentity DXMAT
    D3DXMatrixRotationYawPitchRoll DXMAT, YR, XR, ZR
    D3DXMatrixMultiply DXMAT, DXMAT, ScaleMAT
    DXMAT.m41 = TransX: DXMAT.m42 = TransY: DXMAT.m43 = TransZ
    VisibleTriCount = 0
    For T = 0 To TriangleCount
        V0 = Triangle(T).V0
        V1 = Triangle(T).V1
        V2 = Triangle(T).V2
        D3DXVec4Transform tempVec, Vertex(V0), DXMAT
        tVertex(V0).X = tempVec.X + CenX
        tVertex(V0).Y = tempVec.Y + CenY
        tVertex(V0).z = tempVec.z
        D3DXVec4Transform tempVec, Vertex(V1), DXMAT
        tVertex(V1).X = tempVec.X + CenX
        tVertex(V1).Y = tempVec.Y + CenY
        tVertex(V1).z = tempVec.z
        D3DXVec4Transform tempVec, Vertex(V2), DXMAT
        tVertex(V2).X = tempVec.X + CenX
        tVertex(V2).Y = tempVec.Y + CenY
        tVertex(V2).z = tempVec.z
        X0 = tVertex(V0).X: Y0 = tVertex(V0).Y
        X1 = tVertex(V1).X: Y1 = tVertex(V1).Y
        X2 = tVertex(V2).X: Y2 = tVertex(V2).Y
        TriZ = (tVertex(V0).z + tVertex(V1).z + tVertex(V2).z) / 3
        ZNormal = (X1 - X0) * (Y0 - Y2) - (Y1 - Y0) * (X0 - X2)
        If ZNormal >= 0 Then
            VisibleTrilist(VisibleTriCount, 0) = TriZ
            VisibleTrilist(VisibleTriCount, 1) = T
            VisibleTriCount = VisibleTriCount + 1
            D3DXVec4Transform Triangle(T).tNormal, Triangle(T).Normal, DXMAT
        End If
    Next T
    If VisibleTriCount > 0 Then VisibleTriCount = VisibleTriCount - 1
    QuickSort 0, VisibleTriCount, VisibleTrilist
ErrHandler:
    If Err.Number <> 0 Then
        Exit Sub
    End If
End Sub
Public Sub QuickSort(ByVal ql As Long, ByVal qr As Long, qa() As Double)
    Dim qi As Long, qj As Long
    Dim qT As Double, qm As Double
    Dim TID As Long
    qi = ql
    qj = qr
    qm = qa((ql + qr) \ 2, 0)
    Do
        Do While qa(qi, 0) < qm
            qi = qi + 1
        Loop
        Do While qm < qa(qj, 0)
            qj = qj - 1
        Loop
        If qi <= qj Then
            qT = qa(qi, 0)
            TID = qa(qi, 1)
            qa(qi, 0) = qa(qj, 0)
            qa(qi, 1) = qa(qj, 1)
            qa(qj, 0) = qT
            qa(qj, 1) = TID
            qi = qi + 1
            qj = qj - 1
        End If
    Loop Until qi > qj
    If ql < qj Then QuickSort ql, qj, qa()
    If qi < qr Then QuickSort qi, qr, qa()
End Sub
Sub RenderWorld(ByVal Layer As Object)
    On Error GoTo ErrHandler
    Dim V0 As Long, V1 As Long, V2 As Long
    Dim X0 As Single, Y0 As Single, X1 As Single, Y1 As Single, X2 As Single, Y2 As Single
    Dim VisT As Long
    Dim DotP As Single
    Dim tVector As D3DVECTOR4
    Dim tVectorA As D3DVECTOR4, tVectorB As D3DVECTOR4, tVectorC As D3DVECTOR4
    Dim TriNormal As D3DVECTOR4
    Dim TriPoint(2) As POINTAPI
    Dim TriColor As Integer
    Dim tDc As Long, tBitmap As Long
    tDc = CreateCompatibleDC(Layer.hdc)
    tBitmap = CreateCompatibleBitmap(Layer.hdc, sW, sH)
    SelectObject tDc, tBitmap
    BitBlt tDc, 0, 0, sW, sH, Layer.hdc, 0, 0, vbSrcCopy
    For T = 0 To VisibleTriCount
        VisT = VisibleTrilist(T, 1)
        V0 = Triangle(VisT).V0: V1 = Triangle(VisT).V1: V2 = Triangle(VisT).V2
        X0 = tVertex(V0).X: Y0 = tVertex(V0).Y
        X1 = tVertex(V1).X: Y1 = tVertex(V1).Y
        X2 = tVertex(V2).X: Y2 = tVertex(V2).Y
        '---------------------------------------------------------
        D3DXVec4Normalize TriNormal, Triangle(VisT).tNormal
        DotP = D3DXVec4Dot(TriNormal, NormLight)
        If DotP < 0 Then DotP = 0
        If DotP > 1 Then DotP = 1
        TriColor = DotP * 255
        '---------------------------------------------------------
        TriPoint(0).X = X0: TriPoint(0).Y = Y0
        TriPoint(1).X = X1: TriPoint(1).Y = Y1
        TriPoint(2).X = X2: TriPoint(2).Y = Y2
        tRgn = CreatePolygonRgn(TriPoint(0), 3, WINDING)
        CB.lbColor = RGB(TriColor, TriColor, TriColor)
        tBrush = CreateBrushIndirect(CB)
        If tRgn Then FillRgn tDc, tRgn, tBrush
        DeleteObject tRgn
        DeleteObject tBrush
    Next T
    BitBlt Layer.hdc, 0, 0, sW, sH, tDc, 0, 0, vbSrcCopy
    DeleteDC tDc
    DeleteObject tBitmap
ErrHandler:
    If Err.Number <> 0 Then
        Exit Sub
    End If
End Sub
Sub ClearObject()
    Erase Vertex
    Erase tVertex
    Erase Triangle
    Erase VisibleTrilist
End Sub
