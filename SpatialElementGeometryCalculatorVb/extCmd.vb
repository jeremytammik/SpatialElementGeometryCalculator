Imports Autodesk.Revit.ApplicationServices
Imports System.IO
Imports Autodesk.Revit.DB.Architecture
Imports System.Collections.Generic
Imports System.Diagnostics
Imports Autodesk.Revit.Attributes
Imports Autodesk.Revit.DB
Imports BoundarySegment = Autodesk.Revit.DB.BoundarySegment
Imports Autodesk.Revit.DB.ExternalService
Imports Autodesk.Revit.UI

<Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)> _
<Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)> _
<Autodesk.Revit.Attributes.Journaling(Autodesk.Revit.Attributes.JournalingMode.NoCommandData)> _
Public Class extCmd
    Implements IExternalCommand

    Public Function Execute(commandData As ExternalCommandData, ByRef message As String, elements As ElementSet) As Result Implements IExternalCommand.Execute
        Try

            Dim app As Autodesk.Revit.UI.UIApplication = commandData.Application
            Dim doc As Document = app.ActiveUIDocument.Document

            Dim roomCol As FilteredElementCollector = New FilteredElementCollector(app.ActiveUIDocument.Document).OfClass(GetType(SpatialElement))
            Dim s As String = "Finished populating Rooms with Boundary Data" + vbNewLine + vbNewLine
            For Each e As SpatialElement In roomCol
                Dim room As Room = TryCast(e, Room)
                If room IsNot Nothing Then
                    Try
                        Dim spatialElementBoundaryOptions As New SpatialElementBoundaryOptions()
                        spatialElementBoundaryOptions.SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish
                        Dim calculator1 As New Autodesk.Revit.DB.SpatialElementGeometryCalculator(doc, spatialElementBoundaryOptions)
                        Dim results As SpatialElementGeometryResults = calculator1.CalculateSpatialElementGeometry(room)

                        ' get the solid representing the room's geometry
                        Dim roomSolid As Solid = results.GetGeometry()
                        Dim iSegment As Integer = 1
                        For Each face As Face In roomSolid.Faces
                            Dim subfaceList As IList(Of SpatialElementBoundarySubface) = results.GetBoundaryFaceInfo(face)
                            For Each subface As SpatialElementBoundarySubface In subfaceList
                                If subface.SubfaceType = SubfaceType.Side Then ' only interested in walls in this example
                                    Dim ele As Element = doc.GetElement(subface.SpatialBoundaryElement.HostElementId)
                                    Dim subfaceArea As Double = subface.GetSubface().Area
                                    Dim netArea As Double = sqFootToSquareM(calwallAreaMinusOpenings(subfaceArea, ele, doc, room))
                                    s += "Room " + room.Parameter(BuiltInParameter.ROOM_NUMBER).AsString + " : Wall " + ele.Parameter(BuiltInParameter.DOOR_NUMBER).AsString + " : Area " + netArea.ToString + "m2" + vbNewLine
                                End If

                            Next
                        Next
                        s += vbNewLine
                    Catch ex As Exception
                    End Try

                End If

            Next
            MsgBox(s, MsgBoxStyle.Information, "Room Boundaries")
            Return Result.Succeeded
        Catch ex As Exception
            MsgBox(ex.Message.ToString + vbNewLine +
           ex.StackTrace.ToString)
            Return Result.Failed
        End Try
    End Function

   
    Private Function calwallAreaMinusOpenings(ByVal subfaceArea As Double, ByVal ele As Element, ByVal doc As Document, ByVal room As Room) As Double
        Dim fiCol As FilteredElementCollector = New FilteredElementCollector(doc).OfClass(GetType(FamilyInstance))
        Dim lstTotempDel As New List(Of ElementId)
        'Now find the familyInstances that are associated to the current room
        For Each fi As FamilyInstance In fiCol
            If fi.Parameter(BuiltInParameter.HOST_ID_PARAM).AsValueString = ele.Id.ToString Then
                If fi.Room IsNot Nothing Then
                    If fi.Room.Id = room.Id Then
                        lstTotempDel.Add(fi.Id)
                        Continue For
                    End If
                End If
                If fi.FromRoom IsNot Nothing Then
                    If fi.FromRoom.Id = room.Id Then
                        lstTotempDel.Add(fi.Id)
                        Continue For
                    End If
                End If
                If fi.ToRoom IsNot Nothing Then
                    If fi.ToRoom.Id = room.Id Then
                        lstTotempDel.Add(fi.Id)
                        Continue For
                    End If
                End If
            End If
        Next

        If lstTotempDel.Count > 0 Then
            Dim t As New Transaction(doc, "tmp Delete")
            Dim wallnetArea As Double = ele.Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble
            t.Start()
            doc.Delete(lstTotempDel)
            doc.Regenerate()
            Dim wallGrossArea As Double = ele.Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble
            t.RollBack()
            Dim fiArea As Double = wallGrossArea - wallnetArea
            Return subfaceArea - fiArea
        Else
            Return subfaceArea
        End If
    End Function



    Private Function sqFootToSquareM(ByVal sqFoot As Double) As Double
        Return Math.Round(sqFoot * 0.092903, 2)
    End Function

End Class


