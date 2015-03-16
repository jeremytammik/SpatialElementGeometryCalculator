#region Namespaces
using System;
using System.Collections.Generic;
using System.Diagnostics;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.DB.Architecture;
#endregion

namespace SpatialElementGeometryCalculator
{
  [Transaction( TransactionMode.Manual )]
  public class Command : IExternalCommand
  {
    private double sqFootToSquareM(double sqFoot)
    {
        return Math.Round((double) (sqFoot * 0.092903), 2);
    }

    private double calwallAreaMinusOpenings(double subfaceArea, Element ele, Document doc, Room room)
    {
        FilteredElementCollector fiCol = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance));
        List<ElementId> lstTotempDel = new List<ElementId>();
        foreach (FamilyInstance fi in fiCol)
        {
            if (fi.get_Parameter(BuiltInParameter.HOST_ID_PARAM).AsValueString() == ele.Id.ToString())
            {
                if ((fi.Room != null) && (fi.Room.Id == room.Id))
                {
                    lstTotempDel.Add(fi.Id);
                }
                else if ((fi.FromRoom != null) && (fi.FromRoom.Id == room.Id))
                {
                    lstTotempDel.Add(fi.Id);
                }
                else if ((fi.ToRoom != null) && (fi.ToRoom.Id == room.Id))
                {
                    lstTotempDel.Add(fi.Id);
                }
            }
        }
        if (lstTotempDel.Count > 0)
        {
            Transaction t = new Transaction(doc, "tmp Delete");
            double wallnetArea = ele.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble();
            t.Start();
            doc.Delete(lstTotempDel);
            doc.Regenerate();
            double wallGrossArea = ele.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble();
            t.RollBack();
            double fiArea = wallGrossArea - wallnetArea;
            return (subfaceArea - fiArea);
        }
        return subfaceArea;
    }


    double calwallAreaMinusOpenings( double subfaceArea, Element ele, Document doc, Room room )
    {
      FilteredElementCollector fiCol = new FilteredElementCollector( doc ).OfClass( typeof( FamilyInstance ) );
      List<ElementId> lstTotempDel = new List<ElementId>();
      foreach( FamilyInstance fi in fiCol )
      {
        if( fi.get_Parameter( BuiltInParameter.HOST_ID_PARAM ).AsValueString() == ele.Id.ToString() )
        {
          if( ( fi.Room != null ) && ( fi.Room.Id == room.Id ) )
          {
            lstTotempDel.Add( fi.Id );
          }
          else if( ( fi.FromRoom != null ) && ( fi.FromRoom.Id == room.Id ) )
          {
            lstTotempDel.Add( fi.Id );
          }
          else if( ( fi.ToRoom != null ) && ( fi.ToRoom.Id == room.Id ) )
          {
            lstTotempDel.Add( fi.Id );
          }
        }
      }
      if( lstTotempDel.Count > 0 )
      {
        Transaction t = new Transaction( doc, "tmp Delete" );
        double wallnetArea = ele.get_Parameter( BuiltInParameter.HOST_AREA_COMPUTED ).AsDouble();
        t.Start();
        doc.Delete( lstTotempDel );
        doc.Regenerate();
        double wallGrossArea = ele.get_Parameter( BuiltInParameter.HOST_AREA_COMPUTED ).AsDouble();
        t.RollBack();
        double fiArea = wallGrossArea - wallnetArea;
        return ( subfaceArea - fiArea );
      }
      return subfaceArea;
    }


    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        Result Execute;
        try
        {
            UIApplication app = commandData.Application;
            Document doc = app.ActiveUIDocument.Document;
            FilteredElementCollector roomCol = new FilteredElementCollector(app.ActiveUIDocument.Document).OfClass(typeof(SpatialElement));
            string s = "Finished populating Rooms with Boundary Data\r\n\r\n";
            foreach (SpatialElement e in roomCol)
            {
                Room room = e as Room;
                if (room != null)
                {
                    try
                    {
                        IEnumerator VB$t_ref$L1;
                        SpatialElementBoundaryOptions spatialElementBoundaryOptions = new SpatialElementBoundaryOptions();
                        spatialElementBoundaryOptions.SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish;
                        SpatialElementGeometryResults results = new SpatialElementGeometryCalculator(doc, spatialElementBoundaryOptions).CalculateSpatialElementGeometry(room);
                        Solid roomSolid = results.GetGeometry();
                        try
                        {
                            VB$t_ref$L1 = roomSolid.Faces.GetEnumerator();
                            while (VB$t_ref$L1.MoveNext())
                            {
                                IEnumerator<SpatialElementBoundarySubface> VB$t_ref$L2;
                                Face face = (Face) VB$t_ref$L1.Current;
                                IList<SpatialElementBoundarySubface> subfaceList = results.GetBoundaryFaceInfo(face);
                                try
                                {
                                    VB$t_ref$L2 = subfaceList.GetEnumerator();
                                    while (VB$t_ref$L2.MoveNext())
                                    {
                                        SpatialElementBoundarySubface subface = VB$t_ref$L2.Current;
                                        if (subface.SubfaceType == SubfaceType.Side)
                                        {
                                            Element ele = doc.GetElement(subface.SpatialBoundaryElement.HostElementId);
                                            double subfaceArea = subface.GetSubface().Area;
                                            double netArea = this.sqFootToSquareM(this.calwallAreaMinusOpenings(subfaceArea, ele, doc, room));
                                            s = s + "Room " + room.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsString() + " : Wall " + ele.get_Parameter(BuiltInParameter.ALL_MODEL_MARK).AsString() + " : Area " + netArea.ToString() + "m2\r\n";
                                        }
                                    }
                                }
                                finally
                                {
                                    if (VB$t_ref$L2 != null)
                                    {
                                        VB$t_ref$L2.Dispose();
                                    }
                                }
                            }
                        }
                        finally
                        {
                            if (VB$t_ref$L1 is IDisposable)
                            {
                                (VB$t_ref$L1 as IDisposable).Dispose();
                            }
                        }
                        s = s + "\r\n";
                    }
                    catch (Exception exception1)
                    {
                        ProjectData.SetProjectError(exception1);
                        Exception exception = exception1;
                        ProjectData.ClearProjectError();
                    }
                }
            }
            Interaction.MsgBox(s, MsgBoxStyle.Information, "Room Boundaries");
            Execute = Result.Succeeded;
        }
        catch (Exception exception2)
        {
            ProjectData.SetProjectError(exception2);
            Exception ex = exception2;
            Interaction.MsgBox(ex.Message.ToString() + "\r\n" + ex.StackTrace.ToString(), MsgBoxStyle.OkOnly, null);
            Execute = Result.Failed;
            ProjectData.ClearProjectError();
            return Execute;
            ProjectData.ClearProjectError();
        }
        return Execute;
    }
  }
}
