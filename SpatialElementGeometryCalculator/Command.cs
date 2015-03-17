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
    /// <summary>
    /// Convert square feet to square meters with two decimal places precision.
    /// </summary>
    double sqFootToSquareM( double sqFoot )
    {
      return Math.Round( (double) ( sqFoot * 0.092903 ), 2 );
    }

    /// <summary>
    /// Calculate wall area minus openings
    /// </summary>
    /// <param name="subfaceArea">Initial gross subface area</param>
    /// <param name="wall"></param>
    /// <param name="doc"></param>
    /// <param name="room"></param>
    /// <returns></returns>
    double calwallAreaMinusOpenings( double subfaceArea, Element wall, Document doc, Room room )
    {
      FilteredElementCollector fiCol = new FilteredElementCollector( doc ).OfClass( typeof( FamilyInstance ) );
      List<ElementId> lstTotempDel = new List<ElementId>();
      foreach( FamilyInstance fi in fiCol )
      {
        if( fi.get_Parameter( BuiltInParameter.HOST_ID_PARAM ).AsValueString() == wall.Id.ToString() )
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
        double wallnetArea = wall.get_Parameter( BuiltInParameter.HOST_AREA_COMPUTED ).AsDouble();
        t.Start();
        doc.Delete( lstTotempDel );
        doc.Regenerate();
        double wallGrossArea = wall.get_Parameter( BuiltInParameter.HOST_AREA_COMPUTED ).AsDouble();
        t.RollBack();
        double fiArea = wallGrossArea - wallnetArea;
        return ( subfaceArea - fiArea );
      }
      return subfaceArea;
    }

    public Result Execute( ExternalCommandData commandData, ref string message, ElementSet elements )
    {
      Result rc;
      try
      {
        UIApplication app = commandData.Application;
        Document doc = app.ActiveUIDocument.Document;
        FilteredElementCollector roomCol = new FilteredElementCollector( app.ActiveUIDocument.Document ).OfClass( typeof( SpatialElement ) );
        string s = "Finished populating Rooms with Boundary Data\r\n\r\n";
        foreach( SpatialElement e in roomCol )
        {
          Room room = e as Room;
          if( room != null )
          {
            try
            {
              SpatialElementBoundaryOptions spatialElementBoundaryOptions = new SpatialElementBoundaryOptions();
              spatialElementBoundaryOptions.SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish;
              Autodesk.Revit.DB.SpatialElementGeometryCalculator calculator1 = new Autodesk.Revit.DB.SpatialElementGeometryCalculator( doc, spatialElementBoundaryOptions );
              SpatialElementGeometryResults results = calculator1.CalculateSpatialElementGeometry( room );
              Solid roomSolid = results.GetGeometry();
              foreach( Face face in roomSolid.Faces )
              {
                IList<SpatialElementBoundarySubface> subfaceList = results.GetBoundaryFaceInfo( face );
                foreach( SpatialElementBoundarySubface subface in subfaceList )
                {
                  if( subface.SubfaceType == SubfaceType.Side )
                  {
                    Element wall = doc.GetElement( subface.SpatialBoundaryElement.HostElementId );
                    double subfaceArea = subface.GetSubface().Area;
                    double netArea = this.sqFootToSquareM( calwallAreaMinusOpenings( subfaceArea, wall, doc, room ) );
                    s = s + "Room " + room.get_Parameter( BuiltInParameter.ROOM_NUMBER ).AsString() + " : Wall " + wall.get_Parameter( BuiltInParameter.ALL_MODEL_MARK ).AsString() + " : Area " + netArea.ToString() + "m2\r\n";
                  }
                }
              }
              s = s + "\r\n";
            }
            catch( Exception ex )
            {
            }
          }
        }
        TaskDialog.Show( "Room Boundaries", s);
        rc = Result.Succeeded;
      }
      catch( Exception ex )
      {
         TaskDialog.Show( "Room Boundaries", ex.Message.ToString() + "\r\n" + ex.StackTrace.ToString() );
        rc = Result.Failed;
      }
      return rc;
    }
  }
}
