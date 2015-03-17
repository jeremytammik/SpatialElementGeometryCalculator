#region Namespaces
using System;
using System.Collections.Generic;
using System.Diagnostics;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
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
    /// Calculate wall area minus openings. Temporarily
    /// delete all openings in a transaction that is
    /// rolled back.
    /// </summary>
    /// <param name="subfaceArea">Initial gross subface area</param>
    /// <param name="wall"></param>
    /// <param name="doc"></param>
    /// <param name="room"></param>
    /// <returns></returns>
    double calwallAreaMinusOpenings( 
      double subfaceArea, 
      Element wall,
      Room room )
    {
      Document doc = wall.Document;

      Debug.Assert( 
        room.Document.ProjectInformation.UniqueId.Equals( 
          doc.ProjectInformation.UniqueId ),
        "expected wall and room from same document" );

      // Determine all openings in the given wall.

      FilteredElementCollector fiCol 
        = new FilteredElementCollector( doc )
          .OfClass( typeof( FamilyInstance ) );

      List<ElementId> lstTotempDel 
        = new List<ElementId>();
      
      foreach( FamilyInstance fi in fiCol )
      {
        // The family instances hosted by this wall
        // could be filtered out more efficiently using 
        // a filtered element collector parameter filter.

        if( fi.get_Parameter( 
          BuiltInParameter.HOST_ID_PARAM )
            .AsElementId().IntegerValue.Equals( 
              wall.Id.IntegerValue ) )
        {
          if( ( fi.Room != null ) 
            && ( fi.Room.Id == room.Id ) )
          {
            lstTotempDel.Add( fi.Id );
          }
          else if( ( fi.FromRoom != null ) 
            && ( fi.FromRoom.Id == room.Id ) )
          {
            lstTotempDel.Add( fi.Id );
          }
          else if( ( fi.ToRoom != null ) 
            && ( fi.ToRoom.Id == room.Id ) )
          {
            lstTotempDel.Add( fi.Id );
          }
        }
      }

      // Determine total area of all openings.

      double openingArea = 0;

      if( 0 < lstTotempDel.Count )
      {
        Transaction t = new Transaction( doc );

        double wallAreaNet = wall.get_Parameter( 
          BuiltInParameter.HOST_AREA_COMPUTED )
            .AsDouble();

        t.Start( "tmp Delete" );
        doc.Delete( lstTotempDel );
        doc.Regenerate();
        double wallAreaGross = wall.get_Parameter( 
          BuiltInParameter.HOST_AREA_COMPUTED )
            .AsDouble();
        t.RollBack();

        openingArea = wallAreaGross - wallAreaNet;
      }

      return subfaceArea - openingArea;
    }

    public Result Execute( 
      ExternalCommandData commandData, 
      ref string message, 
      ElementSet elements )
    {
      UIApplication app = commandData.Application;
      Document doc = app.ActiveUIDocument.Document;

      SpatialElementBoundaryOptions sebOptions 
        = new SpatialElementBoundaryOptions();

      sebOptions.SpatialElementBoundaryLocation 
        = SpatialElementBoundaryLocation.Finish;

      Result rc;

      try
      {
        FilteredElementCollector roomCol 
          = new FilteredElementCollector( doc )
            .OfClass( typeof( SpatialElement ) );

        string s = "Finished populating Rooms with "
          + "Boundary Data\r\n\r\n";

        foreach( SpatialElement e in roomCol )
        {
          Room room = e as Room;

          if( room != null )
          {
            try
            {
              Autodesk.Revit.DB
                .SpatialElementGeometryCalculator 
                  calc = new Autodesk.Revit.DB
                    .SpatialElementGeometryCalculator( 
                      doc, sebOptions );

              SpatialElementGeometryResults results 
                = calc.CalculateSpatialElementGeometry( 
                  room );

              Solid roomSolid = results.GetGeometry();

              foreach( Face face in roomSolid.Faces )
              {
                IList<SpatialElementBoundarySubface> 
                  subfaceList = results.GetBoundaryFaceInfo( 
                    face );

                foreach( SpatialElementBoundarySubface 
                  subface in subfaceList )
                {
                  if( subface.SubfaceType 
                    == SubfaceType.Side )
                  {
                    Element wall = doc.GetElement( 
                      subface.SpatialBoundaryElement
                        .HostElementId );

                    double subfaceArea = subface
                      .GetSubface().Area;

                    double netArea = sqFootToSquareM( 
                      calwallAreaMinusOpenings( 
                        subfaceArea, wall, room ) );

                    s = s + "Room " 
                      + room.get_Parameter( 
                        BuiltInParameter.ROOM_NUMBER )
                          .AsString() 
                      + " : Wall " + wall.get_Parameter( 
                        BuiltInParameter.ALL_MODEL_MARK )
                          .AsString() 
                      + " : Area " + netArea.ToString() 
                      + "m2\r\n";
                  }
                }
              }
              s = s + "\r\n";
            }
            catch( Exception )
            {
            }
          }
        }
        TaskDialog.Show( "Room Boundaries", s);

        rc = Result.Succeeded;
      }
      catch( Exception ex )
      {
        TaskDialog.Show( "Room Boundaries", 
          ex.Message.ToString() + "\r\n" 
          + ex.StackTrace.ToString() );

        rc = Result.Failed;
      }
      return rc;
    }
  }
}
