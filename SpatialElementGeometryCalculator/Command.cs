#region

using System;
using System.Collections.Generic;
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
    public Result Execute( ExternalCommandData commandData, ref string message, ElementSet elements )
    {
      var app = commandData.Application;
      var doc = app.ActiveUIDocument.Document;
      Result rc;
      string s = string.Empty;
      try
      {
        var roomCol = new FilteredElementCollector( doc ).OfClass( typeof( SpatialElement ) );

        foreach( var e in roomCol )
        {
          var room = e as Room;
          if( room == null ) continue;
          if( room.Location == null ) continue;

          var sebOptions = new SpatialElementBoundaryOptions { SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish };
          var calc = new Autodesk.Revit.DB.SpatialElementGeometryCalculator( doc, sebOptions );
          var results = calc.CalculateSpatialElementGeometry( room );

          // To keep track of each wall and its total area in the room
          var walls = new Dictionary<string, double>();

          foreach( Face face in results.GetGeometry().Faces )
          {
            foreach( var subface in results.GetBoundaryFaceInfo( face ) )
            {
              if( subface.SubfaceType != SubfaceType.Side ) { continue; }
              var wall = doc.GetElement( subface.SpatialBoundaryElement.HostElementId ) as HostObject;
              if( wall == null ) { continue; }
              var grossArea = subface.GetSubface().Area;
              if( !walls.ContainsKey( wall.UniqueId ) )
              {
                walls.Add( wall.UniqueId, grossArea );
              }
              else
              {
                walls[wall.UniqueId] += grossArea;
              }
            }
          }

          foreach( var id in walls.Keys )
          {
            var wall = (HostObject) doc.GetElement( id );
            var openings = CalculateWallOpeningArea( wall, room );

            s += string.Format( "Room: {2} Wall: {0} Area: {1} m2\r\n",
                wall.get_Parameter( BuiltInParameter.ALL_MODEL_MARK ).AsString(),
                SqFootToSquareM( walls[id] - openings ),
                room.get_Parameter( BuiltInParameter.ROOM_NUMBER ).AsString() );
          }

        }

        TaskDialog.Show( "Room Boundaries", s );
        rc = Result.Succeeded;
      }
      catch( Exception ex )
      {
        TaskDialog.Show( "Room Boundaries", ex.Message + "\r\n" + ex.StackTrace );
        rc = Result.Failed;
      }

      return rc;
    }

    /// <summary>
    ///     Convert square feet to square meters
    ///     with two decimal places precision.
    /// </summary>
    private static double SqFootToSquareM( double sqFoot )
    {
      return Math.Round( sqFoot * 0.092903, 2 );
    }

    /// <summary>
    ///     Calculate wall area minus openings. Temporarily
    ///     delete all openings in a transaction that is
    ///     rolled back.
    /// </summary>
    /// <param name="wall"></param>
    /// <param name="room"></param>
    private static double CalculateWallOpeningArea( HostObject wall, Room room )
    {
      var doc = wall.Document;
      var wallAreaNet = wall.get_Parameter( BuiltInParameter.HOST_AREA_COMPUTED ).AsDouble();
      var t = new Transaction( doc );
      t.Start( "Temp" );
      foreach( var id in wall.FindInserts( true, true, true, true ) )
      {
        var insert = doc.GetElement( id );
        if( insert is FamilyInstance && IsInRoom( room, (FamilyInstance) insert ) )
        {
          doc.Delete( id );
        }
      }

      doc.Regenerate();
      var wallAreaGross = wall.get_Parameter( BuiltInParameter.HOST_AREA_COMPUTED ).AsDouble();
      t.RollBack();

      return wallAreaGross - wallAreaNet;
    }

    private static bool IsInRoom( Room room, FamilyInstance f )
    {
      return ( ( f.Room != null && f.Room.Id == room.Id ) || ( f.ToRoom != null && f.ToRoom.Id == room.Id ) || ( f.FromRoom != null && f.FromRoom.Id == room.Id ) );
    }
  }
}
