using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.IFC;
using Autodesk.Revit.DB.Architecture;

namespace SpatialElementGeometryCalculator
{
  class OpeningHandler
  {
    public double GetOpeningArea(
      Wall elemHost,
      Element elemInsert,
      Room room,
      Solid roomSolid )
    {
      Document doc = room.Document;
      double openingArea = 0;

      if( elemInsert is FamilyInstance )
      {
        FamilyInstance fi = elemInsert as FamilyInstance;

        if( IsInRoom( room, fi ) )
        {
          if( elemHost is Wall )
          {
            Wall wall = elemHost as Wall;
            openingArea = GetWallCutArea( fi, wall );
          }

          //if( openingArea.Equals( 0 ) )
          //{
          //  openingArea = GetDoorWinAreaFromParameter( doc, fi );
          //}
        }
      }

      if( elemInsert is Wall )
      {
        SolidHandler solidHandler = new SolidHandler();
        openingArea = solidHandler.GetWallAsOpeningArea( 
          elemInsert, roomSolid );
      }
      return openingArea;
    }

    public double GetWallCutArea( 
      FamilyInstance fi, 
      Wall wall )
    {
      Document doc = fi.Document;

      XYZ cutDir = null;

      CurveLoop curveLoop 
        = ExporterIFCUtils.GetInstanceCutoutFromWall( 
          fi.Document, wall, fi, out cutDir );

      IList<CurveLoop> loops = new List<CurveLoop>( 1 );
      loops.Add( curveLoop );

      if( !wall.IsStackedWallMember )
      {
        return ExporterIFCUtils.ComputeAreaOfCurveLoops( loops );
      }
      else
      {
        // Will not get multiple stacked walls with 
        // varying thickness due to the nature of rooms.
        // Use ReferenceIntersector if we can identify 
        // those missing room faces...open for suggestions.

        SolidHandler solHandler = new SolidHandler();
        Options optCompRef 
          = doc.Application.Create.NewGeometryOptions();
        if( null != optCompRef )
        {
          optCompRef.ComputeReferences = true;
          optCompRef.DetailLevel = ViewDetailLevel.Medium;
        }

        GeometryElement geomElemHost 
          = wall.get_Geometry( optCompRef ) 
            as GeometryElement;

        Solid solidOpening = GeometryCreationUtilities
          .CreateExtrusionGeometry( loops, 
            cutDir.Negate(), .1 );

        Solid solidHost 
          = solHandler.CreateSolidFromBoundingBox( 
            null, geomElemHost.GetBoundingBox(), null );

        // We dont really care about the boundingbox 
        // rotation as we only need the intersected solid.

        if( solidHost == null )
        {
          return 0;
        }

        Solid intersectSolid = BooleanOperationsUtils
          .ExecuteBooleanOperation( solidOpening, 
            solidHost, BooleanOperationsType.Intersect );

        if( intersectSolid.Faces.Size.Equals( 0 ) )
        {
          solidOpening = GeometryCreationUtilities
            .CreateExtrusionGeometry( loops, cutDir, .1 );

          intersectSolid = BooleanOperationsUtils
            .ExecuteBooleanOperation( solidOpening, 
              solidHost, BooleanOperationsType.Intersect );
        }

        if( DebugHandler.EnableSolidUtilityVolumes )
        {
          using( Transaction t = new Transaction( doc ) )
          {
            t.Start( "Stacked1" );
            ShapeCreator.CreateDirectShape( doc, 
              intersectSolid, "stackedOpening" );
            t.Commit();
          }
        }
        return solHandler.GetLargestFaceArea( 
          intersectSolid );
      }
    }

    /// <summary>
    /// Predicate to determine whether the given 
    /// family instance belongs to the given room.
    /// </summary>
    static bool IsInRoom(
      Room room,
      FamilyInstance f )
    {
      ElementId id = room.Id;
      return ( ( f.Room != null && f.Room.Id == id )
        || ( f.ToRoom != null && f.ToRoom.Id == id )
        || ( f.FromRoom != null && f.FromRoom.Id == id ) );
    }


    //static double GetDoorWinAreaFromParameter( Document doc, FamilyInstance insert )
    //{
    //  ElementType insertType = doc.GetElement( insert.GetTypeId() ) as ElementType;

    //  //This is somewhat problematic as some familys use Height and some Rough Height.
    //  //Or completely different parameters and in some cases not connected to the family geometry at all.
    //  //A solid intersection with the room is probably the surest thing
    //  Parameter height = insertType.get_Parameter(BuiltInParameter.FAMILY_HEIGHT_PARAM);
    //  Parameter width = insertType.get_Parameter(BuiltInParameter.FAMILY_WIDTH_PARAM);

    //  return width.AsDouble() * height.AsDouble();
    //}

    static Solid GetLargestSolid( 
      GeometryElement geomElem )
    {
      // Not correct if the wall is very thick or 
      // opening is unusually narrow though this 
      // should be quite rare.

      IList<Solid> lstHostSolids = new List<Solid>();
      Solid solidMax = null;
      IList<double> lstVolumes = new List<double>();

      double maxVolVal;

      foreach( GeometryObject geomObj in geomElem )
      {
        Solid geomSolid = geomObj as Solid;

        if( null != geomSolid )
        {
          lstHostSolids.Add( geomSolid );
          lstVolumes.Add( geomSolid.Volume );
        }

        maxVolVal = lstVolumes.Max();

        foreach( Solid sol in lstHostSolids )
        {
          if( sol.Volume == maxVolVal )
          {
            solidMax = sol;
          }
        }
      }
      return solidMax;
    }
  }
}
