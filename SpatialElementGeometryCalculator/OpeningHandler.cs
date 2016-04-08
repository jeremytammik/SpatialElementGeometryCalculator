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
      SolidHandler solidHandler, 
      Element elemHost, 
      Element elemInsert, 
      Room room, 
      Solid roomSolid, 
      bool isStacked )
    {
      Document doc = room.Document;
      double openingArea = 0;

      Parameter demoInsert = elemInsert.get_Parameter( 
        BuiltInParameter.PHASE_DEMOLISHED ) as Parameter;

      if( demoInsert != null )
      {
        if( !demoInsert.AsValueString().Equals( "None" ) )
        {
          return openingArea;
        }
      }

      if( elemInsert is FamilyInstance )
      {
        FamilyInstance fi = elemInsert as FamilyInstance;

        if( IsToOrFromThisRoom( room, fi ) )
        {
          if( elemHost is Wall )
          {
            try
            {
              Wall wall = elemHost as Wall;
              openingArea = GetWallCutArea( doc, fi, wall, isStacked );
            }
            catch
            {
              openingArea = 0;
              LogCreator.LogEntry( "WallCut Failed on " 
                + elemHost.Id.ToString() );
            }
          }

          if( openingArea.Equals( 0 ) )
          {
            openingArea = GetDoorWinAreaFromParameter( doc, fi );
          }

          if( !openingArea.Equals( 0 ) )
          {
            LogCreator.LogEntry( ";_______OPENINGAREA;" 
              + elemInsert.Id.ToString() + ";" 
              + elemInsert.Category.Name + ";" 
              + elemInsert.Name + ";" 
              + ( openingArea * 0.09290304 ).ToString() );
          }
          return openingArea;
        }
      }

      if( elemInsert is Wall )
      {
        return solidHandler.GetWallAsOpeningArea( 
          elemInsert, roomSolid );
      }

      return openingArea;
    }

    public double GetWallCutArea( 
      Document doc, 
      FamilyInstance fi, 
      Wall wall, 
      bool isStacked )
    {
      Options optCompRef 
        = doc.Application.Create.NewGeometryOptions();

      if( null != optCompRef )
      {
        optCompRef.ComputeReferences = true;
        optCompRef.DetailLevel = ViewDetailLevel.Medium;
      }

      SolidHandler solHandler = new SolidHandler();
      XYZ cutDir = null;

      if( !isStacked )
      {
        CurveLoop curveLoop 
          = ExporterIFCUtils.GetInstanceCutoutFromWall( 
            fi.Document, wall, fi, out cutDir );

        IList<CurveLoop> loops = new List<CurveLoop>( 1 );
        loops.Add( curveLoop );

        return ExporterIFCUtils.ComputeAreaOfCurveLoops( 
          loops );
      }

      else if( isStacked )
      {
        CurveLoop curveLoop 
          = ExporterIFCUtils.GetInstanceCutoutFromWall( 
            fi.Document, wall, fi, out cutDir );

        IList<CurveLoop> loops = new List<CurveLoop>( 1 );
        loops.Add( curveLoop );

        GeometryElement geomElemHost = wall.get_Geometry( 
          optCompRef ) as GeometryElement;

        Solid solidOpening 
          = GeometryCreationUtilities.CreateExtrusionGeometry( 
            loops, cutDir.Negate(), .1 );

        Solid solidHost = solHandler.CreateSolidFromBoundingBox( 
          null, geomElemHost.GetBoundingBox(), null );

        if( solidHost == null )
        {
          return 0;
        }

        Solid intersectSolid 
          = BooleanOperationsUtils.ExecuteBooleanOperation( 
            solidOpening, solidHost, 
            BooleanOperationsType.Intersect );

        if( intersectSolid.Faces.Size.Equals( 0 ) )
        {
          solidOpening 
            = GeometryCreationUtilities.CreateExtrusionGeometry( 
              loops, cutDir, .1 );

          intersectSolid 
            = BooleanOperationsUtils.ExecuteBooleanOperation( 
              solidOpening, solidHost, 
              BooleanOperationsType.Intersect );
        }

        if( DebugHandler.EnableSolidUtilityVolumes )
        {
          using( Transaction trans = new Transaction( doc ) )
          {
            trans.Start( "stacked1" );
            ShapeCreator.CreateDirectShape( doc, 
              intersectSolid, "stackedOpening" );
            trans.Commit();
          }
        }
        return solHandler.GetLargestFaceArea( intersectSolid );
      }
      return 0;
    }

    private static bool IsToOrFromThisRoom( 
      Room room, 
      FamilyInstance fi )
    {
      bool isInRoom = false;

      if( fi.Room != null )
      {
        if( fi.Room.Id == room.Id )
        {
          return true;
        }
      }

      if( fi.ToRoom != null )
      {
        isInRoom = fi.ToRoom.Id.Equals( room.Id );
      }

      if( !isInRoom )
      {
        if( fi.FromRoom != null )
        {
          isInRoom = fi.FromRoom.Id.Equals( room.Id );
        }
      }
      return isInRoom;
    }

    private static double GetDoorWinAreaFromParameter( 
      Document doc, 
      FamilyInstance insert )
    {
      ElementType insertType = doc.GetElement( insert.GetTypeId() ) 
        as ElementType;

      double openingArea = 0;

      //check instance first then type
      Parameter height = insert.get_Parameter( BuiltInParameter.FAMILY_HEIGHT_PARAM );
      Parameter width = insert.get_Parameter( BuiltInParameter.FAMILY_WIDTH_PARAM );

      if( width != null && height != null )
      {
        openingArea = width.AsDouble() * height.AsDouble();
      }

      if( openingArea.Equals( 0 ) )
      {
        width = insertType.get_Parameter( BuiltInParameter.FAMILY_WIDTH_PARAM );
        height = insertType.get_Parameter( BuiltInParameter.FAMILY_HEIGHT_PARAM );
      }

      if( width != null && height != null )
      {
        openingArea = width.AsDouble() * height.AsDouble();
      }
      return openingArea;
    }

    private static Solid GetLargestSolid( GeometryElement geomElem )
    {
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
