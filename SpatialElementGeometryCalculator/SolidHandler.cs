using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace SpatialElementGeometryCalculator
{
  class SolidHandler
  {
    public double GetWallAsOpeningArea( 
      Element elemOpening, 
      Solid solidRoom )
    {
      Document doc = elemOpening.Document;

      Wall wallAsOpening = elemOpening as Wall;  // most likely an embedded curtain wall

      Options options = doc.Application.Create.NewGeometryOptions();
      options.ComputeReferences = true;
      options.IncludeNonVisibleObjects = true;

      List<Element> walls = new List<Element>();
      walls.Add( wallAsOpening );

      // To my recollection this won't 
      // pick up an edited wall profile

      List<List<XYZ>> polygons = GetWallProfilePolygons( 
        walls, options );

      IList<CurveLoop> solidProfile 
        = XYZAsCurveloop( polygons.First() );

      Solid solidOpening = GeometryCreationUtilities
        .CreateExtrusionGeometry( solidProfile, 
          wallAsOpening.Orientation, 1 );

      Solid intersectSolid = BooleanOperationsUtils
        .ExecuteBooleanOperation( solidOpening, 
          solidRoom, BooleanOperationsType.Intersect );

      if( intersectSolid.Faces.Size.Equals( 0 ) )
      {
        // Then we are extruding in the wrong direction

        solidOpening = GeometryCreationUtilities
          .CreateExtrusionGeometry( solidProfile, 
            wallAsOpening.Orientation.Negate(), 1 );

        intersectSolid = BooleanOperationsUtils
          .ExecuteBooleanOperation( solidOpening, 
            solidRoom, BooleanOperationsType.Intersect );
      }

      if( DebugHandler.EnableSolidUtilityVolumes )
      {
        using( Transaction t = new Transaction( doc ) )
        {
          t.Start( "Test1" );
          ShapeCreator.CreateDirectShape( doc, 
            intersectSolid, "namesolid" );
          t.Commit();
        }
      }

      double openingArea = GetLargestFaceArea( 
        intersectSolid );

      LogCreator.LogEntry( ";_______OPENINGAREA;" 
        + elemOpening.Id.ToString() + ";" 
        + elemOpening.Category.Name + ";" 
        + elemOpening.Name + ";"
        + ( openingArea * 0.09290304 ).ToString() );

      return openingArea;
    }

    public IList<CurveLoop> XYZAsCurveloop( 
      List<XYZ> polyPoints )
    {
      CurveLoop curveLoop = new CurveLoop();

      for( int i = 0; i < polyPoints.Count - 1; i++ )
      {
        curveLoop.Append( Line.CreateBound( 
          polyPoints[i], polyPoints[i + 1] ) );
      }

      curveLoop.Append( Line.CreateBound( 
        polyPoints[polyPoints.Count - 1], polyPoints[0] ) );

      IList<CurveLoop> curveLoops = new List<CurveLoop>();
      curveLoops.Add( curveLoop );

      return curveLoops;
    }

    static public List<List<XYZ>> GetWallProfilePolygons( 
      List<Element> walls, 
      Options opt )
    {
      XYZ p, q, v, w;

      List<List<XYZ>> polygons = new List<List<XYZ>>();

      foreach( Wall wall in walls )
      {
        LocationCurve curve
          = wall.Location as LocationCurve;

        if( null == curve )
        {
          return null;
        }
        p = curve.Curve.GetEndPoint( 0 );
        q = curve.Curve.GetEndPoint( 1 );
        v = q - p;
        v = v.Normalize();
        w = XYZ.BasisZ.CrossProduct( v ).Normalize();

        if( wall.Flipped ) { w = -w; }

        GeometryElement geo = wall.get_Geometry( opt );

        foreach( GeometryObject obj in geo )
        {
          Solid solid = obj as Solid;
          if( solid != null )
          {
            GetProfile( polygons, solid, v, w );
          }
        }
      }
      return polygons;
    }

    const double _offset = 0;

    private static bool GetProfile( 
      List<List<XYZ>> polygons, 
      Solid solid, 
      XYZ v, 
      XYZ w )
    {
      double d, dmax = 0;
      PlanarFace outermost = null;
      FaceArray faces = solid.Faces;
      foreach( Face f in faces )
      {
        PlanarFace pf = f as PlanarFace;
        if( null != pf && Util.IsVertical( pf ) 
          && Util.IsZero( v.DotProduct( pf.FaceNormal ) ) )
        {
          d = pf.Origin.DotProduct( w );
          if( ( null == outermost ) || ( dmax < d ) )
          {
            outermost = pf;
            dmax = d;
          }
        }
      }

      if( null != outermost )
      {
        XYZ voffset = _offset * w;
        XYZ p, q = XYZ.Zero;
        bool first;
        int i, n;
        EdgeArrayArray loops = outermost.EdgeLoops;
        foreach( EdgeArray loop in loops )
        {
          List<XYZ> vertices = new List<XYZ>();
          first = true;
          foreach( Edge e in loop )
          {
            IList<XYZ> points = e.Tessellate();
            p = points[0];
            if( !first )
            {
              if( !p.IsAlmostEqualTo( q ) )
              {
                LogCreator.LogEntry( "Expected "
                  + "subsequent start point to equal "
                  + "previous end point" );
              }
            }

            n = points.Count;
            q = points[n - 1];
            for( i = 0; i < n - 1; ++i )
            {
              XYZ a = points[i];
              a += voffset;
              vertices.Add( a );
            }
          }

          q += voffset;

          if( !q.IsAlmostEqualTo( vertices[0] ) )
          {
            LogCreator.LogEntry( "Expected last end "
              + "point to equal first start point" );
          }

          polygons.Add( vertices );
        }
      }

      return null != outermost;
    }

    public double GetLargestFaceArea( 
      Solid intersectSolid )
    {
      FaceArray faceArray = intersectSolid.Faces;

      //Face targetFace = null;
      double maxFaceArea = 0;

      foreach( Face face in faceArray )
      {
        double a = face.Area;

        if( a > maxFaceArea )
        {
          //targetFace = face;
          maxFaceArea = a;
        }
      }
      return maxFaceArea;
    }

    public Solid CreateSolidFromBoundingBox( 
      Transform lcs, 
      BoundingBoxXYZ boundingBoxXYZ, 
      SolidOptions solidOptions )
    {
      // Check that the bounding box is valid.

      if( boundingBoxXYZ == null
        || !boundingBoxXYZ.Enabled )
      {
        return null;
      }

      try
      {
        // Create a transform based on the incoming 
        // local coordinate system and the bounding 
        // box coordinate system.

        Transform bboxTransform = ( lcs == null ) 
          ? boundingBoxXYZ.Transform 
          : lcs.Multiply( boundingBoxXYZ.Transform );

        XYZ[] profilePts = new XYZ[4];
        profilePts[0] = bboxTransform.OfPoint( boundingBoxXYZ.Min );
        profilePts[1] = bboxTransform.OfPoint( new XYZ( boundingBoxXYZ.Max.X, boundingBoxXYZ.Min.Y, boundingBoxXYZ.Min.Z ) );
        profilePts[2] = bboxTransform.OfPoint( new XYZ( boundingBoxXYZ.Max.X, boundingBoxXYZ.Max.Y, boundingBoxXYZ.Min.Z ) );
        profilePts[3] = bboxTransform.OfPoint( new XYZ( boundingBoxXYZ.Min.X, boundingBoxXYZ.Max.Y, boundingBoxXYZ.Min.Z ) );

        XYZ upperRightXYZ = bboxTransform.OfPoint( boundingBoxXYZ.Max );

        // If we assumed that the transforms had no scaling, 
        // then we could simply take boundingBoxXYZ.Max.Z - boundingBoxXYZ.Min.Z.
        // This code removes that assumption.

        XYZ origExtrusionVector = new XYZ( 
          boundingBoxXYZ.Min.X, boundingBoxXYZ.Min.Y, 
          boundingBoxXYZ.Max.Z ) - boundingBoxXYZ.Min;

        XYZ extrusionVector = bboxTransform.OfVector( 
          origExtrusionVector );

        double extrusionDistance = extrusionVector.GetLength();
        XYZ extrusionDirection = extrusionVector.Normalize();

        CurveLoop baseLoop = new CurveLoop();

        for( int i = 0; i < 4; i++ )
        {
          baseLoop.Append( Line.CreateBound( 
            profilePts[i], profilePts[( i + 1 ) % 4] ) );
        }

        IList<CurveLoop> baseLoops = new List<CurveLoop>();
        baseLoops.Add( baseLoop );

        if( solidOptions == null )
          return GeometryCreationUtilities
            .CreateExtrusionGeometry( baseLoops, 
              extrusionDirection, extrusionDistance );
        else
          return GeometryCreationUtilities
            .CreateExtrusionGeometry( baseLoops, 
              extrusionDirection, extrusionDistance, 
              solidOptions );
      }
      catch
      {
        return null;
      }
    }
  }
}
