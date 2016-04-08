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

      Wall wallAsOpening = elemOpening as Wall;  // most likely an embedded curtainwall acting as window

      Options options = wallAsOpening.Document.Application.Create.NewGeometryOptions();
      options.ComputeReferences = true;
      options.IncludeNonVisibleObjects = true;

      List<Element> walls = new List<Element>();
      walls.Add( wallAsOpening );

      List<List<XYZ>> polygons = GetWallProfilePolygons( walls, options );
      IList<CurveLoop> solidProfile = XYZAsCurveloop( polygons.First() );
      Solid solidOpening = GeometryCreationUtilities.CreateExtrusionGeometry( solidProfile, wallAsOpening.Orientation, 1 );
      Solid intersectSolid = BooleanOperationsUtils.ExecuteBooleanOperation( solidOpening, solidRoom, BooleanOperationsType.Intersect );

      if( intersectSolid.Faces.Size.Equals( 0 ) )
      {
        solidOpening = GeometryCreationUtilities.CreateExtrusionGeometry( solidProfile, wallAsOpening.Orientation.Negate(), 1 );
        intersectSolid = BooleanOperationsUtils.ExecuteBooleanOperation( solidOpening, solidRoom, BooleanOperationsType.Intersect );
      }

      if( DebugHandler.EnableSolidUtilityVolumes )
      {
        Transaction trans = new Transaction( doc, "test" );
        trans.Start( "test1" );
        ShapeCreator.CreateDirectShape( doc, intersectSolid, "namesolid" );

        Creator creator = new Creator( wallAsOpening.Document );
        creator.DrawPolygon( polygons.First() );
        wallAsOpening.Document.Regenerate();

        trans.Commit();
      }

      double openingArea = GetLargestFaceArea( intersectSolid );

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

      for( int ii = 0; ii < polyPoints.Count - 1; ii++ )
      {
        curveLoop.Append( Line.CreateBound( polyPoints[ii], polyPoints[ii + 1] ) );
      }

      curveLoop.Append( Line.CreateBound( polyPoints[polyPoints.Count - 1], polyPoints[0] ) );

      IList<CurveLoop> curveLoops = new List<CurveLoop>();
      curveLoops.Add( curveLoop );

      return curveLoops;
    }

    static public List<List<XYZ>> GetWallProfilePolygons( List<Element> walls, Options opt )
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

    private static bool GetProfile( List<List<XYZ>> polygons, Solid solid, XYZ v, XYZ w )
    {
      double d, dmax = 0;
      PlanarFace outermost = null;
      FaceArray faces = solid.Faces;
      foreach( Face f in faces )
      {
        PlanarFace pf = f as PlanarFace;
        if( null != pf && Util.IsVertical( pf ) && Util.IsZero( v.DotProduct( pf.FaceNormal ) ) )
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
                LogCreator.LogEntry( "Expected subsequent start point to equal previous end point" );
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
            LogCreator.LogEntry( "Expected last end point to equal first start point" );
          }

          polygons.Add( vertices );
        }
      }
      return null != outermost;
    }

    private double GetRadianFromPoint( XYZ point )
    {
      double dy = ( point.Y );
      double dx = ( point.X );

      double radian = Math.Atan2( dy, dx );

      double angle = (/*90 - */( ( radian * 180 ) / Math.PI ) ) % 360;
      return radian;
    }

    private double RadianToDegree( double rad )
    {
      return rad * ( 180.0 / Math.PI );
    }

    public double GetLargestFaceArea( Solid intersectSolid )
    {
      FaceArray faceArray = intersectSolid.Faces;

      Face targetFace = null;
      double dblFaceMaxSoFar = 0;

      foreach( Face face in faceArray )
      {
        double dblAreaThisFace = face.Area;

        if( dblAreaThisFace >= dblFaceMaxSoFar )
        {
          targetFace = face;
          dblFaceMaxSoFar = dblAreaThisFace;
        }
      }
      return dblFaceMaxSoFar;
    }

    //public Solid GetBoundingBoxSolid(Application app, Element opening)
    //{
    //    Autodesk.Revit.DB.Options optCompRef = app.Create.NewGeometryOptions();

    //    if (null != optCompRef)
    //    {
    //        optCompRef.ComputeReferences = true;
    //        optCompRef.DetailLevel = ViewDetailLevel.Medium;
    //    }

    //    GeometryElement geomElemBuilding = opening.get_Geometry(optCompRef) as GeometryElement;
    //    BoundingBoxXYZ boundingBox = BoundingBoxInModelCoordinate(geomElemBuilding.GetBoundingBox());

    //    Solid solidOpening = CreateTransientSolid(boundingBox);
    //    return solidOpening;
    //}

    //public Solid GetBoundingBoxSolid(Application app, Element opening,Transform transform)
    //{
    //    Autodesk.Revit.DB.Options optCompRef = app.Create.NewGeometryOptions();

    //    if (null != optCompRef)
    //    {
    //        optCompRef.ComputeReferences = true;
    //        optCompRef.DetailLevel = ViewDetailLevel.Medium;
    //    }

    //    GeometryElement geomElemBuilding = opening.get_Geometry(optCompRef) as GeometryElement;
    //    BoundingBoxXYZ boundingBox = BoundingBoxInModelCoordinate(geomElemBuilding.GetBoundingBox());

    //    Transaction trans = new Transaction(opening.Document, "test");

    //    trans.Start("start");
    //    Creator creator = new Creator(opening.Document);
    //    List<XYZ> polygon1 = Util.GetCorners(boundingBox).ToList();

    //    creator.DrawPolygon(polygon1);

    //    BoundingBoxXYZ boundingBoxRotated = RotateBoundingBox(boundingBox, transform);

    //    List<XYZ> polygon2 = Util.GetCorners(boundingBoxRotated).ToList();

    //    creator.DrawPolygon(polygon2);

    //    opening.Document.Regenerate();

    //    trans.Commit();



    //    Solid solidOpening = CreateTransientSolid(boundingBoxRotated);

    //    return solidOpening;
    //}

    public Solid CreateSolidFromBoundingBox( Transform lcs, BoundingBoxXYZ boundingBoxXYZ, SolidOptions solidOptions )
    {
      // Check that the bounding box is valid.
      if( boundingBoxXYZ == null || !boundingBoxXYZ.Enabled )
        return null;

      try
      {
        // Create a transform based on the incoming local coordinate system and the bounding box coordinate system.
        Transform bboxTransform = ( lcs == null ) ? boundingBoxXYZ.Transform : lcs.Multiply( boundingBoxXYZ.Transform );

        XYZ[] profilePts = new XYZ[4];
        profilePts[0] = bboxTransform.OfPoint( boundingBoxXYZ.Min );
        profilePts[1] = bboxTransform.OfPoint( new XYZ( boundingBoxXYZ.Max.X, boundingBoxXYZ.Min.Y, boundingBoxXYZ.Min.Z ) );
        profilePts[2] = bboxTransform.OfPoint( new XYZ( boundingBoxXYZ.Max.X, boundingBoxXYZ.Max.Y, boundingBoxXYZ.Min.Z ) );
        profilePts[3] = bboxTransform.OfPoint( new XYZ( boundingBoxXYZ.Min.X, boundingBoxXYZ.Max.Y, boundingBoxXYZ.Min.Z ) );

        XYZ upperRightXYZ = bboxTransform.OfPoint( boundingBoxXYZ.Max );

        // If we assumed that the transforms had no scaling, 
        // then we could simply take boundingBoxXYZ.Max.Z - boundingBoxXYZ.Min.Z.
        // This code removes that assumption.
        XYZ origExtrusionVector = new XYZ( boundingBoxXYZ.Min.X, boundingBoxXYZ.Min.Y, boundingBoxXYZ.Max.Z ) - boundingBoxXYZ.Min;
        XYZ extrusionVector = bboxTransform.OfVector( origExtrusionVector );

        double extrusionDistance = extrusionVector.GetLength();
        XYZ extrusionDirection = extrusionVector.Normalize();

        CurveLoop baseLoop = new CurveLoop();

        for( int ii = 0; ii < 4; ii++ )
        {
          baseLoop.Append( Line.CreateBound( profilePts[ii], profilePts[( ii + 1 ) % 4] ) );
        }

        IList<CurveLoop> baseLoops = new List<CurveLoop>();
        baseLoops.Add( baseLoop );

        if( solidOptions == null )
          return GeometryCreationUtilities.CreateExtrusionGeometry( baseLoops, extrusionDirection, extrusionDistance );
        else
          return GeometryCreationUtilities.CreateExtrusionGeometry( baseLoops, extrusionDirection, extrusionDistance, solidOptions );
      }
      catch
      {
        return null;
      }
    }

    public static BoundingBoxXYZ BoundingBoxInModelCoordinate( BoundingBoxXYZ bbox )
    {
      if( bbox == null )
        return null;

      double[] xVals = new double[] { bbox.Min.X, bbox.Max.X };
      double[] yVals = new double[] { bbox.Min.Y, bbox.Max.Y };
      double[] zVals = new double[] { bbox.Min.Z, bbox.Max.Z };

      XYZ toTest;

      double minX, minY, minZ, maxX, maxY, maxZ;
      minX = minY = minZ = double.MaxValue;
      maxX = maxY = maxZ = double.MinValue;

      // Get the max and min coordinate from the 8 vertices
      for( int iX = 0; iX < 2; iX++ )
      {
        for( int iY = 0; iY < 2; iY++ )
        {
          for( int iZ = 0; iZ < 2; iZ++ )
          {
            toTest = bbox.Transform.OfPoint( new XYZ( xVals[iX], yVals[iY], zVals[iZ] ) );
            minX = Math.Min( minX, toTest.X );
            minY = Math.Min( minY, toTest.Y );
            minZ = Math.Min( minZ, toTest.Z );

            maxX = Math.Max( maxX, toTest.X );
            maxY = Math.Max( maxY, toTest.Y );
            maxZ = Math.Max( maxZ, toTest.Z );
          }
        }
      }

      BoundingBoxXYZ returnBox = new BoundingBoxXYZ();
      returnBox.Max = new XYZ( maxX, maxY, maxZ );
      returnBox.Min = new XYZ( minX, minY, minZ );

      return returnBox;
    }

    private static BoundingBoxXYZ RotateBoundingBox( BoundingBoxXYZ box, Transform transform )
    {
      double height = box.Max.Z - box.Min.Z;

      // Four corners: lower left, lower right, 
      // upper right, upper left:

      XYZ[] corners = Util.GetBottomCorners( box );

      XyzComparable[] cornersTransformed
        = corners.Select<XYZ, XyzComparable>(
          p => new XyzComparable( transform.OfPoint( p ) ) )
            .ToArray();

      box.Min = cornersTransformed.Min();
      box.Max = cornersTransformed.Max();
      box.Max += height * XYZ.BasisZ;

      return box;
    }

    class XyzComparable : XYZ, IComparable<XYZ>
    {
      public XyzComparable( XYZ a )
        : base( a.X, a.Y, a.Z )
      {
      }

      int IComparable<XYZ>.CompareTo( XYZ a )
      {
        return Util.Compare( this, a );
      }
    }

    private static Solid CreateTransientSolid( 
      BoundingBoxXYZ bb )
    {
      XYZ p1 = bb.Min;
      XYZ p2 = new XYZ( bb.Min.X, bb.Max.Y, bb.Min.Z );
      XYZ p3 = new XYZ( bb.Max.X, bb.Max.Y, bb.Min.Z );
      XYZ p4 = new XYZ( bb.Max.X, bb.Min.Y, bb.Min.Z );
      double height = bb.Max.Z - bb.Min.Z;

      CurveLoop rectangleLoop = new CurveLoop();
      rectangleLoop.Append( Line.CreateBound( p1, p2 ) );
      rectangleLoop.Append( Line.CreateBound( p2, p3 ) );
      rectangleLoop.Append( Line.CreateBound( p3, p4 ) );
      rectangleLoop.Append( Line.CreateBound( p4, p1 ) );

      List<CurveLoop> curveloops = new List<CurveLoop>();
      curveloops.Add( rectangleLoop );

      return GeometryCreationUtilities.CreateExtrusionGeometry( 
        curveloops, XYZ.BasisZ, height );
    }
  }
}
