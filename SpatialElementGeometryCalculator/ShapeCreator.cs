using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace SpatialElementGeometryCalculator
{
  public class ShapeCreator
  {
    public static DirectShape CreateDirectShape(
      Document doc,
      Solid transientSolid,
      string dsName )
    {
      ElementId catId = new ElementId(
      BuiltInCategory.OST_GenericModel );

      AddInId addInId = doc.Application.ActiveAddInId;

      DirectShape ds
        = DirectShape.CreateElement( doc, catId,
          addInId.GetGUID().ToString(), "" );

      if( ds.IsValidGeometry( transientSolid ) )
      {
        ds.SetShape( new GeometryObject[] { 
          transientSolid } );
      }
      else
      {
        TessellatedShapeBuilderResult result 
          = GetTessellatedSolid( doc, transientSolid );

        ds.SetShape( result.GetGeometricalObjects() );
      }

      ds.Name = dsName;

      return ds;
    }

    static TessellatedShapeBuilderResult GetTessellatedSolid( 
      Document doc, 
      Solid transientSolid )
    {
      TessellatedShapeBuilder builder 
        = new TessellatedShapeBuilder();

      ElementId idMaterial
        = new FilteredElementCollector( doc )
          .OfClass( typeof( Material ) )
          .FirstElementId();

      ElementId idGraphicsStyle
        = new FilteredElementCollector( doc )
          .OfClass( typeof( GraphicsStyle ) )
          .FirstOrDefault<Element>( gs 
            => gs.Name.Equals( "Walls" ) )
          .Id;

      builder.OpenConnectedFaceSet( true );

      FaceArray faceArray = transientSolid.Faces;

      foreach( Face face in faceArray )
      {
        List<XYZ> triFace = new List<XYZ>( 3 );
        Mesh mesh = face.Triangulate();

        int triCount = mesh.NumTriangles;

        for( int i = 0; i < triCount; i++ )
        {
          triFace.Clear();

          for( int n = 0; n < 3; n++ )
          {
            triFace.Add( mesh.get_Triangle( i ).get_Vertex( n ) );
          }

          builder.AddFace( new TessellatedFace( 
            triFace, idMaterial ) );
        }
      }

      builder.CloseConnectedFaceSet();

      return builder.Build( 
        TessellatedShapeBuilderTarget.Solid, 
        TessellatedShapeBuilderFallback.Abort, 
        idGraphicsStyle );
    }
  }
}
