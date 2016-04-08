using System;
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
      ElementId idMaterial
        = new FilteredElementCollector( doc )
         .OfClass( typeof( Material ) )
         .FirstOrDefault()
         .Id;

      //GraphicsStyle style
      //  = new FilteredElementCollector( doc )
      //    .OfClass( typeof( GraphicsStyle ) )
      //  .Cast<GraphicsStyle>()
      //  .FirstOrDefault<GraphicsStyle>(
      //    gs => gs.Name.Equals( "Walls" ) );

      //ElementId idGraphicsStyle = null;

      //if( style != null )
      //{
      //  idGraphicsStyle = style.Id;
      //}
      //else
      //{
      //  LogCreator.LogEntry( "Cant create DirectShape because the Graphic Style was not found." );
      //  return null;
      //}

      try
      {
        ElementId catId = new ElementId( BuiltInCategory.OST_GenericModel );
        DirectShape dsUtilityVolume = DirectShape.CreateElement( doc, catId, "06713861-8D80-4BCE-9B42-657695D45DC8", "" );

        bool isValid = dsUtilityVolume.IsValidGeometry( transientSolid );

        if( isValid )
        {
          dsUtilityVolume.SetShape( new GeometryObject[] { transientSolid } );
        }
        else
        {
          return null;
        }

        dsUtilityVolume.Name = dsName;

        return dsUtilityVolume;
      }
      catch
      {
        LogCreator.LogEntry( "DirectShape creation failed." );
        return null;
      }
    }

    //public static IList<GeometryObject> CreateMeshesFromSolid(Solid solid)
    //{
    //    IList<GeometryObject> triangulations = new List<GeometryObject>();

    //    foreach (Face face in solid.Faces)
    //    {
    //        Mesh faceMesh = face.Triangulate();
    //        if (faceMesh != null && faceMesh.NumTriangles > 0)
    //            triangulations.Add(faceMesh);
    //    }

    //    return triangulations;
    //}
  }
}
