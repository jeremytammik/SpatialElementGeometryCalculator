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
        ds.SetShape( new GeometryObject[] { transientSolid } );
        ds.Name = dsName;
      }
      else
      {
        ds = null;
      }
      return ds;
    }
  }
}
