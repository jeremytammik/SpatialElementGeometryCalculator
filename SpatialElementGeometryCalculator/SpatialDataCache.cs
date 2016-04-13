using System;
using Autodesk.Revit.DB;

namespace SpatialElementGeometryCalculator
{
  class SpatialBoundaryCache
  {
    public string roomName;
    public ElementId idElement;
    public ElementId idMaterial;
    public double dblNetArea;
    public double dblOpeningArea;

    public string AreaReport
    {
      get
      {
        return string.Format(
          "net {0}; opening {1}; gross {2}",
          dblNetArea, dblOpeningArea, 
          dblNetArea + dblOpeningArea );
      }
    }
  }
}
