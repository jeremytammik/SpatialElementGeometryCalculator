using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
using Autodesk.Revit.DB;
//using Autodesk.Revit.DB.Architecture;

namespace SpatialElementGeometryCalculator
{
  class SpatialBoundaryCache
  {
    public string roomName;
    public ElementId idElement;
    public ElementId idMaterial;
    public double dblGrossArea;
    public double dblOpeningArea;
  }
}
