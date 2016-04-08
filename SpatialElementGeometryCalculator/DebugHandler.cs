using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpatialElementGeometryCalculator
{
  public class DebugHandler
  {
    static bool _enableSolidUtilityVolumes = false;

    public static bool EnableSolidUtilityVolumes
    {
      get
      {
        return _enableSolidUtilityVolumes;
      }
      set
      {
        _enableSolidUtilityVolumes = value;
      }
    }
  }
}
