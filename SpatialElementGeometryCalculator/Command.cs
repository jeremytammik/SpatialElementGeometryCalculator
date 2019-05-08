#region Namespaces
using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System.Text;
#endregion

namespace SpatialElementGeometryCalculator
{
  [Transaction( TransactionMode.Manual )]
  public class Command : IExternalCommand
  {
    public Result Execute(
      ExternalCommandData commandData,
      ref string message,
      ElementSet elements )
    {
      UIApplication uiapp = commandData.Application;
      Document doc = uiapp.ActiveUIDocument.Document;
      Result rc;

      try
      {
        SpatialElementBoundaryOptions sebOptions
          = new SpatialElementBoundaryOptions
          {
            SpatialElementBoundaryLocation
              = SpatialElementBoundaryLocation.Finish
          };

        IEnumerable<Element> rooms
          = new FilteredElementCollector( doc )
            .OfClass( typeof( SpatialElement ) )
            .Where<Element>( e => ( e is Room ) );

        List<string> compareWallAndRoom = new List<string>();
        OpeningHandler openingHandler = new OpeningHandler();

        List<SpatialBoundaryCache> lstSpatialBoundaryCache
          = new List<SpatialBoundaryCache>();

        foreach( Room room in rooms )
        {
          if( room == null ) continue;
          if( room.Location == null ) continue;
          if( room.Area.Equals( 0 ) ) continue;

          Autodesk.Revit.DB.SpatialElementGeometryCalculator calc =
            new Autodesk.Revit.DB.SpatialElementGeometryCalculator(
              doc, sebOptions );

          SpatialElementGeometryResults results
            = calc.CalculateSpatialElementGeometry(
              room );

          Solid roomSolid = results.GetGeometry();

          foreach( Face face in results.GetGeometry().Faces )
          {
            IList<SpatialElementBoundarySubface> boundaryFaceInfo
              = results.GetBoundaryFaceInfo( face );

            foreach( var spatialSubFace in boundaryFaceInfo )
            {
              if( spatialSubFace.SubfaceType
                != SubfaceType.Side )
              {
                continue;
              }

              SpatialBoundaryCache spatialData
                = new SpatialBoundaryCache();

              Wall wall = doc.GetElement( spatialSubFace
                .SpatialBoundaryElement.HostElementId )
                  as Wall;

              if( wall == null )
              {
                continue;
              }

              WallType wallType = doc.GetElement(
                wall.GetTypeId() ) as WallType;

              if( wallType.Kind == WallKind.Curtain )
              {
                // Leave out, as curtain walls are not painted.

                LogCreator.LogEntry( "WallType is CurtainWall" );

                continue;
              }

              HostObject hostObject = wall as HostObject;

              IList<ElementId> insertsThisHost
                = hostObject.FindInserts(
                  true, false, true, true );

              double openingArea = 0;

              foreach( ElementId idInsert in insertsThisHost )
              {
                string countOnce = room.Id.ToString()
                  + wall.Id.ToString() + idInsert.ToString();

                if( !compareWallAndRoom.Contains( countOnce ) )
                {
                  Element elemOpening = doc.GetElement(
                    idInsert ) as Element;

                  openingArea = openingArea
                    + openingHandler.GetOpeningArea(
                      wall, elemOpening, room, roomSolid );

                  compareWallAndRoom.Add( countOnce );
                }
              }

              // Cache SpatialElementBoundarySubface info.

              spatialData.roomName = room.Name;
              spatialData.idElement = wall.Id;
              spatialData.idMaterial = spatialSubFace
                .GetBoundingElementFace().MaterialElementId;
              spatialData.dblNetArea = Util.sqFootToSquareM(
                spatialSubFace.GetSubface().Area - openingArea ); 
              spatialData.dblOpeningArea = Util.sqFootToSquareM(
                openingArea );

              lstSpatialBoundaryCache.Add( spatialData );

            } // end foreach subface from which room bounding elements are derived

          } // end foreach Face

        } // end foreach Room

        List<string> t = new List<string>();

        List<SpatialBoundaryCache> groupedData
          = SortByRoom( lstSpatialBoundaryCache );

        foreach( SpatialBoundaryCache sbc in groupedData )
        {
          t.Add( sbc.roomName
            + "; all wall types and materials: "
            + sbc.AreaReport );
        }

        Util.InfoMsg2( "Total Net Area in m2 by Room",
          string.Join( System.Environment.NewLine, t ) );

        t.Clear();

        groupedData = SortByRoomAndWallType(
          lstSpatialBoundaryCache );

        foreach( SpatialBoundaryCache sbc in groupedData )
        {
          Element elemWall = doc.GetElement(
            sbc.idElement ) as Element;

          t.Add( sbc.roomName + "; " + elemWall.Name
            + "(" + sbc.idElement.ToString() + "): "
            + sbc.AreaReport );
        }

        Util.InfoMsg2( "Net Area in m2 by Wall Type",
          string.Join( System.Environment.NewLine, t ) );

        t.Clear();

        groupedData = SortByRoomAndMaterial(
          lstSpatialBoundaryCache );

        foreach( SpatialBoundaryCache sbc in groupedData )
        {
          string materialName
            = ( sbc.idMaterial == ElementId.InvalidElementId )
              ? string.Empty
              : doc.GetElement( sbc.idMaterial ).Name;

          t.Add( sbc.roomName + "; " + materialName + ": "
            + sbc.AreaReport );
        }

        Util.InfoMsg2(
          "Net Area in m2 by Outer Layer Material",
          string.Join( System.Environment.NewLine, t ) );

        rc = Result.Succeeded;
      }
      catch( Exception ex )
      {
        TaskDialog.Show( "Room Boundaries",
          ex.Message + "\r\n" + ex.StackTrace );

        rc = Result.Failed;
      }
      return rc;
    }

    /// <summary>
    /// Convert square feet to square meters
    /// with two decimal places precision.
    /// </summary>
    static double SqFootToSquareM( double sqFoot )
    {
      return Math.Round( sqFoot * 0.092903, 2 );
    }

    List<SpatialBoundaryCache> SortByRoom(
      List<SpatialBoundaryCache> lstRawData )
    {
      var sortedCache
        = from rawData in lstRawData
          group rawData by new { room = rawData.roomName }
            into sortedData
            select new SpatialBoundaryCache()
            {
              roomName = sortedData.Key.room,
              idElement = ElementId.InvalidElementId,
              dblNetArea = sortedData.Sum( x => x.dblNetArea ),
              dblOpeningArea = sortedData.Sum(
                y => y.dblOpeningArea ),
            };

      return sortedCache.ToList();
    }

    List<SpatialBoundaryCache> SortByRoomAndWallType(
      List<SpatialBoundaryCache> lstRawData )
    {
      var sortedCache
        = from rawData in lstRawData
          group rawData by new
          {
            room = rawData.roomName,
            wallid = rawData.idElement
          }
            into sortedData
            select new SpatialBoundaryCache()
            {
              roomName = sortedData.Key.room,
              idElement = sortedData.Key.wallid,
              dblNetArea = sortedData.Sum( x => x.dblNetArea ),
              dblOpeningArea = sortedData.Sum(
                y => y.dblOpeningArea ),
            };

      return sortedCache.ToList();
    }

    List<SpatialBoundaryCache> SortByRoomAndMaterial(
      List<SpatialBoundaryCache> lstRawData )
    {
      var sortedCache
        = from rawData in lstRawData
          group rawData by new
          {
            room = rawData.roomName,
            mid = rawData.idMaterial
          }
            into sortedData
            select new SpatialBoundaryCache()
            {
              roomName = sortedData.Key.room,
              idMaterial = sortedData.Key.mid,
              dblNetArea = sortedData.Sum( x => x.dblNetArea ),
              dblOpeningArea = sortedData.Sum(
                y => y.dblOpeningArea ),
            };

      return sortedCache.ToList();
    }

    /// <summary>
    /// Return wall openings using GetDependentElements
    /// </summary>
    static IList<ElementId> GetOpenings( Wall wall )
    {
      ElementMulticategoryFilter emcf
        = new ElementMulticategoryFilter(
          new List<ElementId>() {
            new ElementId(BuiltInCategory.OST_Windows),
            new ElementId(BuiltInCategory.OST_Doors) } );

      return wall.GetDependentElements( emcf );
    }
  }
}
