using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

// General Information about an assembly is controlled through the following
// set of attributes. Change these attribute values to modify the information
// associated with an assembly.
[assembly: AssemblyTitle( "SpatialElementGeometryCalculator" )]
[assembly: AssemblyDescription( "Revit Add-In Description" )]
[assembly: AssemblyConfiguration( "" )]
[assembly: AssemblyCompany( "Autodesk Inc." )]
[assembly: AssemblyProduct( "SpatialElementGeometryCalculator" )]
[assembly: AssemblyCopyright( "Copyright 2015-2019 (C) Phillip Miller, Kiwi Codes Solutions Ltd., Jeremy Tammik, Autodesk Inc., Hakon Clausen, Håvard Dagsvik" )]
[assembly: AssemblyTrademark( "" )]
[assembly: AssemblyCulture( "" )]

// Setting ComVisible to false makes the types in this assembly not visible
// to COM components.  If you need to access a type in this assembly from
// COM, set the ComVisible attribute to true on that type.
[assembly: ComVisible( false )]

// The following GUID is for the ID of the typelib if this project is exposed to COM
[assembly: Guid( "321044f7-b0b2-4b1c-af18-e71a19252be0" )]

// Version information for an assembly consists of the following four values:
//
//      Major Version
//      Minor Version
//      Build Number
//      Revision
//
// You can specify all the values or you can default the Build and Revision Numbers
// by using the '*' as shown below:
// [assembly: AssemblyVersion("1.0.*")]
//
// 2015-03-17 2015.0.0.0 initial porting from VB.NET, cleaned up and enhanced the filtering 
// 2015-03-19 2015.0.0.1 use FindInserts instead of filtered element collector
// 2015-04-16 2015.0.0.3 reimplementation by Hakon Clausen: issue #1, pull request #2
// 2015-12-03 2015.0.0.4 fixed bug caused by element deletion within filtered element collector MoveNext /a/doc/revit/tbc/img/spatial_element_geomatry_calculator_error.png http://thebuildingcoder.typepad.com/blog/2015/03/findinserts-retrieves-all-openings-in-all-wall-types.html#comment-2380592080
// 2015-12-03 2015.0.0.5 cleanup before migrating to Revit 2016
// 2015-12-03 2016.0.0.0 flat migration to Revit 2016 - zero changes, zero errors, zero warnings
// 2016-04-08 2016.0.0.1 integrated added OpeningHandler.cs by Håvard Dagsvik
// 2016-04-08 2016.0.0.2 cleanup but not yet tested
// 2016-04-13 2016.0.0.3 integrated and executed testing code by Håvard Dagsvik
// 2019-05-07 2020.0.0.0 flat migration to Revit 2020
// 2019-05-07 2020.0.0.1 fixed architecture mismatch warning
// 2019-05-07 2020.0.0.2 applied fix to subtract opening area suggested by dan in comment https://thebuildingcoder.typepad.com/blog/2016/04/determining-wall-opening-areas-per-room.html#comment-4452689539
// 2019-05-08 2020.0.0.3 added GetOpenings as suggested by Håvard in comment https://thebuildingcoder.typepad.com/blog/2016/04/determining-wall-opening-areas-per-room.html#comment-4454441990
//
[assembly: AssemblyVersion( "2020.0.0.3" )]
[assembly: AssemblyFileVersion( "2020.0.0.3" )]
