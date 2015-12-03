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
[assembly: AssemblyCopyright( "Copyright 2015 (C) Phillip Miller, Kiwi Codes Solutions Ltd., Jeremy Tammik, Autodesk Inc., Hakon Clausen" )]
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
//
[assembly: AssemblyVersion( "2015.0.0.4" )]
[assembly: AssemblyFileVersion( "2015.0.0.4" )]
