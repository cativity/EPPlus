/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/

using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Security;

// General Information about an assembly is controlled through the following 
// set of attributes. Change these attribute values to modify the information
// associated with an assembly.
[assembly:
    InternalsVisibleTo("EPPlusTest, PublicKey=00240000048000009400000006020000002400005253413100040000010001001dd11308ec93a6ebcec727e183a8972dc6f95c23ecc34aa04f40cbfc9c17b08b4a0ea5c00dcd203bace44d15a30ce8796e38176ae88e960ceff9cc439ab938738ba0e603e3d155fc298799b391c004fc0eb4393dd254ce25db341eb43303e4c488c9500e126f1288594f0710ec7d642e9c72e76dd860649f1c48249c00e31fba")]
[assembly:
    InternalsVisibleTo("DynamicProxyGenAssembly2, PublicKey=0024000004800000940000000602000000240000525341310004000001000100c547cac37abd99c8db225ef2f6c8a3602f3b3606cc9891605d02baa56104f4cfc0734aa39b93bf7852f7d9266654753cc297e7d2edfe0bac1cdcf9f717241550e0a7b191195b7667bb4f64bcb8e2121380fd1d9d46ad2d92d2d15605093924cceaf74c4861eff62abf69b9291ed0a340e113be11e6a7d3113e92484cf7045cc7")]

// The following GUID is for the ID of the typelib if this project is exposed to COM
[assembly: Guid("9dd43b8d-c4fe-4a8b-ad6e-47ef83bbbb01")]

// Version information for an assembly consists of the following four values:
//
//      Major Version
//      Minor Version 
//      Build Number
//      Revision
//
// You can specify all the values or you can default the Revision and Build Numbers 
// by using the '*' as shown below:
#if (!Core)
    //[assembly: AssemblyTitle("EPPlus")]
    //[assembly: AssemblyDescription("A spreadsheet library for .NET framework and .NET core")]
    //[assembly: AssemblyConfiguration("")]
    //[assembly: AssemblyCompany("EPPlus Software AB")]
    //[assembly: AssemblyProduct("EPPlus")]
    //[assembly: AssemblyCopyright("EPPlus Software AB")]
    //[assembly: AssemblyTrademark("")]
    //[assembly: AssemblyCulture("")]
    //[assembly: ComVisible(false)]

    //[assembly: AssemblyVersion("5.5.0.0")]
    //[assembly: AssemblyFileVersion("5.5.0.0")]
#endif
[assembly: AllowPartiallyTrustedCallers]