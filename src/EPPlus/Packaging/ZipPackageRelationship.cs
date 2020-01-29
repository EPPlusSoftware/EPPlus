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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Packaging
{
    /// <summary>
    /// A relation ship between two parts in a package
    /// </summary>
    public class ZipPackageRelationship
    {
        /// <summary>
        /// The uri to the source part
        /// </summary>
        public Uri SourceUri { get; internal set; }
        /// <summary>
        /// The relationship type
        /// </summary>
        public string RelationshipType { get; internal set; }
        /// <summary>
        /// Target, internal or external
        /// </summary>
        public TargetMode TargetMode { get; internal set; }
        /// <summary>
        /// The relationship Id
        /// </summary>
        public string Id { get; internal set; }
        /// <summary>
        /// The uri to the target part
        /// </summary>
        public Uri TargetUri { get; set; }
        /// <summary>
        /// The target if it's not a valid uri, for example an internal reference to a cell withing the package.
        /// </summary>
        public string Target { get; internal set; }

    }
}
