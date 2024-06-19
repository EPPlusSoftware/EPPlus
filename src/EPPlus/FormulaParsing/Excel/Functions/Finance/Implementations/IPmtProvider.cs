/*************************************************************************************************
  * This Source Code Form is subject to the terms of the Mozilla Public
  * License, v. 2.0. If a copy of the MPL was not distributed with this
  * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/20/2020         EPPlus Software AB       Implemented function
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations
{
    /// <summary>
    /// IPmtProvider
    /// </summary>
    public interface IPmtProvider
    {
        /// <summary>
        /// GetPmt
        /// </summary>
        /// <param name="Rate"></param>
        /// <param name="NPer"></param>
        /// <param name="PV"></param>
        /// <param name="FV"></param>
        /// <param name="Due"></param>
        /// <returns></returns>
        double GetPmt(double Rate, double NPer, double PV, double FV = 0, PmtDue Due = PmtDue.EndOfPeriod);
    }
}
