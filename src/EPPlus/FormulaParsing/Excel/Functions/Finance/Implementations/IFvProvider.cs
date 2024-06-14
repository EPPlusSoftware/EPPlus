/*************************************************************************************************
  * This Source Code Form is subject to the terms of the Mozilla Public
  * License, v. 2.0. If a copy of the MPL was not distributed with this
  * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/20/2020         EPPlus Software AB       Implemented function
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance
{
    /// <summary>
    /// IFvProvider
    /// </summary>
    public interface IFvProvider
    {
        /// <summary>
        /// GetFv
        /// </summary>
        /// <param name="Rate"></param>
        /// <param name="NPer"></param>
        /// <param name="Pmt"></param>
        /// <param name="PV"></param>
        /// <param name="Due"></param>
        /// <returns></returns>
        double GetFv(double Rate, double NPer, double Pmt, double PV = 0, PmtDue Due = PmtDue.EndOfPeriod);
    }
}
