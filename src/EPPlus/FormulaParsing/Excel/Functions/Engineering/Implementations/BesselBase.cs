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

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Engineering.Implementations
{
    public abstract class BesselBase
    {
        protected const double f_PI = 3.1415926535897932385;
        protected const double f_PI_DIV_2 = f_PI / 2.0;
        protected const double f_PI_DIV_4 = f_PI / 4.0;
        protected const double f_2_DIV_PI = 2.0 / f_PI;
    }
}
