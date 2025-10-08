/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/07/2025         EPPlus Software AB       Initial release EPPlus 8.3
 *************************************************************************************************/
using System.Collections;
using System.Collections.Generic;

namespace OfficeOpenXml.Connection
{
    public class ExcelConnectionCollection : IEnumerable<ExcelConnection>
    {
        List<ExcelConnection> _connection = new List<ExcelConnection>();
        public IEnumerator<ExcelConnection> GetEnumerator()
        {
            return _connection.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _connection.GetEnumerator();
        }

    }
}
