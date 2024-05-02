/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  //2024         EPPlus Software AB       Initial release EPPlus 
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing
{
    internal class ExcelEquation : ExcelShapeBase
    {

        /*NOTES:
         * 
         * Equations verkar börja på <mc:alternatecontent>
         * Har sedan en <mc:choice> som innehåller equation i <m>
         * På samma nivå finns en <mc:fallback> som innehåller ekvationen i <a>
         * 
         * I <mc:choice>
         *      <xdr:sp>
         *          <xdr:nvSpPr>
         *          <xdr:spPr>
         *          <xdr:style>
         *          <xdr:txbody>
         *              <a:p>
         *                  <a:pPr />
         *                  <a14:m>
         *                      <m:oMathPara>
         *                          <m:oMathParaPr>
         *                          <m:oMath> //THE FUN STARTS HERE
         *                              <m:>
         *                                  <m:Pr>
         *                                  <m:t>
         */
        internal ExcelEquation(ExcelDrawings drawings, XmlNode node, ExcelGroupShape shape = null) :
            base(drawings, node, "xdr:sp", "xdr:nvSpPr/xdr:cNvPr", shape)
        {
        }
    }
}
