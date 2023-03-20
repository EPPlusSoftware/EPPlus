using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis.TokenSeparatorHandlers
{
    internal class ExponentialNumberHandler : SeparatorHandler
    {
        public override bool Handle(char c, Token tokenSeparator, TokenizerContext context, ITokenIndexProvider tokenIndexProvider)
        {
            if(!string.IsNullOrEmpty(context.CurrentToken))
            {
                if (c == '-')
                {
                    var currentToken = context.CurrentToken;
                    var arr = currentToken.ToArray();
                    if (arr[arr.Length - 1] != 'E') return false;
                    for (var x = 0; x < arr.Length - 1; x++)
                    {
                        var ch = arr[x];
                        if (char.IsDigit(ch) || ch == '.')
                        {
                            continue;
                        }
                        else
                        {
                            return false;
                        }
                    }
                    context.AppendToCurrentToken('-');
                    return true;
                }
                else if(c == '+')
                {
                    var currentToken = context.CurrentToken;
                    var arr = currentToken.ToArray();
                    if (arr[arr.Length -1] != 'E') return false;
                    
                    for (var x = 0; x < arr.Length - 1; x++)
                    {
                        var ch = arr[x];
                        if (char.IsDigit(ch) || ch == '.')
                        {
                            continue;
                        }
                        else
                        {
                            return false;
                        }
                    }
                    context.AppendToCurrentToken('+');
                    return true;
                }
            }
            
            return false;
        }
    }
}
