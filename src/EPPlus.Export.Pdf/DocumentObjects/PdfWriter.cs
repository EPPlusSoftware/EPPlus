/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  27/11/2025         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/
using EPPlus.Export.Pdf.Settings;
using System.IO;
using System.Text;

internal sealed class PdfWriter
{
    private readonly BinaryWriter _bw;
    private readonly StringBuilder _debug;
    internal readonly bool IsDebug;
    private byte[] _buf = new byte[512];
    internal readonly PdfCompressionLevel CompressionLevel;

    public PdfWriter(BinaryWriter bw, bool captureDebug = false, PdfCompressionLevel compressionLevel = PdfCompressionLevel.Smallest)
    {
        _bw = bw;
        IsDebug = captureDebug;
        CompressionLevel = compressionLevel;
        if (captureDebug) _debug = new StringBuilder();
    }

    public string DebugText => _debug?.ToString();

    public void Text(string s)
    {
        if (_bw != null)
        {
            int n = s.Length;
            if (_buf.Length < n) _buf = new byte[n];
            Encoding.ASCII.GetBytes(s, 0, n, _buf, 0);
            _bw.Write(_buf, 0, n);
        }
        _debug?.Append(s);
    }

    public void Bytes(byte[] b)
    {
        _bw?.Write(b);
        _debug?.Append($"|{b.Length} bytes binary|");
    }
}