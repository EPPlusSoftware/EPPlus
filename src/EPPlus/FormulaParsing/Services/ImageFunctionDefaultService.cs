using System;
using System.Collections.Generic;
using System.Linq;
#if !NET35
using System.Net.Http;
using System.Threading.Tasks;
#else

#endif
using System.Text;
using OfficeOpenXml.Interfaces.Net;

namespace OfficeOpenXml.FormulaParsing.Services
{
    internal class ImageFunctionDefaultService : IHttpsService
    {
        public byte[] Download(string url)
        {
#if !NET35
            using (HttpClient client = new HttpClient())
            {
                return client.GetByteArrayAsync(url).Result;
            }
#else
            return null;
#endif
        }
    }
}
