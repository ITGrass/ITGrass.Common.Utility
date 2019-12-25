using ITGrass.Common.Utility.Http请求;
using System;
using System.Net;
using System.Threading.Tasks;

namespace Test
{
    class Program
    {
        static async Task Main(string[] args)
        {
            
            HttpWebRequest
               Uri uri = new Uri("http://WWW.baidu.com");
            var http = await HttpManager.GetRequest(uri);
        }
    }
}
