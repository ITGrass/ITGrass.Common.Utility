using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ITGrass.Common.Utility.Http请求
{
    public static class HttpManager
    {
        private static readonly HttpClient HttpClient;

        static HttpManager()
        {
            HttpClient = new HttpClient();
        }

        /// <summary>
        /// Get请求
        /// </summary>
        /// <param name="uri">url地址</param>
        /// <returns></returns>
        public static async Task<HttpResponseMessage> GetRequest(Uri uri)
        {
            return await HttpClient.GetAsync(uri);
        }

        /// <summary>
        /// Get请求(带取消操作)
        /// </summary>
        /// <param name="uri">地址</param>
        /// <param name="token">取消标记</param>
        /// <returns></returns>
        public static async Task<HttpResponseMessage> GetRequest(Uri uri, CancellationToken token)
        {
            return await HttpClient.GetAsync(uri, token);
        }

        public static async Task<string> GetRequest(string url)
        {
            return await HttpClient.GetStringAsync(url);
        }

        /// <summary>
        /// post请求
        /// </summary>
        /// <param name="uri">地址</param>
        /// <param name="parms">参数</param>
        /// <returns></returns>
        public static async Task<HttpResponseMessage> PostRequest(Uri uri, IEnumerable<KeyValuePair<string, string>> parms)
        {
            return await HttpClient.PostAsync(uri, new FormUrlEncodedContent(parms));
        }
    }
}
