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
        private static readonly HttpClient _httpClient;

        static HttpManager()
        {
            _httpClient = new HttpClient();
        }

        /// <summary>
        /// Get请求
        /// </summary>
        /// <param name="uri">url地址</param>
        /// <returns></returns>
        public static async Task<HttpResponseMessage> Get(Uri uri)
        {
            return await _httpClient.GetAsync(uri);
        }



        /// <summary>
        /// Get请求
        /// </summary>
        /// <param name="url">地址</param>
        /// <param name="headers">请求头信息</param>
        /// <returns></returns>
        public static async Task<HttpResponseMessage> Get(string url, List<KeyValuePair<string, string>> headers = null)
        {
            HttpRequestMessage request = new HttpRequestMessage()
            {
                RequestUri = new Uri(url),
                Method = HttpMethod.Get,
            };
            if (headers != null && headers.Count > 0)
            {
                request.Headers.Clear();

                foreach (var header in headers)
                {
                    request.Headers.Add(header.Key, header.Value);

                }
            }
            return await _httpClient.SendAsync(request);
        }



        /// <summary>
        /// post 请求
        /// </summary>
        /// <param name="url">地址</param>
        /// <param name="paramList">参数</param>
        /// <param name="headers">头信息</param>
        /// <returns></returns>
        public static async Task<HttpResponseMessage> PostAsync(string url, List<KeyValuePair<String, String>> paramList, List<KeyValuePair<string, string>> headers = null)
        {
            FormUrlEncodedContent formUrlEncodedContent = new FormUrlEncodedContent(paramList);
            if (headers != null && headers.Count > 0)
            {
                formUrlEncodedContent.Headers.Clear();
                foreach (var header in headers)
                {
                    formUrlEncodedContent.Headers.Add(header.Key, header.Value);
                }
            }
            return await _httpClient.PostAsync(new Uri(url), formUrlEncodedContent);

        }


        /// <summary>
        /// Post方法请求 raw data
        /// </summary>
        /// <param name="url">请求地址</param>
        /// <param name="content">raw data</param>
        /// <returns></returns>
        public static async Task<HttpResponseMessage> PostAsync(string url, string content, List<KeyValuePair<string, string>> headers = null)
        {
            StringContent stringContent = new StringContent(content, Encoding.UTF8);
            if (headers != null && headers.Count > 0)
            {
                stringContent.Headers.Clear();
                foreach (var header in headers)
                {
                    stringContent.Headers.Add(header.Key, header.Value);
                }
            }
            return await _httpClient.PostAsync(new Uri(url), stringContent);
        }
    }
}
