using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Text;

namespace ITGrass.Common.Utility.JsonHelper
{
    public static class JsonHelper
    {

        public static string toJson(this object obj)
        {
            if (obj != null)
            {
                return JsonConvert.SerializeObject(obj);
            }
            return string.Empty;
        }

        public static T toObject<T>(this string str) 
        {
            if (string.IsNullOrEmpty(str))
            {
                return default(T);
            }
            return JsonConvert.DeserializeObject<T>(str);
        }

    }
}
