using System;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Text;

namespace ITGrass.Common.Utility
{
    public static class Md5Helper
    {
        public static string MD5Encrypt32(string strText, bool IsLower)
        {
            MD5 md5 = new MD5CryptoServiceProvider();
            byte[] bytes = System.Text.Encoding.UTF8.GetBytes(strText);

            bytes = md5.ComputeHash(bytes);
            md5.Clear();

            string ret = "";
            for (int i = 0; i < bytes.Length; i++)
            {
                ret += bytes[i].ToString("x2");
            }

            return IsLower ? ret.ToLower() : ret.ToUpper();
        }

        public static string MD5Encrypt16(string strText, bool IsLower)
        {
            string md5Pwd = string.Empty;

            //使用加密服务提供程序
            MD5CryptoServiceProvider md5 = new MD5CryptoServiceProvider();

            //将指定的字节子数组的每个元素的数值转换为它的等效十六进制字符串表示形式。
            md5Pwd = BitConverter.ToString(md5.ComputeHash(UTF8Encoding.Default.GetBytes(strText)), 4, 8);

            md5Pwd = md5Pwd.Replace("-", "");

            return IsLower ? md5Pwd.ToLower() : md5Pwd.ToUpper();
        }
    }
}
