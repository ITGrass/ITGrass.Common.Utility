using ITGrass.Common.Utility;
using System;

namespace Test
{
    class Program
    {
        static void Main(string[] args)
        {
            string s1 = Md5Helper.MD5Encrypt32("丁浩", true);
            string s2 = Md5Helper.MD5Encrypt16("丁浩", true);
        }
    }
}
