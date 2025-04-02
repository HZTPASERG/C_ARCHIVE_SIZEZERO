using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace DocSizeZero
{
    public static class IniFileHelper
    {
        [DllImport("kernel32")]
        private static extern long WritePrivateProfileString(string section, string key, string value, string filePath);

        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string defaultValue, StringBuilder result, int size, string filePath);

        public static string ReadFromIniFile(string section, string key, string filePath)
        {
            StringBuilder result = new StringBuilder(255);
            GetPrivateProfileString(section, key, "", result, 255, filePath);
            return result.ToString();
        }

        public static void SaveToIniFile(string section, string key, string value, string filePath)
        {
            WritePrivateProfileString(section, key, value, filePath);
        }
    }
}
