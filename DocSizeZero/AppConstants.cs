using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocSizeZero
{
    public static class AppConstants
    {
        // Títulos y textos
        public const string NameMainWindow = "Докуименти з розміром 0 Мб 'УКРЕЛЕКТРОАПАРАТ'";
        public const string ErrorTitle = "Ошибка !";

        // Archivos
        public const string ErrFile = "ERR.TXT";
        public const string SqlLogFile = "SQLERR.TXT";
        public const string VfpLogFile = "VFPERR.TXT";
        public const string IniFile = "USER.INI";
        public const string SqlIniFile = "CONFIG.INI";

        // Configuración SQL
        public const string SqlDriver = "SQL Server";
        public const string SqlAppRole = "WORK_ROLE";
        public const string RolePass = "wrk";
        public const string SqlServer = "NT3";
        public const string DatabaseName = "master";
        public const string SqlLogin = "LoginForAll";
        public const string LoginPassword = "all";
        public const string SqlLoginAdmin = "AppAdmin";
        public const string AdminPassword = "adm";

        // Directorios
        public const string DirectoryOtherUsers = @"\\nt3\Users$";
        public const string DirectoryOtherUsersProgram = @"DocSizeZero";
        public const string DirectoryOfConfig = "USERS";
        public const int LoginAttemptLimit = 3;
        public const string LockFileName = "SERG.txt";
        public const string DefaultConfigIni = "USER.INI";
        public const string ConfigIni = "USER.INI";
    }
}
