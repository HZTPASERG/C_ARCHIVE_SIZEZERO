using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace DocSizeZero
{
    public class UserCatalog
    {
        private string _basePath;               // Ruta base para los directorios de usuarios
        private string _userLogin;              // Nombre de usuario
        private string _userPath;               // Ruta completa del directorio del usuario
        private string _lockFile;               // Archivo de bloqueo
        private FileStream _lockFileStream;     // Mantiene el archivo de bloqueo abierto

        public UserCatalog(string basePath, string userLogin)
        {
            _basePath = basePath;
            _userLogin = userLogin;
            _userPath = Path.Combine(_basePath, _userLogin, AppConstants.DirectoryOtherUsersProgram);       // Ruta: basePath/userLogin/ARCHIV_USER
            _lockFile = Path.Combine(_userPath, $"{_userLogin}.lock");                                      // Archivo de bloqueo: userPath/userLogin.lock
        }

        public bool EnsureUserCatalog()
        {
            try
            {
                // 1. Crear el directorio del usuario si no existe
                if (!Directory.Exists(_userPath))
                {
                    Directory.CreateDirectory(_userPath);
                }

                // 2. Intentar abrir el archivo en modo exclusivo
                if (File.Exists(_lockFile))
                {
                    // Si ya existe el archivo, intentar abrirlo sin permitir compartirlo
                    _lockFileStream = new FileStream(_lockFile, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
                }
                else
                {
                    // Si no existe el archivo, crearlo en modo exclusivo
                    _lockFileStream = new FileStream(_lockFile, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None);
                }

                // 3. Verificar si existen los archivos de configuración y copiarlos si es necesario
                string defaultConfigFile = Path.Combine(AppConstants.DirectoryOfConfig, AppConstants.DefaultConfigIni);
                string userConfigFile = Path.Combine(_userPath, AppConstants.ConfigIni);

                if (!File.Exists(userConfigFile) && File.Exists(defaultConfigFile))
                {
                    File.Copy(defaultConfigFile, userConfigFile);
                }

                return true; // El directorio está disponible y listo
            }
            catch (IOException)
            {
                // Ocurre si otro proceso ya tiene el archivo de bloqueo abierto
                Console.WriteLine("IOException in EnsureUserCatalog: Каталог користувача зайнятий іншою сесією.");
                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al verificar el directorio del usuario: {ex.Message}");
                return false;
            }
        }

        public void ReleaseLock()
        {
            try
            {
                // Cerrar el FileStream antes de eliminar el archivo
                _lockFileStream?.Close();
                _lockFileStream?.Dispose();

                // Eliminar el archivo de bloqueo al finalizar la sesión
                if (File.Exists(_lockFile))
                {
                    File.Delete(_lockFile);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al liberar el archivo de bloqueo: {ex.Message}");
            }
        }

    }
}
