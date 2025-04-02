using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.IO;

namespace DocSizeZero
{
    /// <summary>
    /// Clase responsable de la autenticación de usuarios, permitiendo autenticación manual (con LoginForm)
    /// y automático (mediante parametros de inicio)
    /// </summary>
    public class UserAuthenticator
    {
        private readonly DatabaseManager _tmpDatabaseManager;   // Conexión temporal principal;
        private int _loginAttempts;                             // Contador de intentos de inicio de sesión

        // Propiedades públicas con la informacón del usuario autentificado.
        public int UserId { get; private set; }
        public string Username { get; private set; }
        public string EncodedPassword { get; private set; }
        public string UserFullName { get; private set; }
        public string Password { get; private set; }
        public string UserRole { get; private set; }
        public bool PasswordChangeRequired { get; private set; }

        /// <summary>
        /// Constructor que recibe una conección temporal con la base de datos
        /// </summary>
        public UserAuthenticator(DatabaseManager tmpDatabaseManager)
        {
            _tmpDatabaseManager = tmpDatabaseManager ?? throw new ArgumentNullException(nameof(tmpDatabaseManager));
            _loginAttempts = AppConstants.LoginAttemptLimit;        // Número maximo de intentos permitidos
        }

        /// <summary>
        /// Realizar la autenticación del usuario manual o automatico
        /// </summary>
        /// <param name="username">Nombre del usuario.</param>
        /// <param name="password">Contraseña en texto plano.</param>
        /// <param name="errorMessage">Mensaje de error en caso de fallo.</param>
        /// <returns>True si la autenticación es exitosa, False si falla.</returns>
        public bool Authenticate(string username, string password, out string errorMessage)
        {
            errorMessage = string.Empty;

            // Verificar que la conexión a la base de datos esté activa.
            if (_tmpDatabaseManager._temporaryConnection == null || _tmpDatabaseManager._temporaryConnection.State != ConnectionState.Open)
            {
                errorMessage = "Тимчасове з'єднання відсутнє.";
                return false;
            }

            // Verificar que los campos no estén vacíos
            if (string.IsNullOrWhiteSpace(username) || string.IsNullOrWhiteSpace(password))
            {
                errorMessage = "Будь-ласка, введіть ваш логін і пароль.";
                return false;
            }

            Username = username.Trim();
            EncodedPassword = PasswordHelper.EncodePassword(password.Trim());

            // Validar el usuario en la base de datos.
            if (!_tmpDatabaseManager.ValidateUser(Username, EncodedPassword, out int id, out string role, out bool pwdChangeRequired, out string fullName))
            {
                _loginAttempts--;
                errorMessage = $"Корстувач або пароль не вірні. Залишилося спроб: {_loginAttempts}";

                if (_loginAttempts <= 0)
                {
                    throw new UnauthorizedAccessException("Ви перевищили кількість спроб розпочати сесію.");
                }
                return false;
            }

            // Configurar datos del usuario autenticado.
            SetAuthenticatedUser(id, role, password, fullName);

            // Guardar el último usuario en el archivo INI.

            return true;
        }

        /// <summary>
        /// Configura la información del usuario autenticado.
        /// </summary>
        private void SetAuthenticatedUser(int id, string role, string password, string fullName)
        {
            UserId = id;
            UserRole = role;
            Password = password;
            UserFullName = fullName;
        }

        /// <summary>
        /// Guardar el último usuario en el archivo INI.
        /// </summary>
        private void SaveLastUsername()
        {
            string iniFilePath = Path.Combine(AppConstants.DirectoryOfConfig, AppConstants.SqlIniFile);
            IniFileHelper.SaveToIniFile("Options", "LastLoginUser", Username, iniFilePath);
        }
    }
}
