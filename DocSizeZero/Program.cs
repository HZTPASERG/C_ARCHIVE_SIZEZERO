using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DocSizeZero
{
    static class Program
    {
        /// <summary>
        /// Punto de entrada principal para la aplicación.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            // Configurar estilos visuales de la aplicación
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // Crear instancia principal de la aplicación
            MainApp app = new MainApp();

            int userId = 0;
            string userName = null;
            string encodedPassword = null;
            string userRole = null;
            string userFullName = null;
            string errorMessage = string.Empty;
            bool isAuthenticated = false;

            // 1️ **Verificar si hay parámetros de entrada (usuario y contraseña)
            // autenticación automática si hay parámetros**
            if (args.Length >= 2)
            {
                userName = args[0];             // Nombre de usuario
                encodedPassword = args[1];      // Contraseña encriptada

                // Intentar inicio de sesión automático
                UserAuthenticator autenticador = new UserAuthenticator(app.DatabaseManagerTemp);
                if (autenticador.Authenticate(userName, encodedPassword, out errorMessage))
                {
                    userId = autenticador.UserId;
                    userRole = autenticador.UserRole;
                    userFullName = autenticador.UserFullName;
                    isAuthenticated = true;
                }
                else
                {
                    MessageBox.Show($"Помилка при автоматичній авторизації: {errorMessage}", "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

            }

            // 2 **Si la autenticación automática falla, mostrar la ventana de LoginForm**
            if (!isAuthenticated)
            {
                using (LoginForm loginForm = new LoginForm(app.DatabaseManagerTemp))
                {
                    if (loginForm.ShowDialog() != DialogResult.OK)
                    {
                        return;     // Salir si el usuario cancela el login
                    }

                    userId = loginForm.UserId;
                    userName = loginForm.Username;
                    encodedPassword = loginForm.EncodedPassword;
                    userRole = loginForm.UserRole;
                    userFullName = loginForm.UserFullName;
                }
            }

            // 3 **Iniciar la aplicación con la información del usuario autenticado**
            if (!app.StartApp(userId, userName, encodedPassword, userRole, userFullName))
            {
                return;         // Salir si el inicio de la aplicación falla
            }

            // 4 **Ejecutar el formulario principal
            Application.Run(new MenuForm(app));

            // Liberar los recursos asociados a la conexión temporal
            app.DatabaseManagerTemp.DisposeTemp();
            
        }
    }
}
