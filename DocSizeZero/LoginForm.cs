using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using System.Diagnostics;

namespace DocSizeZero
{
    public partial class LoginForm : Form
    {
        private readonly DatabaseManager _tmpDatabaseManager; // Conexión temporal principal;
        private int _loginAttempts;

        public int UserId { get; private set; }
        public string Username { get; private set; }
        public string EncodedPassword { get; private set; }
        public string UserFullName { get; private set; }
        public string UserRole { get; private set; }
        public bool PasswordChangeRequired { get; private set; }

        public LoginForm(DatabaseManager tmpDatabaseManager)
        {
            if (tmpDatabaseManager == null)
                throw new ArgumentNullException(nameof(tmpDatabaseManager), "Об'кт DatabaseManager не може бути NULL.");

            if (tmpDatabaseManager._temporaryConnection == null || tmpDatabaseManager._temporaryConnection.State != ConnectionState.Open)
                throw new InvalidOperationException("Тимчасове з'єднання відсутнє в об'кті DatabaseManager.");

            InitializeComponent();
            _tmpDatabaseManager = tmpDatabaseManager ?? throw new ArgumentNullException(nameof(tmpDatabaseManager));
            _loginAttempts = AppConstants.LoginAttemptLimit;
            LoadPreviousUsername();
        }

        private void LoadPreviousUsername()
        {
            string iniFilePath = Path.Combine(AppConstants.DirectoryOfConfig, AppConstants.SqlIniFile);
            Username = IniFileHelper.ReadFromIniFile("Options", "LastLoginUser", iniFilePath);
            txtUsername.Text = Username;
            if (!string.IsNullOrEmpty(Username))
            {
                txtPassword.Focus();
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {
            if (_tmpDatabaseManager._temporaryConnection == null || _tmpDatabaseManager._temporaryConnection.State != ConnectionState.Open)
            {
                lblMessage.Text = "Тимчасове з'єднання відсутнє.";
                return;
            }

            Username = txtUsername.Text.Trim();
            string password = txtPassword.Text.Trim();

            if (string.IsNullOrEmpty(Username))
            {
                lblMessage.Text = "Буль-ласка, введіть ваш логін і пароль.";
                return;
            }

            EncodedPassword = PasswordHelper.EncodePassword(password);

            // Validar usuario usando la conexión temporal
            if (!_tmpDatabaseManager.ValidateUser(Username, EncodedPassword, out int id, out string role, out bool pwdChangeRequired, out string fullName))
            {
                _loginAttempts--;
                lblMessage.Text = $"Корстувач або пароль не вірні. Залишилося спроб: {_loginAttempts}";

                if (_loginAttempts <= 0)
                {
                    MessageBox.Show("Ви перевищили кількість спроб розпочати сесію.", "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Application.Exit();
                }
                return;
            }

            // Configuración después de la validación exitosa
            UserId = id;
            UserRole = role;
            PasswordChangeRequired = pwdChangeRequired;
            UserFullName = fullName;

            // Guardar el último usuario en el archivo INI
            string iniFilePath = Path.Combine(AppConstants.DirectoryOfConfig, AppConstants.SqlIniFile);
            IniFileHelper.SaveToIniFile("Options", "LastLoginUser", Username, iniFilePath);

            DialogResult = DialogResult.OK;
            Close();
        }
    }
}