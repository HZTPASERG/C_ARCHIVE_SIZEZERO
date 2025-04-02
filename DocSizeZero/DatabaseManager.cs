using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;

namespace DocSizeZero
{
    public class DatabaseManager : IDisposable
    {
        public SqlConnection _persistentConnection;
        public SqlConnection _temporaryConnection;
        private readonly string _persistentConnectionString;
        private string _temporaryConnectionString;

        // Constructor con parámetros para inicializar la conexión persistente
        public DatabaseManager(string server, string database, string username, string password, string appName)
        {
            _persistentConnectionString = $"Server={server};Database={database};User Id={username};Password={password};Application Name={appName};";
            _persistentConnection = new SqlConnection(_persistentConnectionString);
            OpenPersistentConnection(); // Abre la conexión persistente al inicializar
        }

        // Constructor para una conexión temporal (opcional, para flexibilidad)
        public DatabaseManager() { }

        /// <summary>
        /// Abre una conexión temporal.
        /// </summary>
        public bool OpenTemporaryConnection(string server, string database, string username, string password, string appName)
        {
            try
            {
                Debug.WriteLine("Intentando abrir conexión temporal...");
                Debug.WriteLine($"Server: {server}, Database: {database}, User: {username}, AppName: {appName}");

                // Escapar comillas en appName
                appName = appName.Replace("\"", "\"\"");

                _temporaryConnectionString = $"Server={server};Database={database};User Id={username};Password={password};Application Name={appName};";
                Debug.WriteLine($"Generated Connection String: {_temporaryConnectionString}");

                _temporaryConnection = new SqlConnection(_temporaryConnectionString);
                _temporaryConnection.Open();

                Debug.WriteLine("Conexión temporal abierta correctamente.");
                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error al abrir conexión temporal: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Valida al usuario utilizando  utilizando la conexión temporal y un procedimiento almacenado.
        /// </summary>
        public bool ValidateUser(string username, string password, out int userId, out string role, out bool pwdChangeRequired, out string fullName)
        {
            if (_temporaryConnection == null || _temporaryConnection.State != ConnectionState.Open)
            {
                throw new InvalidOperationException("La conexión temporal no está abierta.");
            }

            Debug.WriteLine("username: " + username + " | password: " + password);

            userId = 0;
            fullName = string.Empty;
            role = string.Empty;
            pwdChangeRequired = false;

            string query = "EXEC DATD..CurUser @username, @password";

            try
            {
                using (SqlCommand cmd = new SqlCommand(query, _temporaryConnection))
                {
                    cmd.Parameters.AddWithValue("@username", username);
                    cmd.Parameters.AddWithValue("@password", password);

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            fullName = reader["fullname"].ToString();
                            userId = Convert.ToInt32(reader["user_id"]);
                            if (userId == 0) return false;

                            role = reader["role"].ToString();
                            pwdChangeRequired = reader["cfg_data"].ToString().Contains("CHANGEPWDONLOGIN=-1");
                            return true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al validar usuario: {ex.Message}");
            }

            return false;
        }

        /// <summary>
        /// Inicia una sesión en la base de datos.
        /// </summary>
        public int StartSession(int userId)
        {
            try
            {
                EnsurePersistentConnection();

                using (SqlCommand command = new SqlCommand("EXEC DATD..Admin_Session_ON @USER_ID", _persistentConnection))
                {
                    command.Parameters.AddWithValue("@USER_ID", userId);

                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            return reader.GetInt32(reader.GetOrdinal("session_id"));
                        }
                        else
                        {
                            Console.WriteLine("Error: No se devolvió ningún resultado de Admin_Session_ON.");
                            return -1;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al iniciar sesión: {ex.Message}");
                return -1;
            }
        }

        /// <summary>
        /// Abre la conexión persistente si no está abierta.
        /// </summary>
        private void OpenPersistentConnection()
        {
            if (_persistentConnection.State != ConnectionState.Open)
            {
                try
                {
                    _persistentConnection.Open();
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException("No se pudo establecer la conexión con la base de datos.", ex);
                }
            }
        }

        /// <summary>
        /// Verifica si la conexión persistente está activa y la restaura si es necesario.
        /// </summary>
        public void EnsurePersistentConnection()
        {
            if (_persistentConnection.State == ConnectionState.Closed || _persistentConnection.State == ConnectionState.Broken)
            {
                _persistentConnection.Close(); // Cierra cualquier conexión rota
                OpenPersistentConnection();   // Intenta reabrir la conexión persistente
            }
        }

        /// <summary>
        /// Finaliza la sesión persistente en la base de datos.
        /// </summary>
        public void EndSession(int userId)
        {
            try
            {
                EnsurePersistentConnection();

                using (SqlCommand command = new SqlCommand("EXEC DATD..Admin_Session_OFF @USER_ID", _persistentConnection))
                {
                    command.Parameters.AddWithValue("@USER_ID", userId);
                    command.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al terminar sesión: {ex.Message}");
            }
        }

        /// <summary>
        /// Libera los recursos asociados a la conexión temporal.
        /// </summary>
        public void DisposeTemp()
        {
            if (_temporaryConnection != null)
            {
                _temporaryConnection.Close();
                _temporaryConnection.Dispose();
            }
        }

        /// <summary>
        /// Libera los recursos asociados a la conexión persistente.
        /// </summary>
        public void Dispose()
        {
            if (_persistentConnection != null)
            {
                _persistentConnection.Close();
                _persistentConnection.Dispose();
            }
        }
    }
}
