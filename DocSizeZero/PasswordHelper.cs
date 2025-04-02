using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocSizeZero
{
    public static class PasswordHelper
    {
        private const int Key = 5;  // Clave constante para codificar

        public static string EncodePassword(string plainPassword)
        {
            if (string.IsNullOrEmpty(plainPassword))
                return string.Empty;

            string trimmedPassword = plainPassword.Trim();
            StringBuilder encodedPassword = new StringBuilder();

            // Agregar la longitud del texto al inicio (2 dígitos)
            encodedPassword.Append(trimmedPassword.Length.ToString("D2"));

            // Codificar cada carácter
            foreach (char ch in trimmedPassword)
            {
                int codeSym = ch + Key;
                encodedPassword.Append((char)codeSym);
            }

            return encodedPassword.ToString();
        }

        public static string DecodePassword(string encodedPassword)
        {
            if (string.IsNullOrEmpty(encodedPassword) || encodedPassword.Length < 2)
                return string.Empty;

            // Leer la longitud del texto original (primeros 2 caracteres)
            int length = int.Parse(encodedPassword.Substring(0, 2));

            StringBuilder decodedPassword = new StringBuilder();

            // Decodificar cada carácter
            for (int i = 2; i < encodedPassword.Length; i++)
            {
                int codeSym = encodedPassword[i] - Key;
                decodedPassword.Append((char)codeSym);
            }

            return decodedPassword.ToString();
        }
    }
}
