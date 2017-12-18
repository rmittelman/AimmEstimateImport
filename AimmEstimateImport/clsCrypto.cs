using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Security.Cryptography;

namespace AimmEstimateImport
{
    public static class clsCrypto
    {
        public const String strPermutation = "ouiveyxaqtd";
        public const Int32 bytePermutation1 = 0x19;
        public const Int32 bytePermutation2 = 0x59;
        public const Int32 bytePermutation3 = 0x17;
        public const Int32 bytePermutation4 = 0x41;

        /// <summary>
        /// Encrypt a text string
        /// </summary>
        /// <param name="text">Text to encrypt.</param>
        /// <param name="password">Optional. If used, must be used to decrypt.</param>
        /// <remarks>Uses a default 8-byte seed.</remarks>
        /// <returns></returns>
        public static string Encrypt(string text, string password = "")
        {
            if(password == "")
                password = "ouiveyxaqtd";
            return Convert.ToBase64String(Encrypt(Encoding.UTF8.GetBytes(text), password));
        }

        /// <summary>
        /// Decrypt a text string
        /// </summary>
        /// <param name="text">Text to decrypt.</param>
        /// <param name="password">Optional. If used to encrypt, must be used to decrypt.</param>
        /// <remarks>Uses a default 8-byte seed.</remarks>
        /// <returns></returns>
        public static string Decrypt(string text, string password = "")
        {
            if(password == "")
                password = "ouiveyxaqtd";
            return Encoding.UTF8.GetString(Decrypt(Convert.FromBase64String(text), password));
        }

        private static byte[] Encrypt(byte[] text, string password)
        {
            byte[] salt = new byte[] { 0x19, 0x59, 0x17, 0x41, 0x13, 0x29, 0x31, 0x66 };
            using(Rfc2898DeriveBytes pdb = new Rfc2898DeriveBytes(password, salt, 10000))
            {
                using(MemoryStream ms = new MemoryStream())
                {

                    using(Aes aes = new AesManaged())
                    {
                        aes.Key = pdb.GetBytes(aes.KeySize / 8);
                        aes.IV = pdb.GetBytes(aes.BlockSize / 8);
                        try
                        {
                            using(CryptoStream cs = new CryptoStream(ms, aes.CreateEncryptor(), 
                                CryptoStreamMode.Write))
                            {
                                cs.Write(text, 0, text.Length);
                                cs.Close();
                                return ms.ToArray();
                            }

                        }
                        catch(Exception)
                        {
                            return new byte[0];
                        }
                    }
                }
            }
        }

        private static byte[] Decrypt(byte[] text, string password)
        {
            byte[] salt = new byte[] { 0x19, 0x59, 0x17, 0x41, 0x13, 0x29, 0x31, 0x66 };
            using(Rfc2898DeriveBytes pdb = new Rfc2898DeriveBytes(password, salt, 10000))
            {
                using(MemoryStream ms = new MemoryStream())
                {
                    using(Aes aes = new AesManaged())
                    {
                        aes.Key = pdb.GetBytes(aes.KeySize / 8);
                        aes.IV = pdb.GetBytes(aes.BlockSize / 8);

                        try
                        {
                            using(CryptoStream cs = new CryptoStream(ms, aes.CreateDecryptor(), 
                                CryptoStreamMode.Write))
                            {
                                cs.Write(text, 0, text.Length);
                                cs.Close();
                                return ms.ToArray();
                            }
                        }
                        catch(Exception)
                        {
                            return new byte[0];
                        }
                    }
                }
            }
        }
    }
}
