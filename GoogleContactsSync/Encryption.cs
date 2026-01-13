using System;
using System.IO;
using System.Runtime.Serialization;
using System.Security.Cryptography;
using System.Text;

namespace GoContactSyncMod
{
    //class AccountEncryption
    internal partial class Encryption
    {
        private static byte[] GetKey(string email)
        {
            var enc = new ASCIIEncoding();

            using (SHA256 shaM = new SHA256Managed())
            {
                return shaM.ComputeHash(enc.GetBytes(email));
            }
        }

        private static byte[] GetIV(string email)
        {
            var enc = new ASCIIEncoding();

            using (var md5 = MD5.Create())
            {
                return md5.ComputeHash(enc.GetBytes(email));
            }
        }

        public static string DecryptPassword(string email, string encryptedPassword)
        {
            try
            {
                using (var rijndael = new RijndaelManaged())
                {
                    //RijndaelManaged rijndael = new RijndaelManaged();

                    rijndael.IV = GetIV(email);
                    rijndael.Key = GetKey(email);

                    var decryptor = rijndael.CreateDecryptor(rijndael.Key, rijndael.IV);

                    //Now decrypt the previously encrypted password using the decryptor obtained in the above step.
                    var encrypted = HexEncoding.GetBytes(encryptedPassword);
                    using (var msDecrypt = new MemoryStream(encrypted))
                    {
                        var csDecrypt = new CryptoStream(msDecrypt, decryptor, CryptoStreamMode.Read);

                        //byte[] fromEncrypt = new byte[encrypted.Length];
                        var fromEncrypt = ReadFully(csDecrypt);//(

                        //Read the data out of the crypto stream.
                        //csDecrypt.Read(fromEncrypt, 0, fromEncrypt.Length);
                        //csDecrypt.Close();
                        //msDecrypt.Close();

                        //Convert the byte array back into a string.
                        var result = new ASCIIEncoding().GetString(fromEncrypt);
                        return result;
                    }
                }
            }
            catch (Exception e)
            {
                throw new PasswordEncryptionException("Error decryption password", e);
            }
        }
        public static string EncryptPassword(string email, string password)
        {
            using (var rijndael = new RijndaelManaged())
            {

                rijndael.IV = GetIV(email);
                rijndael.Key = GetKey(email);

                var encryptor = rijndael.CreateEncryptor(rijndael.Key, rijndael.IV);

                var toEncrypt = new ASCIIEncoding().GetBytes(password);

                //Encrypt the data.
                using (var msEncrypt = new MemoryStream())
                {
                    var csEncrypt = new CryptoStream(msEncrypt, encryptor, CryptoStreamMode.Write);

                    //byte[] encrypted = ReadFully(csEncrypt);
                    /*byte[] codes = new byte[msEncrypt.Capacity];
                    csEncrypt.Read(codes, 0, codes.Length);
                    csEncrypt.Close();
                    msEncrypt.Close();*/

                    //Write all data to the crypto stream and flush it.
                    csEncrypt.Write(toEncrypt, 0, toEncrypt.Length);
                    csEncrypt.FlushFinalBlock();

                    //Convert the byte array back into a string.
                    return HexEncoding.ToString(msEncrypt.ToArray());
                }
            }
        }
        public static byte[] ReadFully(Stream stream)
        {
            var buffer = new byte[32768];
            using (var ms = new MemoryStream())
            {
                while (true)
                {
                    var read = stream.Read(buffer, 0, buffer.Length);
                    if (read <= 0)
                    {
                        return ms.ToArray();
                    }

                    ms.Write(buffer, 0, read);
                }
            }
        }
    }
    [Serializable]
    public class PasswordEncryptionException : Exception
    {
        public PasswordEncryptionException() : base() { }
        public PasswordEncryptionException(string message) : base(message) { }
        public PasswordEncryptionException(string message, Exception innerException) : base(message, innerException) { }
        protected PasswordEncryptionException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }

    /// <summary>
    /// Summary description for HexEncoding.
    /// </summary>
    internal class HexEncoding
    {
        public HexEncoding()
        {
            //
            // TODO: Add constructor logic here
            //
        }
        public static int GetByteCount(string hexString)
        {
            var numHexChars = 0;
            char c;
            // remove all none A-F, 0-9, characters
            for (var i = 0; i < hexString.Length; i++)
            {
                c = hexString[i];
                if (IsHexDigit(c))
                {
                    numHexChars++;
                }
            }
            // if odd number of characters, discard last character
            if (numHexChars % 2 != 0)
            {
                numHexChars--;
            }
            return numHexChars / 2; // 2 characters per byte
        }

        /// <summary>
        /// Creates a byte array from the hexadecimal string. Each two characters are combined
        /// to create one byte. First two hexadecimal characters become first byte in returned array.
        /// Non-hexadecimal characters are ignored.
        /// </summary>
        /// <param name="hexString">string to convert to byte array</param>
        /// <returns>byte array, in the same left-to-right order as the hexString</returns>

        public static byte[] GetBytes(string hexString)
        {
            return GetBytes(hexString, out var discarded);
        }

        /// <summary>
        /// Creates a byte array from the hexadecimal string. Each two characters are combined
        /// to create one byte. First two hexadecimal characters become first byte in returned array.
        /// Non-hexadecimal characters are ignored.
        /// </summary>
        /// <param name="hexString">string to convert to byte array</param>
        /// <param name="discarded">number of characters in string ignored</param>
        /// <returns>byte array, in the same left-to-right order as the hexString</returns>
        public static byte[] GetBytes(string hexString, out int discarded)
        {
            discarded = 0;
            var newString = "";
            char c;
            // remove all none A-F, 0-9, characters
            for (var i = 0; i < hexString.Length; i++)
            {
                c = hexString[i];
                if (IsHexDigit(c))
                {
                    newString += c;
                }
                else
                {
                    discarded++;
                }
            }
            // if odd number of characters, discard last character
            if (newString.Length % 2 != 0)
            {
                discarded++;
                newString = newString.Substring(0, newString.Length - 1);
            }

            var byteLength = newString.Length / 2;
            var bytes = new byte[byteLength];
            string hex;
            var j = 0;
            for (var i = 0; i < bytes.Length; i++)
            {
                hex = new string(new char[] { newString[j], newString[j + 1] });
                bytes[i] = HexToByte(hex);
                j += 2;
            }
            return bytes;
        }

        public static string ToString(byte[] bytes)
        {
            var hexString = "";
            for (var i = 0; i < bytes.Length; i++)
            {
                hexString += bytes[i].ToString("X2");
            }
            return hexString;
        }

        /// <summary>
        /// Determines if given string is in proper hexadecimal string format
        /// </summary>
        /// <param name="hexString"></param>
        /// <returns></returns>
        public static bool InHexFormat(string hexString)
        {
            var hexFormat = true;

            foreach (var digit in hexString)
            {
                if (!IsHexDigit(digit))
                {
                    hexFormat = false;
                    break;
                }
            }
            return hexFormat;
        }

        /// <summary>
        /// Returns true is c is a hexadecimal digit (A-F, a-f, 0-9)
        /// </summary>
        /// <param name="c">Character to test</param>
        /// <returns>true if hex digit, false if not</returns>
        public static bool IsHexDigit(char c)
        {
            int numChar;
            var numA = Convert.ToInt32('A');
            var num1 = Convert.ToInt32('0');
            c = char.ToUpper(c);
            numChar = Convert.ToInt32(c);
            if (numChar >= numA && numChar < (numA + 6))
            {
                return true;
            }

            return numChar >= num1 && numChar < (num1 + 10);
        }
        /// <summary>
        /// Converts 1 or 2 character string into equivalant byte value
        /// </summary>
        /// <param name="hex">1 or 2 character string</param>
        /// <returns>byte</returns>
        private static byte HexToByte(string hex)
        {
            if (hex.Length > 2 || hex.Length <= 0)
            {
                throw new ArgumentException("hex must be 1 or 2 characters in length");
            }

            var newByte = byte.Parse(hex, System.Globalization.NumberStyles.HexNumber);
            return newByte;
        }
    }
}