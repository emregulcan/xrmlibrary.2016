using System;
using System.IO;
using Sasha.Exceptions;
using Sasha.ExtensionMethods;

namespace XrmLibrary.EntityHelpers.Utils
{
    /// <summary>
    /// Includes helper methods for file manupulation with <see cref="System.IO"/>.
    /// </summary>
    public class FileHelper
    {
        #region | Public Methods |

        /// <summary>
        /// Retrieve <see cref="FileInfo"/> data from <c>path</c>.
        /// </summary>
        /// <param name="path"></param>
        /// <returns>
        /// <see cref="FileInfo"/>
        /// </returns>
        public static FileInfo GetFileInfoFromPath(string path)
        {
            ExceptionThrow.IfNullOrEmpty(path, "path");
            ExceptionThrow.IfFileNotFound(path, "path");

            return new FileInfo(path);
        }

        /// <summary>
        /// Retrieve <see cref="FileInfo"/> data from <c>Base64</c> content.
        /// </summary>
        /// <param name="content"><c>Base64</c> content</param>
        /// <returns>
        /// <see cref="FileInfo"/>
        /// </returns>
        public static FileInfo GetFileInfoFromBase64(string content)
        {
            ExceptionThrow.IfNullOrEmpty(content, "Content");

            string tempFileName = Path.GetTempFileName();

            byte[] fileByte = Convert.FromBase64String(content);
            File.WriteAllBytes(tempFileName, fileByte);

            FileInfo result = GetFileInfoFromPath(tempFileName);
            File.Delete(tempFileName);

            return result;
        }

        /// <summary>
        /// Retrieve <see cref="FileInfo"/> data from <see cref="Stream"/> content.
        /// </summary>
        /// <param name="stream"></param>
        /// <returns>
        /// <see cref="FileInfo"/>
        /// </returns>
        public static FileInfo GetFileInfoFromStream(Stream stream)
        {
            string tempFileName = Path.GetTempFileName();

            using (FileStream fs = File.OpenWrite(tempFileName))
            {
                stream.CopyTo(fs);
            }

            FileInfo result = GetFileInfoFromPath(tempFileName);
            File.Delete(tempFileName);

            return result;
        }

        /// <summary>
        /// Retrieve data from <see cref="Stream"/> content.
        /// </summary>
        /// <param name="stream"></param>
        /// <returns>
        /// 
        /// </returns>
        public static byte[] GetByteFromStream(Stream stream)
        {
            return stream.ToByteArray();
        }

        #endregion
    }
}
