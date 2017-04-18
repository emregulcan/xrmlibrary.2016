using System;
using System.IO;
using System.Web;
using Sasha.Exceptions;
using Sasha.ExtensionMethods;

namespace XrmLibrary.EntityHelpers.Utils
{
    /// <summary>
    /// 
    /// </summary>
    public class AttachmentHelper
    {
        #region | Public Methods |

        /// <summary>
        /// Create an <see cref="AttachmentData"/> from <c>Base64</c> content.
        /// </summary>
        /// <param name="content"><c>Base64</c> content</param>
        /// <param name="fileName"></param>
        /// <returns>
        /// <see cref="AttachmentData"/>
        /// </returns>
        public static AttachmentData CreateFromBase64(string content, string fileName)
        {
            ExceptionThrow.IfNullOrEmpty(content, "content");
            ExceptionThrow.IfNullOrEmpty(fileName, "fileName");

            string tempFileName = Path.GetTempFileName();

            byte[] fileByte = Convert.FromBase64String(content);
            File.WriteAllBytes(tempFileName, fileByte);

            FileInfo fileInfo = new FileInfo(tempFileName);

            AttachmentData result = new AttachmentData()
            {
                Data = new AttachmentFileData()
                {
                    Base64 = fileByte.ToBase64String(),
                    ByteArray = fileByte
                },
                Meta = new AttachmentFileMeta
                {
                    Directory = fileInfo.Directory.FullName,
                    Extension = fileInfo.Extension,
                    FullPath = fileInfo.FullName,
                    MimeType = MimeMapping.GetMimeMapping(fileInfo.Name),
                    Name = !string.IsNullOrEmpty(fileName) ? fileName : fileInfo.Name,
                    Size = fileByte.Length
                }
            };

            File.Delete(tempFileName);

            return result;
        }

        /// <summary>
        /// Create an <see cref="AttachmentData"/> from <c>path</c>.
        /// </summary>
        /// <param name="path">File full path (with <c>drive\folder\filename.extension</c>)</param>
        /// <param name="fileName">If you want override filename on CRM please provide that. Otherwise leave blank.</param>
        /// <returns>
        /// <see cref="AttachmentData"/>
        /// </returns>
        public static AttachmentData CreateFromPath(string path, string fileName)
        {
            ExceptionThrow.IfNullOrEmpty(path, "path");
            ExceptionThrow.IfFileNotFound(path, "path");

            FileInfo fileInfo = new FileInfo(path);
            byte[] fileByte = File.ReadAllBytes(path);

            AttachmentData result = new AttachmentData()
            {
                Data = new AttachmentFileData()
                {
                    Base64 = fileByte.ToBase64String(),
                    ByteArray = fileByte
                },
                Meta = new AttachmentFileMeta
                {
                    Directory = fileInfo.Directory.FullName,
                    Extension = fileInfo.Extension,
                    FullPath = fileInfo.FullName,
                    MimeType = MimeMapping.GetMimeMapping(fileInfo.Name),
                    Name = !string.IsNullOrEmpty(fileName) ? fileName : fileInfo.Name,
                    Size = fileByte.Length
                }
            };

            return result;
        }

        /// <summary>
        /// Create an <see cref="AttachmentData"/> from <see cref="Stream"/>.
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="fileName"></param>
        /// <returns>
        /// <see cref="AttachmentData"/>
        /// </returns>
        public static AttachmentData CreateFromStream(Stream stream, string fileName)
        {
            ExceptionThrow.IfNullOrEmpty(fileName, "fileName");

            string tempFileName = Path.GetTempFileName();

            using (FileStream fs = File.OpenWrite(tempFileName))
            {
                stream.CopyTo(fs);
            }

            var result = CreateFromPath(tempFileName, fileName);
            File.Delete(tempFileName);

            return result;
        }

        /// <summary>
        /// Creates an <see cref="AttachmentData"/> from <c>Byte Array</c> content.
        /// </summary>
        /// <param name="content"></param>
        /// <param name="fileName"></param>
        /// <returns>
        /// <see cref="AttachmentData"/>
        /// </returns>
        public static AttachmentData CreateFromByte(byte[] content, string fileName)
        {
            ExceptionThrow.IfNull(content, "content");
            ExceptionThrow.IfNullOrEmpty(fileName, "fileName");

            string tempFileName = Path.GetTempFileName();

            File.WriteAllBytes(tempFileName, content);

            FileInfo fileInfo = new FileInfo(tempFileName);

            var result = CreateFromPath(tempFileName, fileName);
            File.Delete(tempFileName);

            return result;
        }

        #endregion
    }
}
