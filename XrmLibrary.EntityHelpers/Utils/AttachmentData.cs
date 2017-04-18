using System;

namespace XrmLibrary.EntityHelpers.Utils
{
    public class AttachmentData
    {
        public AttachmentFileData Data { get; set; }
        public AttachmentFileMeta Meta { get; set; }
    }

    public class AttachmentFileData
    {
        public string Base64 { get; set; }
        public byte[] ByteArray { get; set; }
    }

    public class AttachmentFileMeta
    {
        public string Name { get; set; }
        public string Extension { get; set; }
        public string MimeType { get; set; }
        public long Size { get; set; }
        public string Directory { get; set; }
        public string FullPath { get; set; }
    }
}
