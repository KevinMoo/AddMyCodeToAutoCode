namespace MSC.CommonLib
{
    using System;
    using System.Drawing;
    using System.Drawing.Imaging;
    using System.IO;

    public class ImageOperation
    {
        public static string CrateImageFileByBytes(byte[] bytes)
        {
            string filename = Path.GetTempPath() + GenerateStringID() + ".JPG";
            if ((bytes != null) && (bytes.Length != 0))
            {
                Image.FromStream(new MemoryStream(bytes)).Save(filename, ImageFormat.Jpeg);
                return filename;
            }
            return null;
        }

        public static string GenerateStringID()
        {
            long num = 1L;
            foreach (byte num2 in Guid.NewGuid().ToByteArray())
            {
                num *= num2 + 1;
            }
            return string.Format("{0:x}", num - DateTime.Now.Ticks);
        }

        public static byte[] GetBytesFromFilename(string pFileName)
        {
            if (!File.Exists(pFileName))
            {
                return null;
            }
            Image image = Image.FromFile(pFileName);
            MemoryStream stream = new MemoryStream();
            image.Save(stream, ImageFormat.Jpeg);
            stream.Position = 0L;
            byte[] buffer = new byte[stream.Length];
            stream.Read(buffer, 0, buffer.Length);
            return buffer;
        }
    }
}

