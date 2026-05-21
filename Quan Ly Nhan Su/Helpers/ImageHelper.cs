using System;
using System.IO;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace TaxPersonnelManagement.Helpers
{
    public static class ImageHelper
    {
        public static BitmapSource LoadAndOrientImage(string path)
        {
            int rotationDegrees = 0;
            bool flipHorizontal = false;

            try
            {
                byte[] bytes = File.ReadAllBytes(path);
                ushort orientation = GetJpegOrientation(bytes);
                
                // Fallback to WPF metadata query if pure byte parser returned 1
                if (orientation == 1)
                {
                    orientation = GetExifOrientationFallback(path);
                }

                ParseOrientationValue(orientation, out rotationDegrees, out flipHorizontal);
            }
            catch
            {
                // Ignore metadata errors
            }

            var bi = new BitmapImage();
            bi.BeginInit();
            bi.UriSource = new Uri(path);
            bi.CacheOption = BitmapCacheOption.OnLoad;
            bi.EndInit();

            if (rotationDegrees != 0 || flipHorizontal)
            {
                var transformGroup = new TransformGroup();
                if (rotationDegrees != 0)
                {
                    transformGroup.Children.Add(new RotateTransform(rotationDegrees));
                }
                if (flipHorizontal)
                {
                    transformGroup.Children.Add(new ScaleTransform(-1, 1));
                }

                var transformed = new TransformedBitmap(bi, transformGroup);
                transformed.Freeze();
                return transformed;
            }

            bi.Freeze();
            return bi;
        }

        public static BitmapSource? LoadAndOrientImage(byte[] bytes)
        {
            if (bytes == null || bytes.Length == 0)
                return null;

            int rotationDegrees = 0;
            bool flipHorizontal = false;

            try
            {
                ushort orientation = GetJpegOrientation(bytes);
                
                // Fallback to WPF metadata query if pure byte parser returned 1
                if (orientation == 1)
                {
                    orientation = GetExifOrientationFallback(bytes);
                }

                ParseOrientationValue(orientation, out rotationDegrees, out flipHorizontal);
            }
            catch
            {
                // Ignore metadata errors
            }

            var bi = new BitmapImage();
            bi.BeginInit();
            bi.StreamSource = new MemoryStream(bytes);
            bi.CacheOption = BitmapCacheOption.OnLoad;
            bi.EndInit();

            if (rotationDegrees != 0 || flipHorizontal)
            {
                var transformGroup = new TransformGroup();
                if (rotationDegrees != 0)
                {
                    transformGroup.Children.Add(new RotateTransform(rotationDegrees));
                }
                if (flipHorizontal)
                {
                    transformGroup.Children.Add(new ScaleTransform(-1, 1));
                }

                var transformed = new TransformedBitmap(bi, transformGroup);
                transformed.Freeze();
                return transformed;
            }

            bi.Freeze();
            return bi;
        }

        private static ushort GetJpegOrientation(byte[] bytes)
        {
            if (bytes == null || bytes.Length < 4) return 1;
            // Check SOI marker (FF D8)
            if (bytes[0] != 0xFF || bytes[1] != 0xD8) return 1;

            int i = 2;
            while (i < bytes.Length - 4)
            {
                // Find next marker (FF xx)
                if (bytes[i] != 0xFF)
                {
                    i++;
                    continue;
                }

                byte marker = bytes[i + 1];
                if (marker == 0xFF)
                {
                    i++;
                    continue;
                }

                // Markers without payload
                if (marker == 0xD8 || marker == 0xD9 || marker == 0xDA)
                    break;

                int len = (bytes[i + 2] << 8) | bytes[i + 3];
                if (marker == 0xE1) // APP1
                {
                    if (i + 10 < bytes.Length &&
                        bytes[i + 4] == 0x45 && // E
                        bytes[i + 5] == 0x78 && // x
                        bytes[i + 6] == 0x69 && // i
                        bytes[i + 7] == 0x66 && // f
                        bytes[i + 8] == 0x00 &&
                        bytes[i + 9] == 0x00)
                    {
                        return ParseExifOrientation(bytes, i + 10, len - 6);
                    }
                }

                i += 2 + len;
            }
            return 1;
        }

        private static ushort ParseExifOrientation(byte[] bytes, int start, int length)
        {
            if (start + 8 >= bytes.Length) return 1;

            bool isLittleEndian = bytes[start] == 0x49 && bytes[start + 1] == 0x49; // "II"
            if (!isLittleEndian && !(bytes[start] == 0x4D && bytes[start + 1] == 0x4D)) // "MM"
                return 1;

            ushort magic = ReadUShort(bytes, start + 2, isLittleEndian);
            if (magic != 42) return 1;

            uint ifdOffset = ReadUInt(bytes, start + 4, isLittleEndian);
            int ifdStart = start + (int)ifdOffset;
            if (ifdStart + 2 >= bytes.Length) return 1;

            ushort entriesCount = ReadUShort(bytes, ifdStart, isLittleEndian);
            int entryOffset = ifdStart + 2;

            for (int i = 0; i < entriesCount; i++)
            {
                if (entryOffset + 12 >= bytes.Length) break;

                ushort tag = ReadUShort(bytes, entryOffset, isLittleEndian);
                if (tag == 0x0112) // Orientation Tag
                {
                    return ReadUShort(bytes, entryOffset + 8, isLittleEndian);
                }

                entryOffset += 12;
            }

            return 1;
        }

        private static ushort ReadUShort(byte[] bytes, int offset, bool isLittleEndian)
        {
            if (isLittleEndian)
                return (ushort)(bytes[offset] | (bytes[offset + 1] << 8));
            else
                return (ushort)((bytes[offset] << 8) | bytes[offset + 1]);
        }

        private static uint ReadUInt(byte[] bytes, int offset, bool isLittleEndian)
        {
            if (isLittleEndian)
                return (uint)(bytes[offset] | (bytes[offset + 1] << 8) | (bytes[offset + 2] << 16) | (bytes[offset + 3] << 24));
            else
                return (uint)((bytes[offset] << 24) | (bytes[offset + 1] << 16) | (bytes[offset + 2] << 8) | bytes[offset + 3]);
        }

        private static ushort GetExifOrientationFallback(string path)
        {
            try
            {
                using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    var decoder = BitmapDecoder.Create(stream, BitmapCreateOptions.DelayCreation, BitmapCacheOption.None);
                    if (decoder.Frames.Count > 0)
                    {
                        var frame = decoder.Frames[0];
                        var metadata = frame.Metadata as BitmapMetadata;
                        if (metadata != null)
                        {
                            const string orientationQuery = "/app1/ifd0/{ushort=274}";
                            if (metadata.ContainsQuery(orientationQuery))
                            {
                                var val = metadata.GetQuery(orientationQuery);
                                if (val != null)
                                {
                                    return Convert.ToUInt16(val);
                                }
                            }
                        }
                    }
                }
            }
            catch
            {
                // Ignore
            }
            return 1;
        }

        private static ushort GetExifOrientationFallback(byte[] bytes)
        {
            try
            {
                using (var stream = new MemoryStream(bytes))
                {
                    var decoder = BitmapDecoder.Create(stream, BitmapCreateOptions.DelayCreation, BitmapCacheOption.None);
                    if (decoder.Frames.Count > 0)
                    {
                        var frame = decoder.Frames[0];
                        var metadata = frame.Metadata as BitmapMetadata;
                        if (metadata != null)
                        {
                            const string orientationQuery = "/app1/ifd0/{ushort=274}";
                            if (metadata.ContainsQuery(orientationQuery))
                            {
                                var val = metadata.GetQuery(orientationQuery);
                                if (val != null)
                                {
                                    return Convert.ToUInt16(val);
                                }
                            }
                        }
                    }
                }
            }
            catch
            {
                // Ignore
            }
            return 1;
        }

        private static void ParseOrientationValue(ushort orientation, out int rotationDegrees, out bool flipHorizontal)
        {
            rotationDegrees = 0;
            flipHorizontal = false;

            switch (orientation)
            {
                case 2:
                    flipHorizontal = true;
                    break;
                case 3:
                    rotationDegrees = 180;
                    break;
                case 4:
                    rotationDegrees = 180;
                    flipHorizontal = true;
                    break;
                case 5:
                    rotationDegrees = 90;
                    flipHorizontal = true;
                    break;
                case 6:
                    rotationDegrees = 90;
                    break;
                case 7:
                    rotationDegrees = 270;
                    flipHorizontal = true;
                    break;
                case 8:
                    rotationDegrees = 270;
                    break;
            }
        }
    }
}
