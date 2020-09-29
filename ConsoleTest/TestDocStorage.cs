using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ConsoleTest
{
    public class TestDocStorage
    {
        public static MemoryStream GetTestFile(string fileName)
        {
            string path = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "../../../../../TestFiles/" + fileName;
            MemoryStream result = new MemoryStream();
            var fileData = File.ReadAllBytes(path);
            result.Write(fileData, 0, fileData.Length);
            return result;
        }

        public static string GetFilePath(string fileName)
        {
            string path = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "../../../../../TestFiles/" + fileName;
            return path;    
        }
        public static void SaveResult(MemoryStream resultStream, string fileName)
        {
            string path = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "../../../../../TestFiles/" + fileName;
            resultStream.Position = 0;
            using (var sw = new System.IO.FileStream(path, FileMode.Create)) {
                resultStream.WriteTo(sw);
            }
        }
    }
}
