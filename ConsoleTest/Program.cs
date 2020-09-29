using ItAintBoring.Functions;
using System;
using System.IO;

namespace ConsoleTest
{
    class Program
    {
        static void Main(string[] args)
        {

            var streamFile1 = TestDocStorage.GetTestFile("File1.docx");

            var streamFile2 = TestDocStorage.GetTestFile("File2.docx");
            DocumentHandler dh = new DocumentHandler();

            //var resultStream = new MemoryStream(0);
            //streamFile1.CopyTo(resultStream);
            dh.Merge(streamFile1, streamFile2);
            string resultFile = TestDocStorage.GetFilePath("result.docx");
            if (File.Exists(resultFile)) File.Delete(resultFile);
            File.WriteAllBytes(resultFile, streamFile1.ToArray());
            //TestDocStorage.SaveResult(resultStream, "Result.docx");

        }
    }
}
