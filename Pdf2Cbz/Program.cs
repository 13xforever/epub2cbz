using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;
using NReco.PdfRenderer;

namespace Pdf2Cbz
{
    internal static class Program
    {
        private const string InputDir = @"C:\Documents\Downloads\Books\";
        private const string OutputDir = InputDir + @"out\";
        private const string toolPath = @"PdfRenderer\pdftoppm.exe";

        internal static async Task Main(string[] args)
        {
            if (!Directory.Exists(OutputDir))
                Directory.CreateDirectory(OutputDir);
            
            var inputFiles = Directory.GetFiles(InputDir, "*.pdf", SearchOption.TopDirectoryOnly);
            inputFiles = inputFiles.OrderByDescending(fn => new FileInfo(fn).Length).ToArray();
            CheckPdfToPpm(inputFiles.Last());
            Parallel.ForEach(
                inputFiles,
                //new() {MaxDegreeOfParallelism = 1},
                Convert
            );
            Console.WriteLine("Done.");
        }

        private static void CheckPdfToPpm(string inputFile)
        {
            if (File.Exists(toolPath))
                return;

            var pdf2image = new PdfToImageConverter();
            using var tmp = new MemoryStream();
            pdf2image.GenerateImage(inputFile, 1, ImageFormat.Jpeg, tmp);
            if (!File.Exists(toolPath))
                throw new InvalidOperationException("Couldn't fine pdftoppm.exe");
        }

        private static void Convert(string inputFile)
        {
            var prefix = Path.GetFileNameWithoutExtension(inputFile);
            var tmpFolder = Path.Combine(Path.GetTempPath(), "pdf2cbz", Guid.NewGuid().ToString("n"));
            Directory.CreateDirectory(tmpFolder);
            try
            {
                var psi = new ProcessStartInfo(toolPath, $"-jpeg -jpegopt quality=97,optimize=y -r 300 -cropbox \"{inputFile}\" i")
                {
                    WorkingDirectory = tmpFolder,
                    WindowStyle = ProcessWindowStyle.Hidden,
                };
                Console.WriteLine($"{prefix}: Rendering pdf to png...");
                using var tool = Process.Start(psi);
                if (tool is null)
                {
                    Console.WriteLine($"{prefix}: failed to start pdftoppm.exe");
                    return;
                }
                
                tool.WaitForExit();
                Console.WriteLine($"{prefix}: Building cbz...");
                var imgCount = 0;
                using var outputStream = File.Open(Path.Combine(OutputDir, prefix + ".cbz"), FileMode.Create, FileAccess.Write, FileShare.Read);
                using var zip = new ZipArchive(outputStream, ZipArchiveMode.Create);
                foreach (var page in Directory.GetFiles(tmpFolder, "*.jpg", SearchOption.TopDirectoryOnly))
                {
                    using var imgStream = File.Open(page, FileMode.Open, FileAccess.Read, FileShare.Read);
                    var entry = zip.CreateEntry($"{imgCount++:0000}.jpg", CompressionLevel.Optimal);
                    using var entryStream = entry.Open();
                    imgStream.CopyTo(entryStream);
                }
                Console.WriteLine($"{prefix}: Completed");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"{prefix}: Failed to convert: " + ex);
            }
            finally
            {
                Directory.Delete(tmpFolder, true);
            }
        }
    }
}