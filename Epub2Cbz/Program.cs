using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace Epub2Cbz
{
    internal static class Program
    {
        private const string InputDir = @"C:\Documents\Downloads\Books\";
        private const string OutputDir = InputDir + @"out\";

        internal static async Task Main(string[] args)
        {
            if (!Directory.Exists(OutputDir))
                Directory.CreateDirectory(OutputDir);
            
            var inputFiles = Directory.GetFiles(InputDir, "*.epub", SearchOption.TopDirectoryOnly);
            Parallel.ForEach(
                inputFiles,
                //new() {MaxDegreeOfParallelism = 1},
                Convert
            );
            Console.WriteLine("Done.");
        }

        private static void Convert(string inputFile)
        {
            var logPrefix = Path.GetFileNameWithoutExtension(inputFile);
            Console.WriteLine($"{logPrefix}: Converting...");
            using var inputStream = File.Open(inputFile, FileMode.Open, FileAccess.Read, FileShare.Read);
            using var inputZip = new ZipArchive(inputStream, ZipArchiveMode.Read);
            var imageList = inputZip.Entries.Select(e => e.FullName).Where(n => Path.GetExtension(n.ToLower()) is ".jpg" or ".jpeg" or ".png").ToList();
            var baseBath = GetImageBasePath(imageList, logPrefix);
            if (baseBath is null)
            {
                Console.WriteLine($"{logPrefix}: NO IMAGES FOUND, skipping");
                return;
            }

            var coverPath = GetCoverPath(inputZip);
            if (coverPath is null)
                Console.WriteLine($"{logPrefix}: NO COVER FOUND");
            else
                imageList.Remove(coverPath);
            imageList.Sort(StringComparer.InvariantCultureIgnoreCase);
            using var outputStream = File.Open(Path.Combine(OutputDir, logPrefix + ".cbz"), FileMode.Create, FileAccess.Write, FileShare.Read);
            using var outputZip = new ZipArchive(outputStream, ZipArchiveMode.Create);
            if (coverPath is not null)
            {
                var cover = inputZip.Entries.First(e => e.FullName == coverPath);
                var outputCover = outputZip.CreateEntry("c_" + Path.GetFileName(coverPath), CompressionLevel.Optimal);
                using var inputCoverStream = cover.Open();
                using var outputCoverStream = outputCover.Open();
                inputCoverStream.CopyTo(outputCoverStream);
            }
            foreach (var img in imageList)
            {
                var inputEntry = inputZip.Entries.First(e => e.FullName == img);
                var outputEntry = outputZip.CreateEntry("i_" + Path.GetFileName(img), CompressionLevel.Optimal);
                using var inputEntryStream = inputEntry.Open();
                using var outputEntryStream = outputEntry.Open();
                inputEntryStream.CopyTo(outputEntryStream);
            }
            Console.WriteLine($"{logPrefix}: successfully converted to CBZ");
        }

        private static string? GetImageBasePath(List<string> imageList, string logPrefix)
        {
            var stats = new Dictionary<string, int>();
            foreach (var imagePath in imageList)
            {
                var dir = Path.GetDirectoryName(imagePath) ?? "";
                stats.TryGetValue(dir, out var count);
                stats[dir] = count + 1;
            }
            if (stats.Count == 0)
                return null;

            var top = stats.OrderByDescending(kvp => kvp.Value).First();
            Console.WriteLine($"{logPrefix}: top image folder is {top.Key} with {(top.Value*100.0/imageList.Count):0.}% of files");
            return top.Key;
        }

        private static string? GetCoverPath(ZipArchive zip)
        {
            var opf = zip.Entries.FirstOrDefault(e => e.Name.ToLower().EndsWith(".opf"));
            string? name = null;
            if (opf is not null)
            {
                using var stream = opf.Open();
                var xml = XDocument.Load(stream);
                var coverId = xml.Root?.Element("metadata")?.Elements("meta").FirstOrDefault(e => e.Attribute("name") is {Value: "cover"})?.Value;
                name = coverId is null ? null : xml.Root?.Element("manifest")?.Elements("item").FirstOrDefault(e => e.Attribute("id")?.Value == coverId)?.Attribute("href")?.Value;
                if (string.IsNullOrEmpty(name))
                    name = xml.Root?.Element("manifest")?.Elements("item").FirstOrDefault(e => e.Attribute("properties") is {Value: "cover-image"})?.Attribute("href")?.Value;
            }
            if (string.IsNullOrEmpty(name))
                name = zip.Entries.FirstOrDefault(e => e.Name.ToLower().Contains("cover") && (Path.GetExtension(e.Name.ToLower()) is ".jpg" or ".jpeg" or ".png"))?.Name;
            if (string.IsNullOrEmpty(name))
                return null;

            name = name.ToLower();
            return zip.Entries.First(e => e.Name.ToLower() == name).FullName;
        }
    }
}