using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;

namespace XlsxGateway.SystemAdapters
{
    public class SystemCompressor : ICompressor
    {
        private string fileName;
        private Dictionary<string, CompressedEntry> entries = new Dictionary<string, CompressedEntry>();

        public void OpenForUpdate(string fileName)
        {
            Open(fileName, ZipArchiveMode.Update);
        }

        public void OpenToRead(string fileName)
        {
            Open(fileName, ZipArchiveMode.Read);
        }

        public void Save()
        {
            string tempFolderName = fileName + ".TempDir";

            if (Directory.Exists(tempFolderName))
                Directory.Delete(tempFolderName, recursive: true);

            Directory.CreateDirectory(tempFolderName);

            foreach (var pair in entries)
            {
                string path = Path.Combine(tempFolderName, pair.Key.ToString());
                string dir = Path.GetDirectoryName(path);
                if (!Directory.Exists(dir))
                    Directory.CreateDirectory(dir);

                var stream = File.CreateText(path);
                stream.Write(pair.Value.Content);
                stream.Flush();
                stream.Close();
            }

            if (File.Exists(fileName))
                File.Delete(fileName);

            ZipFile.CreateFromDirectory(tempFolderName, fileName);

            Directory.Delete(tempFolderName, recursive: true);
        }

        public CompressedEntry GetEntry(string name)
        {
            CompressedEntry value;
            string path = Normalize(name);
            if (entries.TryGetValue(path, out value))
                return value;

            throw new Exception("Archive entry does not exist named: " + path);
        }

        private void Open(string fileName, ZipArchiveMode mode)
        {
            entries.Clear();
            this.fileName = fileName;
            using (var archive = ZipFile.Open(fileName, mode))
            {
                foreach (var archiveEntry in archive.Entries)
                    AddEntry(archiveEntry);
            }
        }

        private void AddEntry(ZipArchiveEntry archiveEntry)
        {
            string path = Normalize(archiveEntry.FullName);

            entries.Add(path, EntryFrom(archiveEntry));
        }

        private static string Normalize(string path)
        {
            return path == null ? null : path.Replace('\\', '/');
        }

        private static string StringFrom(Stream sourceStream)
        {
            string content = new StreamReader(sourceStream)
                .ReadToEnd();

            sourceStream.Close();

            return content;
        }

        private static CompressedEntry EntryFrom(ZipArchiveEntry sourceEntry)
        {
            return new CompressedEntry
            {
                Name = sourceEntry.FullName,
                Content = StringFrom(sourceEntry.Open())
            };
        }
    }
}
