using System;
using System.Collections.Generic;
using System.IO;

namespace Common
{
    public class AccessFileProvider
    {
        public List<FileInfo> Files { get; private set; }
        public string RootDirectory { get; private set; }

        public AccessFileProvider(string rootDirectory)
        {
            if (!Directory.Exists(rootDirectory))
                throw new ArgumentException(
                    string.Format("Directory does not exist: ", rootDirectory));

            RootDirectory = rootDirectory;
            Files = new List<FileInfo>();
            EnumerateFiles();
        }

        private void EnumerateFiles()
        {
            var dirInfo = new DirectoryInfo(RootDirectory);

            Files.AddRange(dirInfo.EnumerateFiles("*.mdb", SearchOption.AllDirectories));
            Files.AddRange(dirInfo.EnumerateFiles("*.accdb", SearchOption.AllDirectories));
        }
    }
}
