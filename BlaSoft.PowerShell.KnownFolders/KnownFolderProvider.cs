namespace BlaSoft.PowerShell.KnownFolders
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.IO;
    using System.Linq;
    using System.Management.Automation;
    using System.Management.Automation.Provider;

    [CmdletProvider("KnownFolders", ProviderCapabilities.None)]
    public sealed class KnownFolderProvider : ContainerCmdletProvider, IContentCmdletProvider
    {
        private readonly GetKnownFolderCommand getKnownFolderCommand = new GetKnownFolderCommand();

        public void ClearContent(string path)
        {
            throw new NotSupportedException();
        }

        public object ClearContentDynamicParameters(string path)
        {
            return null;
        }

        public IContentReader GetContentReader(string path)
        {
            return new KnownFolderContentReader(path, this);
        }

        public object GetContentReaderDynamicParameters(string path)
        {
            return null;
        }

        public IContentWriter GetContentWriter(string path)
        {
            throw new NotSupportedException();
        }

        public object GetContentWriterDynamicParameters(string path)
        {
            return null;
        }

        protected override bool IsValidPath(string path)
        {
            return !string.IsNullOrEmpty(path);
        }

        protected override void GetChildNames(string path, ReturnContainers returnContainers)
        {
            if (string.IsNullOrEmpty(path))
            {
                var items = this.GetItemsDictionary();
                foreach (var kvp in items.OrderBy(kvp => kvp.Key, StringComparer.CurrentCultureIgnoreCase))
                {
                    this.WriteItemObject(kvp.Key, kvp.Key, isContainer: false);
                }
            }
            else
            {
                var kf = this.TryGetItemByName(path);
                if (kf != null)
                {
                    this.WriteItemObject(path, path, isContainer: false);
                }
            }
        }

        protected override void GetChildItems(string path, bool recurse)
        {
            if (string.IsNullOrEmpty(path))
            {
                var items = this.GetItemsDictionary();
                foreach (var kvp in items.OrderBy(kvp => kvp.Key, StringComparer.CurrentCultureIgnoreCase))
                {
                    //var obj = new KeyValuePair<string, string>(kvp.Key, kvp.Value.Path);
                    var obj = new DictionaryEntry(kvp.Key, kvp.Value.Path);
                    this.WriteItemObject(obj, kvp.Key, isContainer: false);
                }
            }
            else
            {
                var kf = this.TryGetItemByName(path);
                if (kf != null)
                {
                    var obj = new DictionaryEntry(kf.Name, kf.Path);
                    this.WriteItemObject(obj, path, isContainer: false);
                }
            }
        }

        protected override void GetItem(string path)
        {
            if (string.IsNullOrEmpty(path))
            {
                var items = this.GetItemsDictionary();
                var obj = items
                    .OrderBy(kvp => kvp.Key, StringComparer.CurrentCultureIgnoreCase)
                    .Select(kvp => new DictionaryEntry(kvp.Key, kvp.Value.Path))
                    .ToList();
                this.WriteItemObject(obj, path, isContainer: true);
            }
            else
            {
                var kf = this.TryGetItemByName(path);
                if (kf != null)
                {
                    var obj = new DictionaryEntry(kf.Name, kf.Path);
                    this.WriteItemObject(obj, path, isContainer: false);
                }
            }
        }

        protected override bool HasChildItems(string path)
        {
            return string.IsNullOrEmpty(path) && 0 < this.GetItemsDictionary().Count;
        }

        protected override bool ItemExists(string path)
        {
            if (string.IsNullOrEmpty(path))
            {
                return true;
            }

            var kf = this.TryGetItemByName(path);
            return kf != null;
        }

        protected override Collection<PSDriveInfo> InitializeDefaultDrives()
        {
            return new Collection<PSDriveInfo>()
            {
                new PSDriveInfo("KF", this.ProviderInfo, string.Empty, "Known Folders", null)
            };
        }

        private IDictionary<string, KnownFolder> GetItemsDictionary()
        {
            var dict = this.getKnownFolderCommand.GetAll()
                .Select(nkf => new KnownFolder(nkf))
                .Select(kf => new KeyValuePair<string, KnownFolder>(kf.Name, kf))
                .Where(kvp => !string.IsNullOrEmpty(kvp.Key))
                .ToDictionary(kvp => kvp.Key, kvp => kvp.Value, StringComparer.CurrentCultureIgnoreCase);
            return dict;
        }

        private KnownFolder TryGetItemByName(string name)
        {
            try
            {
                var nkf = this.getKnownFolderCommand.GetKnownFolderByName(name);
                return new KnownFolder(nkf);
            }
            catch (FileNotFoundException)
            {
                return null;
            }
        }

        private sealed class KnownFolderContentReader : IContentReader
        {
            private readonly string path;

            private readonly KnownFolderProvider provider;

            private bool contentRead;

            public KnownFolderContentReader(string path, KnownFolderProvider provider)
            {
                this.path = path;
                this.provider = provider;
            }

            public void Close()
            {
            }

            public IList Read(long readCount)
            {
                IList result = null;
                if (!this.contentRead)
                {
                    var item = provider.TryGetItemByName(this.path);
                    if (item != null)
                    {
                        result = new[] { item.Path };
                    }

                    this.contentRead = true;
                }

                return result;
            }

            public void Seek(long offset, SeekOrigin origin)
            {
                throw new NotSupportedException();
            }

            public void Dispose()
            {
                this.Close();
                GC.SuppressFinalize(this);
            }
        }
    }
}
