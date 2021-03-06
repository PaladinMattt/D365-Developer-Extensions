using System.Collections.Generic;

namespace PluginDeployer.Spkl
{
    public interface IDirectoryService
    {
        string GetApplicationDirectory();
        string SimpleSearch(string path, string search);
        List<string> Search(string path, string search);
        void SaveFile(string filename, byte[] content, bool overwrite);
    }
}