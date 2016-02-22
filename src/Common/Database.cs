using System.IO;

namespace Common
{
    public class Database
    {
        public string FullName { get; private set; }
        public string Name { get { return Path.GetFileName(FullName); } }

        public Database(string fullName)
        {
            FullName = fullName;
        }
    }
}
