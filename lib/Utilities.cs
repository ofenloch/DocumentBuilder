namespace dblib
{
    public class Utilities
    {
        // create the given directory / path
        static public string CreateDirectory(string path)
        {
            path = Path.GetFullPath(path);
            try
            {
                // Determine whether the directory exists.
                if (Directory.Exists(path))
                {
                    // TODO: use logger Console.WriteLine("That path \"{0}\" exists already.", path);
                    return path;
                }
                // Try to create the directory.
                DirectoryInfo di = Directory.CreateDirectory(path);
                // TODO: use logger Console.WriteLine("The directory \"{0}\" was created successfully at {1}.", di.FullName, Directory.GetCreationTime(path));
                return path;
            }
            catch (Exception e)
            {
                Console.WriteLine("CreateDirectory failed: {0}", e.ToString());
                return "";
            }
            finally { }
        } // static public string CreateDirectory(string path)

    } // public class Utilities

} // namespace dblib