using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToFileHashCsv
{
    class Program
    {
        static void Main(string[] args)
        {
            ArgumentValidator argValidator = new ArgumentValidator();
            try
            {
                // Validate the arguments length is right.
                argValidator.ValidateLength(args);
                // Validate the argument is not null or whitespace.
                String path = args[0];
                argValidator.ValidateNull(path);
                // Check if a Path is a File or a Direcory.
                // Check a Path is accessible or not.
                argValidator.ValidatePath(path);
                argValidator.Dispose();

                // Check existence of files in the directory.
                if (!ExcelUtils.Exists(path)) throw new Exception("指定されたパスに拡張子.xls? のファイルが見つかりませんでした。");

                String outputfullname =Path.Combine(Directory.GetCurrentDirectory(), DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv");
                using (StreamWriter writer = new StreamWriter(outputfullname, false, new UTF8Encoding(true)))
                {
                    String header = "\"fullname\",\"md5\"";
                    writer.WriteLine(header);
                    StringBuilder sBuilder = new StringBuilder(1024);
                    String quote = "\"";
                    String intermediate = "\",\"";
                    foreach (KeyValuePair<String, String> item in ExcelUtils.SortedFileHashes(path))
                    {
                        sBuilder.Append(quote);
                        sBuilder.Append(item.Key);
                        sBuilder.Append(intermediate);
                        sBuilder.Append(item.Value);
                        sBuilder.Append(quote);
                        writer.WriteLine(sBuilder.ToString());
                        sBuilder.Clear();
                    }
                }
            }
            catch (ArgumentException e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine(argValidator.ArgumentRequirement);
                argValidator.Dispose();
                Environment.Exit(1);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Environment.Exit(1);
            }
        }
        internal class ArgumentValidator
        {
            public String ArgumentRequirement { get; private set; } = "usage: ExcelToFileHashCsv.exe <引数はディレクトリの絶対パスまたはUNC パスです。引数にブランクが含まれる場合はダブルクォーテーションで囲んでください。>";
            // Validate the length of the argValidator
            internal void ValidateLength(string[] paramArray)
            {
                if (paramArray.Length != 1)
                {
                    string message = "引数が1 つではありません。";
                    throw new ArgumentException(message);
                }
            }
            internal void ValidateNull(string param)
            {
                if (String.IsNullOrWhiteSpace(param))
                {
                    string message = "引数がNull またはエンプティまたはブランクです。";
                    throw new ArgumentException(message);
                }
            }
            internal void ValidatePath(string path)
            {
                String pathRoot = Path.GetPathRoot(path);
                if (pathRoot == String.Empty)
                {
                    String message = "ルートディレクトリが取得できませんでした。引数が正しいか確認してください。";
                    throw new ArgumentException(message);
                }
                DirectoryInfo di = new DirectoryInfo(pathRoot);
                if (!di.Exists)
                {
                    String message = "引数のルートディレクトリを参照できませんでした。権限またはネットワークを確認してください。";
                    throw new ArgumentException(message);
                }
                FileAttributes attr = File.GetAttributes(path);
                if (!attr.HasFlag(FileAttributes.Directory))
                {
                    String message = "引数はディレクトリと認識されませんでした。";
                    throw new ArgumentException(message);
                }
                di = new DirectoryInfo(path);
                if (!di.Exists)
                {
                    String message = "引数のディレクトリを参照できませんでした。権限またはネットワークを確認してください。";
                    throw new ArgumentException(message);
                }
            }
            public void Dispose()
            {
                this.ArgumentRequirement = null;
            }
        }
        internal class ExcelUtils
        {
            internal static Boolean Exists(String path)
            {
                var file = Directory.GetFiles(path, "*.xls?", SearchOption.TopDirectoryOnly).FirstOrDefault();
                if (file == null) return false;
                return true;
            }
            internal static SortedList<String, String> SortedFileHashes(String path)
            {
                SortedList<String, String> fullname_hash = new SortedList<String, String>();
                Parallel.ForEach(
                    GetFileHashAsync(path),
                    new ParallelOptions { MaxDegreeOfParallelism = 4 },
                    res =>
                    {
                        lock (fullname_hash)
                        {
                            fullname_hash.Add(res.fullname, res.hash);
                        }
                    }
                );
                return fullname_hash;
            }
            private static IEnumerable<(String fullname, String hash)> GetFileHashAsync(String path)
            {
                FileStream res;
                MD5 hashAlgorithm = MD5.Create();
                Byte[] data;
                StringBuilder sBuilder = new StringBuilder(64);

                DirectoryInfo di = new DirectoryInfo(path);
                foreach (FileInfo fi in di.EnumerateFiles_Safety("*.xls?", (d, ex) => true))
                {
                    res = fi.OpenRead();
                    // md5sum
                    data = hashAlgorithm.ComputeHash(res);
                    for (int i = 0; i < data.Length; i++)
                    {
                        sBuilder.Append(data[i].ToString("x2")); // ToString("x2") turns Byte[] to Uppercase Hexadecimal String.
                    }
                    yield return (fi.FullName, sBuilder.ToString().ToUpper());
                    sBuilder.Clear();
                }
            }
        }
    }
    internal static class ExtensionDirectoryInfo
    {
        public static IEnumerable<FileInfo> EnumerateFiles_Safety(this DirectoryInfo directory,
            String searchPattern, Func<DirectoryInfo, Exception, bool> handleExceptionAccess)
        {
            IEnumerable<FileInfo> file_list = null;

            // Try to get an enumerator on fileSystemInfos of directory
            try
            {
                file_list = directory.EnumerateFiles(searchPattern, SearchOption.TopDirectoryOnly);
            }
            catch (Exception e)
            {
                //Console.WriteLine("\tInaccessible: {0}", directory);
                // If there's a callback delegate and this delegate return true, we don't throw the exception
                if (handleExceptionAccess == null || !handleExceptionAccess(directory, e))
                    throw;
                // If the exception wasn't throw, we make directories reference an empty collection
                file_list = Enumerable.Empty<FileInfo>();
            }

            foreach (FileInfo file in file_list)
            {
                yield return file;
            }
        }
    }
}
