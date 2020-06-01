using System;
using System.IO;

namespace DeleteOutdatedFiles
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string d = args[0];
                string r = args[1].ToUpper();
                if (r != "Y" && r != "N")
                {
                    Console.WriteLine("Incorrect Format, must contain 3 arguments: pathpattern YN 99");
                    Console.WriteLine("Press any key to exit.");
                    Console.ReadKey();
                    Environment.Exit(0);
                }

                string z = args[2];
                Int32 days = Int32.Parse(args[2]);

                String slash = Convert.ToString(Convert.ToChar(92));  //store the slash so it can be used in the filename later

                string filename = d.Substring(d.LastIndexOf(slash) + 1, d.Length - d.LastIndexOf(slash) - 1);
                if (filename.Contains("."))
                {
                    Console.WriteLine("Processing folder " + d);
                }
                else
                {
                    Console.WriteLine("Incorrect Format, must end with *.* or some flavor.");
                    Console.WriteLine("Press any key to exit.");
                    Console.ReadKey();
                    Environment.Exit(0);
                }

                DirSearch(d.Substring(0, d.LastIndexOf(slash)), filename, r, days);

            }
            catch
            {
                Console.WriteLine("Incorrect Format, must contain 3 arguments: pathpattern YN 99");
                Console.WriteLine("Press any key to exit.");
                Console.ReadKey();
                Environment.Exit(0);
            }
            finally
            {
                Environment.Exit(0);
            }
        }

        static void DirSearch(string sDir, string filename, string recurse, int days)
        {
            try
            {
                foreach (string f in Directory.GetFiles(sDir, filename))
                {
                    DateTime filedate = File.GetLastWriteTime(f);
                    TimeSpan difference = DateTime.Today - filedate;

                    if (difference.TotalDays > days)
                    {
                        Console.WriteLine("Will be deleted: " + f);
                        File.Delete(f);
                    }
                }

                if (recurse == "Y")
                {
                    foreach (string d in Directory.GetDirectories(sDir))
                    {
                        foreach (string f in Directory.GetFiles(d, filename))
                        {
                            DateTime filedate = File.GetLastWriteTime(f);
                            TimeSpan difference = DateTime.Today - filedate;

                            if (difference.TotalDays > days)
                            {
                                Console.WriteLine("Will be deleted: " + f);
                                File.Delete(f);
                            }
                            Console.WriteLine("Will be deleted: " + f);
                        }
                        DirSearch(d, filename, recurse, days);
                    }
                }
            }
            catch (System.Exception excpt)
            {
                Console.WriteLine(excpt.Message);
                Console.WriteLine("Press any key to exit.");
                Console.ReadKey();
                Environment.Exit(1);
            }
        }
    }
}

