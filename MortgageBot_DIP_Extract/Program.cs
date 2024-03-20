using System;
using System.IO;
using System.Linq;
using System.Configuration;

namespace MortgageBot_DIP_Extract
{
    class Program
    {
        public static string mappeddoctype { get; internal set; }

        static void Main()
        {
            try
            {               
                String useInPath = ConfigurationManager.AppSettings["inpath"].ToString();
                String useOutPath = ConfigurationManager.AppSettings["outpath"].ToString();
                String useDrivePath = ConfigurationManager.AppSettings["drivepath"].ToString();
                String useBackupPath = ConfigurationManager.AppSettings["backuppath"].ToString();
                String filetype = "2";  //default to image file format

                try
                {
                    Directory.CreateDirectory(useBackupPath);
                }
                catch
                {
                    Console.WriteLine("Directory " + useBackupPath + " already exists.");
                }

                try
                {
                    Directory.CreateDirectory(useOutPath);
                }
                catch
                {
                    Console.WriteLine("Directory " + useOutPath + " already exists.");
                }

                String slash = Convert.ToString(Convert.ToChar(92));  //store the slash so it can be used in the filename later

                //Get all items in all folder and subfolders
                String[] allitems = Directory.GetFiles(useInPath, "*.*", SearchOption.TopDirectoryOnly);
                Int32 numitems = allitems.Count();
                Int32 pathlength = useInPath.Length;
                Int32 workingitemnum = 0;
                string datetimestamp = DateTime.Now.ToString("yyyyMMddHHmm");

                foreach (String item in allitems)
                {
                    ++workingitemnum;
                    Console.WriteLine("Working: " + workingitemnum + " of " + numitems);
                    FileInfo f = new FileInfo(item);

                    Int32 filenamelength = item.Length - pathlength;  //length of the filename+ext minus the fullpath to it
                    String filenamextension = Path.GetExtension(item);
                    Int32 filenamextensionlength = filenamextension.Length + 1;
                    String filename = item.Substring(pathlength + 1, filenamelength - filenamextensionlength);
                    String filenamewithextension = filename + filenamextension;
                    String fullpathfilename = item;

                    if (filenamextension != ".xml") //handle the not xml files first
                    {
                        string[] splittext = filename.Split("_");
                        String appid = splittext[0].Trim();
                        String doctype = splittext[1].Trim();

                        if (filenamextension == ".pdf")
                        {
                            filetype = "16";
                        }
                        else
                        {
                            filetype = "2";
                        }

                        //String outDIPindexfile = ("DIPindex_MBot_" + filename + ".txt").Replace(" ", "");  //the name of the index file to be used for this file     
                        String outDIPindexfile = ("DIPindex_MBot_" + datetimestamp + ".txt").Replace(" ", "");  //the name of the index file to be used for this file     

                        String DIPIndexValue = doctype + "\t" + appid + "\t" + useDrivePath + slash + filenamewithextension + "\t" + filetype + "\t" + "MortgageBot" + "\r\n"; //build the line for the index file

                        //create the DIPIndex file
                        //File.WriteAllText(useOutPath + slash + outDIPindexfile, DIPIndexValue);
                        File.AppendAllText(useOutPath + slash + outDIPindexfile, DIPIndexValue);

                        //copy the source file to the folder with the DIPIndex file
                        File.Copy(fullpathfilename, useOutPath + slash + filenamewithextension, true);
                        File.Copy(fullpathfilename, useBackupPath + slash + filenamewithextension, true);
                        File.Delete(fullpathfilename);
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Config file does not exist or does not meet requirements.");
                Console.WriteLine(e.InnerException == null ? e.Message: e.InnerException.Message);
                Console.WriteLine("Press any key to exit.");
                Console.ReadKey();
                Environment.Exit(1);
            }
            finally
            {
                Environment.Exit(0);
            }
        }
    }
}

