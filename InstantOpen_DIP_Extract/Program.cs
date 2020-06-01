using System;
using System.IO;
using System.Linq;
using System.Xml;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;

namespace InstantOpen_DIP_Extract
{
    class Program
    {
        public static string mappeddoctype { get; internal set; }

        static void Main()
        {
            try
            {
                String useDBServer = ConfigurationManager.AppSettings["dbserver"].ToString();
                String useDatabase = ConfigurationManager.AppSettings["database"].ToString();
                String useTable = ConfigurationManager.AppSettings["table"].ToString();
                String useXMLTable = ConfigurationManager.AppSettings["xmltable"].ToString();
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
                String[] allitems = Directory.GetFiles(useInPath, "*.*", SearchOption.AllDirectories);
                Int32 numitems = allitems.Count();
                Int32 pathlength = useInPath.Length;
                String Docdate;
                Int32 workingitemnum = 0;

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
                        String custname = splittext[0];
                        String ssn = splittext[1];
                        String acctnum = splittext[2];
                        String tranid = splittext[3];
                        String doctype = splittext[4];
                        String docdate = splittext[5];
                        Docdate = docdate.Substring(0, 2) + "/" + docdate.Substring(2, 2) + "/" + docdate.Substring(4, 4);

                        mappeddoctype = "";  //reset to assume not mapped

                        String sqlCmd = "SELECT TOP 1 a.NautilusDoctype FROM  " + useTable + " a WHERE a.OnlineBankingDoctype='" + doctype + "'";
                        String connectionString = "Server=" + useDBServer + ";Database=" + useDatabase + ";User Id=viewer;Password=cprt_hsi";

                        using (SqlConnection connection = new SqlConnection(connectionString))  //connect to the sql server
                        using (SqlCommand cmd = connection.CreateCommand())  //start a  sql command
                        {
                            try
                            {
                                //see if the document name extracted from the filename is in the mapping table
                                cmd.CommandText = sqlCmd;  //set the commandtext to the sqlcmd
                                cmd.CommandType = CommandType.Text;  //set it as a text command
                                try
                                {
                                    connection.Open();  //open the sql server connection to the database
                                }
                                catch
                                {
                                    Console.WriteLine("SQL Server not available.");
                                    Console.WriteLine("Press any key to exit.");
                                    Console.ReadKey();
                                    Environment.Exit(1);
                                }
                                var dbreader = cmd.ExecuteReader();  //run the command and put the results into dbreader

                                //if the reader has rows get the mapper doctype from the table
                                while (dbreader.Read())
                                {
                                    string NautilusDocType = dbreader.GetValue(dbreader.GetOrdinal("NautilusDoctype")).ToString();
                                    mappeddoctype = NautilusDocType;
                                }
                                connection.Close();  //close the sql server connection to the database                      

                                String Description = ""; //default this keyword to nil because if the document mapping doesnt exist we'll put the name from the file into description kw 
                                String DIPDoctype = ""; //for DIP because these will map to DEP Disclosure and the workflow will assign Description to the TYP kw

                                if (mappeddoctype == "")
                                {
                                    DIPDoctype = "DEP Disclosure";
                                    Description = doctype;
                                }
                                else
                                {
                                    DIPDoctype = mappeddoctype;
                                }

                                if (filenamextension == ".pdf")
                                {
                                    filetype = "16";
                                }
                                else
                                {
                                    filetype = "2";
                                }
                                String outDIPindexfile = "DIPindex_" + "_" + filename + ".txt".Replace(" ", "");  //the name of the index file to be used for this file     
                                String DIPIndexValue = DIPDoctype + "\t" + Docdate + "\t" + acctnum + "\t" + custname + "\t" + ssn + "\t" + tranid + "\t" + Description + "\t" + useDrivePath + slash + filenamewithextension + "\t" + filetype; //build the line for the index file

                                //create the DIPIndex file
                                File.WriteAllText(useOutPath + slash + outDIPindexfile, DIPIndexValue);
                                //copy the source file to the folder with the DIPIndex file
                                File.Copy(fullpathfilename, useOutPath + slash + filenamewithextension, true);
                                File.Copy(fullpathfilename, useBackupPath + slash + filenamewithextension, true);
                                File.Delete(fullpathfilename);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("Error: " + ex);
                                Console.WriteLine("Press any key to exit.");
                                Console.ReadKey();
                                Environment.Exit(1);
                            }
                        }
                    }

                    if (filenamextension == ".xml") //handle the xml files
                    {
                        String[] splittext = filename.Split("_");
                        String xmldesc = splittext[1];

                        filename = item.Substring(pathlength + 1, filenamelength - filenamextensionlength);
                        filenamewithextension = filename + filenamextension;

                        String AppID = "";

                        XmlDocument XmlDoc = new XmlDocument();
                        XmlDoc.Load(fullpathfilename);
                        XmlNodeList elemList = XmlDoc.GetElementsByTagName("Form");
                        for (int i = 0; i < elemList.Count; i++)
                        {
                            AppID = elemList[i].Attributes["FormNo"].Value;
                            if (AppID != "")
                            {
                                String sqlCmd = "INSERT INTO " + useXMLTable + "(XMLFileName, ApplicationID) VALUES ('" + filename + "', '" + AppID + "')";
                                String connectionString = "Server=" + useDBServer + ";Database=" + useDatabase + ";User Id=viewer;Password=cprt_hsi";

                                using (SqlConnection connection2 = new SqlConnection(connectionString))  //connect to the sql server
                                using (SqlCommand cmd2 = connection2.CreateCommand())  //start a  sql command
                                {
                                    try
                                    {
                                        cmd2.CommandText = sqlCmd;  //set the commandtext to the sqlcmd
                                        cmd2.CommandType = CommandType.Text;  //set it as a text command
                                        try
                                        {
                                            connection2.Open();  //open the sql server connection to the database
                                        }
                                        catch
                                        {
                                            Console.WriteLine("SQL Server not available.");
                                            Console.WriteLine("Press any key to exit.");
                                            Console.ReadKey();
                                            Environment.Exit(1);
                                        }
                                        int rowsadded = cmd2.ExecuteNonQuery();  //run the command and store the row count inserted
                                        connection2.Close();  //close the sql server connection to the database
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine("Error: " + ex);
                                        Console.WriteLine("Press any key to exit.");
                                        Console.ReadKey();
                                        Environment.Exit(1);
                                    }
                                } 
                            }
                        }
                        Docdate = filename.Substring(filenamelength - filenamextensionlength - 4, 2) + "/" + filename.Substring(filenamelength - filenamextensionlength - 2, 2) + "/" + filename.Substring(filenamelength - filenamextensionlength - 8, 4);

                        String outDIPindexfile = "DIPindex_" + "_" + filename + ".txt".Replace(" ", "");  //the name of the index file to be used for this file
                        String DIPIndexValue = "DEP Instant Open XML " + "\t" + Docdate + "\t" + xmldesc + "\t" + filename + "\t" + useDrivePath + slash + filenamewithextension + "\t" + "32"; //build the line for the index file

                        File.WriteAllText(useOutPath + slash + outDIPindexfile, DIPIndexValue);
                        File.Copy(fullpathfilename, useOutPath + slash + filenamewithextension, true);
                        File.Copy(fullpathfilename, useBackupPath + slash + filenamewithextension, true);
                        File.Delete(fullpathfilename);
                    }
                }
            }
            catch
            {
                Console.WriteLine("Config file does not exist or does not meet requirements.");
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
