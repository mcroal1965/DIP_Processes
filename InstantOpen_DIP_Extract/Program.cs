using System;
using System.IO;
using System.Linq;
using System.Xml;
using System.Data;
using System.Data.SqlClient;

namespace InstantOpen_DIP_Extract
{
    class Program
    {
        public static string mappeddoctype { get; internal set; }
        

        static void Main(string[] args)
        {
            string paramfile = args[0];
            //string client = args[1];

            String ReaderName = null;

            String useDBServer = null;
            String useDatabase = null;
            String useTable = null;
            String useXMLTable = null;
            String useInPath = null;
            String useOutPath = null;
            String useDrivePath = null;

            String slash = Convert.ToString(Convert.ToChar(92));  //store the slash so it can be used in the filename later

            XmlTextReader reader = new XmlTextReader(paramfile);  // store each line of the input xml file into reader

            while (reader.Read())  //process the rows until no more
            {
                switch (reader.NodeType)
                {
                    case XmlNodeType.Element:  //store the name of all node elements into ReaderName
                        ReaderName = reader.Name;
                        break;
                    case XmlNodeType.Text:
                        if (ReaderName is "DBServer")
                        { useDBServer = reader.Value; }

                        if (ReaderName is "Database")
                        { useDatabase = reader.Value; }

                        if (ReaderName is "Table")
                        { useTable = reader.Value; }

                        if (ReaderName is "XMLTable")
                        { useXMLTable = reader.Value; }

                        if (ReaderName is "InPath")
                        { useInPath = reader.Value; }

                        if (ReaderName is "OutPath")
                        { useOutPath = reader.Value; }
                        
                        if (ReaderName is "DrivePath")
                        { useDrivePath = reader.Value; }
                        break;
                }
            }
            //Get all items in all folder and subfolders
            String[] allitems = Directory.GetFiles(useInPath, "*.*", SearchOption.AllDirectories);
            Int32 numitems = allitems.Count();
            Int32 pathlength = useInPath.Length;
            String Docdate;
            Int32 workingitemnum = 0;

            foreach (string item in allitems)
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

                //handle the not xml files first
                if (filenamextension != ".xml")
                {
                    string[] splittext = filename.Split("_");
                    String custname = splittext[0];
                    String ssn = splittext[1];
                    String acctnum = splittext[2];
                    String tranid = splittext[3];
                    String doctype = splittext[4];
                    
                    String docdate = splittext[5];
                    Docdate = docdate.Substring(0, 2) + "/" + docdate.Substring(2, 2) + "/" + docdate.Substring(4, 4);
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
                            connection.Open();  //open the sql server connection to the database
                            var dbreader = cmd.ExecuteReader();  //run the command and put the results into dbreader
                            
                            //if the reader has rows get the mapper doctype from the table
                            while (dbreader.Read())
                            {
                                string NautilusDocType = dbreader.GetValue(dbreader.GetOrdinal("NautilusDoctype")).ToString();
                                mappeddoctype = NautilusDocType;
                            }
                            connection.Close();  //close the sql server connection to the database
                            //default this keyword to nil because if the document mapping doesnt exist we'll put the name from the file into description kw 
                            //for DIP because these will map to DEP Disclosure and the workflow will assign Description to the TYP kw
                            String Description = "";
                            String DIPDoctype = "";
                            
                            if (mappeddoctype == "")
                            {
                                DIPDoctype = "DEP Disclosure";
                                Description = doctype;
                            }
                            else
                            {
                                DIPDoctype = mappeddoctype;
                            }

                            String outDIPindexfile = "DIPindex_" + "_" + filename + ".txt".Replace(" ", "");  //the name of the index file to be used for this file

                            //build the line for the index file
                            String DIPIndexValue = DIPDoctype + "\t" + Docdate + "\t" + acctnum + "\t" + custname + "\t" + ssn + "\t" + tranid + "\t" + Description + "\t" + useDrivePath + slash + filenamewithextension;

                            File.WriteAllText(useOutPath + slash + outDIPindexfile, DIPIndexValue);
                            File.Copy(fullpathfilename, useOutPath+slash+filenamewithextension, true);
        //                    File.Delete(fullpathfilename);

                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Error: " + ex);
                        }
                    }
                }
                //handle the xml files
                if (filenamextension == ".xml")
                {

                }
        
            }
        }
    }
}
