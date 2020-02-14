using System;
using System.IO;
using System.Linq;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;

namespace ASB_COLD_DIP_Process
{
    public class Program
    {
        public static void Main()
        { 
            try
            {
                String useDBServer = ConfigurationManager.AppSettings["dbserver"].ToString();
                String useDatabase = ConfigurationManager.AppSettings["database"].ToString();
                String useTable = ConfigurationManager.AppSettings["table"].ToString();
                String useNoteTable = ConfigurationManager.AppSettings["notetable"].ToString();
                String useStatsTable = ConfigurationManager.AppSettings["statstable"].ToString();
                String useInPath = ConfigurationManager.AppSettings["inpath"].ToString();
                String useOutPath = ConfigurationManager.AppSettings["outpath"].ToString();
                try
                    {
                        Directory.CreateDirectory(useOutPath);
                    }
                catch
                    {
                        Console.WriteLine("Directory " + useOutPath + " already exists.");
                    }
                String useDrivePath = ConfigurationManager.AppSettings["drivepath"].ToString();
                String slash = Convert.ToString(Convert.ToChar(92));  //store the slash so it can be used in the filename later

                String connectionString = "Server=" + useDBServer + ";Database=" + useDatabase + ";User Id=viewer;Password=cprt_hsi";

                //Get date items which are the folder names in
                Int32 datecount = 0;
                String[] DateItems = Directory.GetDirectories(useInPath, "*.");
                Int32 numdateitems = DateItems.Count();

                foreach (string dateitem in DateItems)
                {
                    FileInfo f = new FileInfo(dateitem);
                    String pathdate = f.Name;  //the name of the date folder only eg. 20190512
                    String datepath = dateitem;   //the fullpath of the date folder eg. H:\ASB_COLD\20190512
                    Int32 pathlength = datepath.Length;  //the length of the full path; used to strip that off the filename later

                    ++datecount; // increment 

                    String Docdate = datepath.Substring(useInPath.Length + 1 + 4, 2) + "/" + datepath.Substring(useInPath.Length + 1 + 6, 2) + "/" + datepath.Substring(useInPath.Length + 1, 4);

                    //Get all txt items for date folder
                    String[] txtitems = Directory.GetFiles(datepath, "*.txt");
                    Int32 numtxtitems = txtitems.Count();

                    //Get all not items for date folder
                    String[] notitems = Directory.GetFiles(datepath, "*.not", SearchOption.AllDirectories);
                    Int32 numnotitems = notitems.Count();

                    Int32 txtcount = 0;
                    Int32 notcount = 0;

                    String sqlCmd0 = "Insert Into " + useStatsTable + " ([RptDate],[NumReports],[NumNotes],[DIPStatus]) values ('" + pathdate + "', '" + numtxtitems + "', '" + numnotitems + "', 'PROCESSING')";
                    using (SqlConnection connection0 = new SqlConnection(connectionString))  //connect to the sql server
                    using (SqlCommand cmd0 = connection0.CreateCommand())  //start a  sql command
                    try
                    {
                        cmd0.CommandText = sqlCmd0;  //set the commandtext to the sqlcmd
                        cmd0.CommandType = CommandType.Text;  //set it as a text command
                        connection0.Open();  //open the sql server connection to the database
                        int rowsadded = cmd0.ExecuteNonQuery();  //run the command and store the row count inserted
                        connection0.Close();  //close the sql server connection to the database
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error inserting New Date into Stats: " + ex);
                    }

                    foreach (string txtitem in txtitems)  //txtitem will contain the fullpath to .txt file
                    {
                        Int32 filenamelength = txtitem.Length - pathlength;  //length of the filename+ext minus the fullpath to it
                        String filenamextension = Path.GetExtension(txtitem);
                        Int32 filenamextensionlength = filenamextension.Length + 1;
                        String filename = txtitem.Substring(pathlength + 1, filenamelength - filenamextensionlength);
                        String filenamewithextension = filename + filenamextension;
                        String fullpathfilename = txtitem;
                        String Unknown = "NO";

                        //Console.WriteLine(txtitem + " | " + itemin + " | " + inx + " | " + ifn + " | ");

                        String RPTnumber = filenamewithextension.Substring(0, filenamewithextension.IndexOf("."));
                        ++txtcount;
                        String RPTinstance = filenamewithextension;
                        Console.Write("Working Date Folder: " + pathdate + " # " + datecount + " of " + numdateitems + " | File: " + filename + " " + txtcount + " of " + numtxtitems + " | Note: " + notcount + " of " + numnotitems + "\n");

                        String sqlCmd = "SELECT TOP 1 a.Application, a.ReportTitle, cast(a.RptRetentionDays as varchar(5)) FROM " + useTable + " a WHERE a.ReportNumber='" + RPTnumber + "'";
                        DataTable dt = new DataTable();  //declare a datable to hold the results
                        int rowsreturned;

                        using (SqlConnection connection = new SqlConnection(connectionString))  //connect to the sql server
                        using (SqlCommand cmd = connection.CreateCommand())  //start a  sql command
                        using (SqlDataAdapter sda = new SqlDataAdapter(cmd))  //use the adapter to get the command results back
                        {
                            try
                            {
                                cmd.CommandText = sqlCmd;  //set the commandtext to the sqlcmd
                                cmd.CommandType = CommandType.Text;  //set it as a text command
                                connection.Open();  //open the sql server connection to the database
                                rowsreturned = sda.Fill(dt);  //run the command and put the results into the datatable and store the row count
                                connection.Close();  //close the sql server connection to the database

                                //declare defaults for the report not being in the database
                                String RPTappl = RPTnumber.Substring(0, 3);
                                String RPTtitle = "Unknown-" + RPTnumber;
                                String RPTretention = "9999";
                                Unknown = "YES";

                                if (dt.Rows.Count == 0)
                                {
                                    //no query results, use defaults
                                }
                                else
                                {
                                    foreach (DataRow drstr in dt.Rows)  //assign each row returned into the variable drstr
                                    {
                                        string[] values = new string[dt.Columns.Count];  //create a string array with the same number of columns as the datatable (3 in this case)

                                        for (int i = 0; i < dt.Columns.Count; ++i)    //iterate through each column in the datatable for this row
                                        {
                                            values[i] = ((string)drstr[i]) ?? "";  //set the value(i) to the column(i) in the datatable but if it is nul ?? then set it to ""
                                        }
                                        //use the values from the table as additional keywords in the index file
                                        RPTappl = values[0];
                                        RPTtitle = values[1];
                                        RPTretention = values[2];
                                        Unknown = "NO";
                                    }
                                }
                                //Calculate Retention End Date keyword for use in workflow to delete when exceeded
                                String RED = "";
                                if (RPTretention == "0")
                                {
                                    //0 so no retention date;permanent record
                                }
                                else
                                {
                                    DateTime Docdate2 = Convert.ToDateTime(Docdate);
                                    DateTime RetentionEndDate = Docdate2.AddDays(Convert.ToInt32(RPTretention));
                                    RED = RetentionEndDate.ToString("MM/dd/yyyy");
                                }

                                String Desc = RPTnumber + " - " + RPTtitle;  //Description keyword part of autoname string

                                String outDIPindexfile = "DIPindex_" + pathdate + "_" + filename + ".txt".Replace(" ", "");  //the name of the index file to be used for this report txt file
                                String outCOPYfile = pathdate + "_" + filenamewithextension.Replace(" ", "");  //the name of the file prefixed with the folder date it is in to handle multiple report files for same day

                                String DIPIndexValue = "ASB Report Legacy\t" + Docdate + "\t" + RPTinstance + "\t" + RPTnumber + "\t" + RPTappl + "\t" + RPTtitle + "\t" + RED + "\t" + Desc + "\t" + useDrivePath + slash + outCOPYfile;

                                //update the report control table if the report number was unknown so the next time we run into it will have values
                                if (Unknown == "YES")
                                {
                                    sqlCmd = "Insert Into " + useTable + " ([Application],[ReportNumber],[ReportTitle],[RPTretentionDays]) values ('" + RPTappl + "', '" + RPTnumber + "', '" + RPTtitle + "', '" + RPTretention + "')";
                                    using (SqlConnection connection2 = new SqlConnection(connectionString))  //connect to the sql server
                                    using (SqlCommand cmd2 = connection2.CreateCommand())  //start a  sql command
                                        try
                                        {
                                            cmd2.CommandText = sqlCmd;  //set the commandtext to the sqlcmd
                                            cmd2.CommandType = CommandType.Text;  //set it as a text command
                                            connection2.Open();  //open the sql server connection to the database
                                            int rowsadded = cmd2.ExecuteNonQuery();  //run the command and store the row count inserted
                                            connection2.Close();  //close the sql server connection to the database
                                        }
                                        catch (Exception ex)
                                        {
                                            Console.WriteLine("Error Inserting Unknown Report: " + ex);
                                        }
                                }
                                //now that we have created the index values for this report, write the file out, copy the report from the temp folder to the DIP folder
                                File.WriteAllText(useOutPath + slash + outDIPindexfile, DIPIndexValue);
                                File.Copy(fullpathfilename, useOutPath + slash + outCOPYfile, true);
                                File.Delete(fullpathfilename);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("Error reading Report Name metadata from table: " + ex);
                            }
                        }
                    }
                    String sqlCmd3 = "UPDATE " + useStatsTable + " set [DIPStatus]='REPORTS COMPLETE' where [RptDate]='" + pathdate + "'";
                    using (SqlConnection connection3 = new SqlConnection(connectionString))  //connect to the sql server
                    using (SqlCommand cmd3 = connection3.CreateCommand())  //start a  sql command
                    try
                    {
                        cmd3.CommandText = sqlCmd3;  //set the commandtext to the sqlcmd
                        cmd3.CommandType = CommandType.Text;  //set it as a text command
                        connection3.Open();  //open the sql server connection to the database
                        int rowsadded = cmd3.ExecuteNonQuery();  //run the command and store the row count inserted
                        connection3.Close();  //close the sql server connection to the database
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error updating Stats table: " + ex);
                    }

                    foreach (string notitem in notitems)
                    {
                        Int32 notpathlength = notitem.Length;
                        Int32 filenamelength = notitem.Length - notitem.LastIndexOf(slash) - 1;  //length of the filename+ext minus the fullpath and \Notes to it
                        String filenamextension = Path.GetExtension(notitem);
                        Int32 filenamextensionlength = filenamextension.Length;
                        String filename = notitem.Substring(notitem.LastIndexOf(slash) + 1, filenamelength - filenamextension.Length);
                        String filenamewithextension = filename + filenamextension;
                        String fullpathfilename = notitem;

                        String RPTnumber = filenamewithextension.Substring(0, filenamewithextension.IndexOf("."));
                        ++notcount;
                        Console.WriteLine("Working Date Folder: " + pathdate + "# " + datecount + " of " + numdateitems + " | Report: " + txtcount + " of " + numtxtitems + " | Note: " + notcount + " of " + numnotitems);

                        RPTnumber = filenamewithextension.Substring(0, filenamewithextension.IndexOf("."));
                        String RPTinstance = filenamewithextension.Substring(0, filenamelength - 10);
                        String RPTnote = File.ReadAllText(fullpathfilename);
                        String sqlCmd4 = "Insert Into " + useNoteTable + " ([ReportNumber],[ReportNumberInstance],[RptDate],[RptNote]) values ('" + RPTnumber + "', '" + RPTinstance + "', '" + Docdate + "', '" + RPTnote + "')";
                        using (SqlConnection connection4 = new SqlConnection(connectionString))  //connect to the sql server
                        using (SqlCommand cmd4 = connection4.CreateCommand())  //start a  sql command
                        try
                        {
                            cmd4.CommandText = sqlCmd4;  //set the commandtext to the sqlcmd
                            cmd4.CommandType = CommandType.Text;  //set it as a text command
                            connection4.Open();  //open the sql server connection to the database
                            int rowsadded = cmd4.ExecuteNonQuery();  //run the command and store the row count inserted
                            connection4.Close();  //close the sql server connection to the database
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Error inserting into Notes table: " + ex);
                        }
                        String sqlCmd5 = "UPDATE " + useStatsTable + " set [DIPStatus]='ALL COMPLETE' where [RptDate]='" + pathdate + "'";
                        using (SqlConnection connection5 = new SqlConnection(connectionString))  //connect to the sql server
                        using (SqlCommand cmd5 = connection5.CreateCommand())  //start a  sql command
                        try
                        {
                            cmd5.CommandText = sqlCmd5;  //set the commandtext to the sqlcmd
                            cmd5.CommandType = CommandType.Text;  //set it as a text command
                            connection5.Open();  //open the sql server connection to the database
                            int rowsadded = cmd5.ExecuteNonQuery();  //run the command and store the row count inserted
                            connection5.Close();  //close the sql server connection to the database
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Error updating Stats table: " + ex);
                        }
                        File.Delete(notitem);
                    }

                    //remove date folder Notes folder
                    if (Directory.Exists(dateitem + slash + "Notes"))
                    {
                        Directory.Delete(dateitem + slash + "Notes");
                    }
                    //remove date folder
                    Directory.Delete(dateitem);
                }
            }
            catch
            {
                Console.WriteLine("App.config does not exist or does not meet requirements.");
                Environment.Exit(0);
            }
        }
    }
}
