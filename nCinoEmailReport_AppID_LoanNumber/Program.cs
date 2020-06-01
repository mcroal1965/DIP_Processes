using System.Xml;
using System;
using System.IO;
using System.Configuration;

namespace nCinoEmailReport_AppID_LoanNumber
{
    class Program
    {
        public static String LOANnumber = "0";
        public static String lineout;
        public static String lineout2;
        public static String loannumber;
        public static String appid;
        public static Int32 colcount;
        public static Int32 foundhtml;

        public static void Main()
        {
            try
            {
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

                String slash = Convert.ToString(Convert.ToChar(92));  //store the slash so it can be used in the filename later

                #region Handle msg formatted file
                try
                {
                    //Get all outlook message items in the folder, these mail messages are exported to the useInPath folder from the workflow action after the mailbox importer imports the message sent by nCino
                    String[] allitems = Directory.GetFiles(useInPath, "*.msg");

                    foreach (String item in allitems)
                    {
                        FileInfo f = new FileInfo(item);
                        String itemname = f.Name;  //the name of the message file which is the dochandle that the workflow export for network folder gave it

                        //write the text portion of the body out to a temporary file replacing tabs with new lines
                        String temptxtfile = useInPath + slash + itemname + ".txt";
                        File.Delete(temptxtfile);
                        //setup the filename for the one that will contain the COLD records, delete it if it already exists
                        String useOUTfile = useOutPath + slash + "GIM_Repair_nCinoEmail_" + itemname + ".txt";
                        File.Delete(useOUTfile);

                        try
                        {
                            using (var msg = new MsgReader.Outlook.Storage.Message(item))
                            {
                                var textBody = msg.BodyText.Replace("\t", "\n").Replace("_R1", "");
                                File.WriteAllText(temptxtfile, textBody);
                            }
                        }
                        catch
                        {
                            Console.WriteLine("File is not an Outlook formated message.");
                            Console.WriteLine("Press any key to exit.");
                            Console.ReadKey();
                            Environment.Exit(0);
                        }

                        //read in the file that contains the text version of the message body
                        foreach (String lineintxt in File.ReadLines(temptxtfile))
                        {
                            try
                            {
                                //identify any large integers as loan numbers
                                Int64 lineinint = Convert.ToInt64(lineintxt);
                                if (lineinint > 9999)
                                {
                                    LOANnumber = Convert.ToString(lineinint);
                                }
                            }
                            catch
                            {
                                //identify nCino AppIDs as anything that starts with AN-
                                if (lineintxt.StartsWith("AN-"))
                                {
                                    String ANnumber = lineintxt;
                                    if (Convert.ToInt64(LOANnumber) > 0)
                                    {
                                        String GIMDocType = "WF GIM Repair           ".Substring(0, 20);  //COLD config is doctype in 1st 20 characters
                                        String GIMMaintType = "nCinoLoan#Swap                       ".Substring(0, 25);  //COLD config is desc kw in 21 for 25 characters
                                        String lineout = GIMDocType + GIMMaintType + "||" + ANnumber + "|" + LOANnumber + "\n";
                                        File.AppendAllText(useOUTfile, lineout);
                                    }
                                    //some AN- may not be preceeded by a loan number so this ensures only records that have a Loan# and an AN- number are handled
                                    LOANnumber = "0";
                                }
                            }
                        }
                        //delete the .msg item 
                        File.Delete(item);
                        //delete the .txt file that stored the message body text
                        File.Delete(temptxtfile);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Exception encountered .msg region:" + ex);
                    Console.WriteLine("Press any key to exit.");
                    Console.ReadKey();
                    Environment.Exit(0);
                }
                #endregion

                #region Handle eml formatted file
                try
                {
                    //Get all eml message items in the folder, these mail messages are exported to the useInPath folder from the workflow action after the mailbox importer imports the message sent by nCino
                    String[] allitemseml = Directory.GetFiles(useInPath, "*.eml");

                    foreach (String item in allitemseml)
                    {
                        foundhtml = 0;
                        FileInfo f = new FileInfo(item);
                        String itemname = f.Name;  //the name of the message file which is the dochandle that the workflow export for network folder gave it

                        //write the text portion of the body out to a temporary file replacing tabs with new lines
                        String temptxtfile = useInPath + slash + itemname + ".html";
                        File.Delete(temptxtfile);
                        String temptxtfile2 = useInPath + slash + itemname + ".newhtmltxt";
                        File.Delete(temptxtfile2);
                        //setup the filename for the one that will contain the COLD records, delete it if it already exists
                        String useOUTfile = useOutPath + slash + "GIM_Repair_nCinoEmail_" + itemname + ".txt";
                        File.Delete(useOUTfile);

                        #region CreateGoodHTML
                        Console.WriteLine("Processing file: " + item + " step 1 of 3");
                        string[] badhtmllines = File.ReadAllLines(item);
                        foreach (string badhtmlline in badhtmllines)
                        {
                            bool foundhtmlstart = badhtmlline.Contains("<html");
                            bool foundbodystart = badhtmlline.Contains("<body");
                            
                            if (foundhtmlstart == true)
                            {
                                lineout = "<html>";
                                File.AppendAllText(temptxtfile, lineout);
                            }
                            else
                            {
                                if (foundbodystart == true)
                                {
                                    foundhtml = 1;
                                }
                                if (foundhtml == 1)
                                {
                                    lineout = badhtmlline;
                                    lineout = lineout.Replace("<br>", "").Replace("&nbsp;", "").Replace("_R1", "");
                                    if (lineout.Length > 0 && lineout.Substring(lineout.Length - 1, 1) == "=")
                                    {
                                        lineout = lineout.Substring(0, lineout.Length - 1);
                                    }
                                    File.AppendAllText(temptxtfile, lineout.Replace("=3D", "=").Replace("<o:p>", "").Replace("</o:p>", ""));
                                }
                            }
                        }
                        #endregion
                        #region CleanUpHTMLerrors
                        //read temptxtfile as one line, do replaces as in lines 162 & 167 and any others that crop up and write out as tmptxtfile4
                        Console.WriteLine("Processing file: " + item + " step 2 of 3");
                        try
                        {
                            string[] goodhtmllines = File.ReadAllLines(temptxtfile);
                            foreach (string goodhtmlline in goodhtmllines)
                            {
                                string goodhtmllineout = goodhtmlline.Replace("<", "\r<");//.Replace("=3D", "=").Replace("<o:p>", "").Replace("</o:p>", "").Replace("<br>", "").Replace("&nbsp;", "").Replace("_R1", "").Replace("_M1", "");
                                goodhtmllineout = goodhtmllineout.Replace("\r<html>","<html>");
                                File.AppendAllText(temptxtfile2, goodhtmllineout);
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Exception encountered CleanUpHTMLerrors region:" + ex);
                            Console.WriteLine("Press any key to exit.");
                            Console.ReadKey();
                            Environment.Exit(1);
                        }
                        #endregion

                        #region ReadLinesAsTableColumns
                        Console.WriteLine("Processing file: " + item + " step 3 of 3");
                        string[] tdlinein = File.ReadAllLines(temptxtfile2);
                        foreach (string tdline in tdlinein)
                        {
                            if (tdline.Length > 10)
                            {
                                if (tdline.Substring(0, 9) == "<tr class")
                                {
                                    colcount = 0;
                                    lineout2 = "";
                                    loannumber = "";
                                    appid = "";
                                }
                            }
                            if (tdline.Length > 3)
                            {
                                if (tdline.Substring(0, 3) == "<td")
                                {
                                    colcount++;
                                }
                            }
                            if (colcount == 1 && tdline != "</td>")
                            {
                                loannumber = tdline.Substring(tdline.IndexOf(">") + 1, tdline.Length - tdline.IndexOf(">") - 1);
                                if (loannumber.IndexOf("_") > 0)
                                {
                                    loannumber = loannumber.Substring(0, loannumber.IndexOf("_") - 1);
                                }
                            }
                            if (colcount == 3)
                            {
                                if (loannumber.Length >= 2)
                                {
                                    appid = tdline.Substring(tdline.IndexOf(">") + 1, tdline.Length - tdline.IndexOf(">") - 1);
                                    if (loannumber != "-" && appid != "" && loannumber.Substring(0, 2) != "AN")
                                    {
                                        String GIMDocType = "WF GIM Repair           ".Substring(0, 20);  //COLD config is doctype in 1st 20 characters
                                        String GIMMaintType = "nCinoLoan#Swap                       ".Substring(0, 25);  //COLD config is desc kw in 21 for 25 characters
                                        lineout2 = GIMDocType + GIMMaintType + "||" + appid + "|" + loannumber + "\n";
                                        File.AppendAllText(useOUTfile, lineout2);
                                    }
                                }
                            }
                        }
                        Console.WriteLine("**********Finished processing file: " + item + "***********");
                        #endregion

                        #region CleanUp
                        //delete the .msg item 
                        File.Delete(item);
                        //delete the .txt file that stored the message body has html
                        File.Delete(temptxtfile);
                        //delete the .txt file that stored the good html rows
                        File.Delete(temptxtfile2);
                        
                        #endregion
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Exception encountered .eml region:" + ex);
                    Console.WriteLine("Press any key to exit.");
                    Console.ReadKey();
                    Environment.Exit(1);
                }
                #endregion
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
