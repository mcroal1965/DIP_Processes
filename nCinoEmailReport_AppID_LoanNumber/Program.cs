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
                    String temptxtfile2 = useInPath + slash + itemname + ".trtxt";
                    File.Delete(temptxtfile2);
                    String temptxtfile3 = useInPath + slash + itemname + ".tdtxt";
                    File.Delete(temptxtfile3);
                    String temptxtfile4 = useInPath + slash + itemname + ".col13txt";
                    File.Delete(temptxtfile4);
                    //setup the filename for the one that will contain the COLD records, delete it if it already exists
                    String useOUTfile = useOutPath + slash + "GIM_Repair_nCinoEmail_" + itemname + ".txt";
                    File.Delete(useOUTfile);

                    #region CreateGoodHTML
                    Console.WriteLine("Processing file: " + item + " step 1 of 4");
                    string[] badhtmllines = File.ReadAllLines(item);
                    foreach (string badhtmlline in badhtmllines)
                    {
                        if (badhtmlline == "<html><head>")
                        {
                            foundhtml = 1;
                        }
                        if (foundhtml == 1 && badhtmlline != "<html><head>" && badhtmlline != "</html>")
                        {
                            lineout = badhtmlline.Replace("<meta http-equiv=3D\"Content-Type\" content=3D\"text/html; charset=3Dutf-8\">", "");
                            lineout = lineout.Replace("=3D", "=").Replace("<br>", "").Replace("&nbsp;", "");
                            File.AppendAllText(temptxtfile, lineout.Substring(0, lineout.Length - 1));
                        }
                        else
                        {
                            if (foundhtml == 1)
                            {
                                File.AppendAllText(temptxtfile, badhtmlline);
                            }
                        }
                    }
                    #endregion
                    #region ReadGoodHTMLWriteTableAsHTMLlines
                    Console.WriteLine("Processing file: " + item + " step 2 of 4");
                    XmlDocument xDoc = new XmlDocument();
                    xDoc.Load(temptxtfile);
                    foreach (XmlNode node in xDoc.DocumentElement.ChildNodes)
                    {
                        foreach (XmlNode allNodes in node)
                        {
                            string x = allNodes.InnerText;

                            if (x.Length >= 16)
                            {
                                if (x.Substring(0, 16) == "Loan NumberLoan:")
                                {
                                    string y = allNodes.InnerXml;
                                    Int32 x1 = y.IndexOf("<tr");
                                    Int32 x2 = y.IndexOf("</tr");
                                    string y1 = x.Substring(x1, x2 - x1);
                                    lineout = y.Replace("<tr", "\r<tr").Replace("</tr", "\r</tr");
                                    File.AppendAllText(temptxtfile2, lineout);
                                }
                            }
                        }
                    }
                    #endregion

                    #region WriteLinesAsRowColumns
                    Console.WriteLine("Processing file: " + item + " step 3 of 4");
                    string[] trlinein = File.ReadAllLines(temptxtfile2);
                    foreach (string trline in trlinein)
                    {
                        if (trline.Length > 20)
                        {
                            if (trline.Substring(0, 20) == "<tr class=\"dataRow\">")
                            {
                                string z = trline.Replace("<td", "\r<td").Replace("</td", "\r</td").Replace("<tr", "\r<tr");
                                File.AppendAllText(temptxtfile3, z);
                            }
                        }

                    }
                    #endregion

                    #region ReadLinesAsTableColumns
                    Console.WriteLine("Processing file: " + item + " step 4 of 4");
                    string[] tdlinein = File.ReadAllLines(temptxtfile3);
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
                        }
                        if (colcount == 3)
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
                    #endregion

                    #region CleanUp
                    //delete the .msg item 
                    File.Delete(item);
                    //delete the .txt file that stored the message body has html
                    File.Delete(temptxtfile);
                    //delete the .txt file that stored the html rows
                    File.Delete(temptxtfile2);
                    //delete the .txt file that stored the html columns
                    File.Delete(temptxtfile3);
                    #endregion
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception encountered .eml region:" + ex);
                Console.WriteLine("Press any key to exit.");
                Console.ReadKey();
                Environment.Exit(0);
            }
            #endregion
        }
    }
}
