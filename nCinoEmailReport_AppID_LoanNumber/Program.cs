using System;
using System.IO;
using System.Configuration;

namespace nCinoEmailReport_AppID_LoanNumber
{
    class Program
    {
        public static String LOANnumber = "0";

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
    
            //Get all outlook message items in the folder, these mail messages are exported to the useInPath folder from the workflow action after the mailbox importer imports the message sent by nCino
            String[] allitems = Directory.GetFiles(useInPath, "*.msg");
            
            foreach (String item in allitems)
            {
                FileInfo f = new FileInfo(item);
                String itemname = f.Name;  //the name of the message file which is the dochandle that the workflow export for network folder gave it
                
                //write the text portion of the body out to a temporary file replacing tabs with new lines
                String temptxtfile = useInPath + slash + itemname + ".txt";
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
                                String lineout = GIMDocType + GIMMaintType + "||" + ANnumber + "|" + LOANnumber+"\n";
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
    }
}
