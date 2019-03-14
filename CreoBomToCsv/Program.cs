using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using pfcls;
using System.IO;

namespace CreoBomToCsv
{
    class Program
    {

        
        static IpfcAsyncConnection asyncConnection;
        static string lastModelDirectory = "";
        static string lastModelName = "";
        static string outputDirectory = "";
        

        static void Main(string[] args)
        {
            CheckEnvironment();

            if (args.Length > 0)
            {
                string arg0 = args[0];
                if (arg0 == "/?")
                {
                    WriteUsage();
                    Environment.Exit(0);
                }
                else {
                    arg0 = arg0.Replace("\"", "").TrimEnd(Path.DirectorySeparatorChar);
                    if (System.IO.Directory.Exists(arg0))
                    {
                        //set the output directory to arg0
                        outputDirectory = arg0;
                        Console.WriteLine("Output Directory: {0}", outputDirectory);
                    }
                }
                
            }


            // Get the current Creo session
            ConnectToSession();


            GetTableData();
            



            DisconnectSession();

        }

        private static void GetTableData()
        {



            IpfcBaseSession CreoSession;

            CreoSession = (IpfcBaseSession)asyncConnection.Session;



            IpfcModel CreoModel;
            CreoModel = CreoSession.CurrentModel;

            if (CreoModel == null)
            {
                Console.WriteLine("Could not get active model!");
                return;
            }

            lastModelName = CreoModel.FullName;

            Console.WriteLine("Current Model:  {0}", lastModelName);

            FileInfo fi = new FileInfo(CreoModel.Origin);
            lastModelDirectory = fi.Directory.FullName;

            Console.WriteLine("Model Directory:  {0}", lastModelDirectory);


            pfcls.IpfcTableOwner TableObj = (pfcls.IpfcTableOwner)CreoModel;

            CpfcTables table_list = TableObj.ListTables();

            if (table_list.Count == 0)
            {
                Console.WriteLine("Could not find any tables!");
                return;
            }

            // [===========================================]
            Console.WriteLine("Gathering data...");


            bool boolFoundTable = false;
            string strOutput = "";

            for (int i = 0; i < table_list.Count; i++)
            {
                if (table_list[i].GetRowCount() > 1)
                {

                    // if the table was already found... exit outer loop
                    if (boolFoundTable) { break; }

                    IpfcTable table = table_list[i];


                    int RowCount = table.GetRowCount();
                    int ColumnCount = table.GetColumnCount();



                    for (int j = 1; j <= RowCount; j++)
                    {

                        //Has multiple rows.. check to see if Row 0 matches our test string
                        string rowData = "";
                        bool validRow = false;

                        for (int k = 1; k <= ColumnCount; k++)
                        {

                            CCpfcTableCell tableCellCreate = new CCpfcTableCell();
                            IpfcTableCell tableCell = tableCellCreate.Create(j, k);

                            string ItemValue = "";
                            try
                            {

                                Cstringseq stringSeq = table.GetText(tableCell, 0);

                                //Console.WriteLine("stringSeq.Count = {0}", stringSeq.Count);
                                for (int x = 0; x < stringSeq.Count; x++)
                                {
                                    ItemValue += stringSeq[x].ToString();
                                    //Console.WriteLine(stringSeq(x).ToString());
                                }
                            }
                            catch (Exception ex)
                            {
                                //Console.WriteLine(ex.Message);
                            }

                            //If this is the FIRST column of the FIRST row.. 
                            //see if the text matches our search string
                            //DON'T include it in the output though...
                            if (j == 1 && k == 1)
                            {
                                if (ItemValue.Contains("AUTO BOM"))
                                {
                                    boolFoundTable = true;
                                }
                                else
                                {
                                    break;
                                }
                            }
                            else
                            {
                                if (!validRow)
                                {
                                    if (!String.IsNullOrEmpty(ItemValue.Trim()))
                                    {
                                        validRow = true;
                                    }
                                }

                                if (validRow)
                                {
                                    //Console.Write(ItemValue + ",");
                                    rowData += "\"" + ItemValue + "\"";
                                    if (k < ColumnCount)
                                    {
                                        rowData += ",";
                                    }
                                }

                            }


                        }
                        if (validRow)
                        {
                            strOutput += rowData + "\r\n";
                            Console.WriteLine(rowData);
                        }


                        if (!boolFoundTable)
                        {
                            break;
                        }

                        //progress update?
                    }


                }



            }

            if (!boolFoundTable)
            {
                Console.WriteLine("Could not find AUTO BOM table!");
            }


            //Save file
            string strFilename;
            if (!string.IsNullOrEmpty(outputDirectory) && System.IO.Directory.Exists(outputDirectory))
                {
                strFilename = outputDirectory + "\\" + lastModelName + ".csv";
            } else
            {
                strFilename = lastModelDirectory + "\\" + lastModelName + ".csv";
            }
                //= lastModelDirectory + "\\" + lastModelName + ".csv";
            File.WriteAllText(strFilename, strOutput);
        
            Console.WriteLine("Done!");
            Console.WriteLine("Saved to: {0}",strFilename);


        }

        private static void DisconnectSession()
        {
             if (asyncConnection != null && asyncConnection.IsRunning()) {
                try 
                {
                        asyncConnection.Disconnect(1);
                } 
                catch (Exception ex) 
                {
                    //Do nothing
                }
            }
            
        }

        private static void ConnectToSession()
        {

            Console.WriteLine("Connecting to Creo session...");

            try
            {
                asyncConnection = (IpfcAsyncConnection)new CCpfcAsyncConnection().Connect(null, null, null, 5);
                Console.WriteLine("Connected!");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error connecting to Creo Parametric Session!");
                Console.WriteLine(ex.Message);
                WriteTips();
                Environment.Exit(1);
            }
            

        }

        private static void WriteTips()
        {
            Console.WriteLine();
            Console.WriteLine("Tips:");
            Console.WriteLine("1) Make sure Creo is running");
            Console.WriteLine("2) Register the API using the \"vb_api_register.bat\" From the Creo directory(run elevated!)");
            Console.WriteLine("3) Make sure the environment variables are set correctly:");
            Console.WriteLine("     PRO_COMM_MSG_EXE");
            Console.WriteLine("     PRO_DIRECTORY");
            Console.WriteLine("4) If the process \"pfclscom.exe\" is running, terminate it");

            Console.WriteLine();

        }

        private static void WriteUsage()
        {
            
            Assembly currentAssem = Assembly.GetExecutingAssembly();
            AssemblyName name = currentAssem.GetName();
            //string version = string.Format("{0}.{1:2}.{2:2}.{3:2}",
            
            Console.WriteLine();
            Console.WriteLine("Name:         {0}",name.Name);
            Console.WriteLine("Version:      {0}", name.Version);
            Console.WriteLine("Author:       Joe Ostrander");
            Console.WriteLine("Build date:   {0}",DateTime.Now);
            Console.WriteLine("Description:  Extract the BOM table from a Creo drawing");
            Console.WriteLine();
            Console.WriteLine("Optional usage:  CreoBomToCsv.exe <output directory>");
            Console.WriteLine();


        }

        private static void CheckEnvironment()
        {

            string strEnv1 = Environment.GetEnvironmentVariable("PRO_COMM_MSG_EXE");
            string strEnv2 = Environment.GetEnvironmentVariable("PRO_DIRECTORY");


            if (string.IsNullOrEmpty(strEnv1) || String.IsNullOrEmpty(strEnv2))
            {
                Console.WriteLine("Warning:  Set the environment variables!");
                Console.WriteLine();
                Console.WriteLine("Example values:");
                Console.WriteLine("PRO_COMM_MSG_EXE=\"c:\\Program Files\\PTC\\Creo 4.0\\M030\\Common Files\\x86e_win64\\obj\\pro_comm_msg.exe\"");
                Console.WriteLine("PRO_DIRECTORY=\"C:\\Program Files\\PTC\\Creo 4.0\\M030\\Common Files\"");
                System.Threading.Thread.Sleep(10000);
                Environment.Exit(1);
            }
        }
    }
}
