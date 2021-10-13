/* FixtureLabelReprint_ByPart - Reprint Fixture Labels for Parts by Part Number                                            */
/* This project prints Fixture Labels for a specified Part. This is a .NET conversion of the original E9 Progress program. */
/* Modifications                                                                                                           */
/* Date        By      Description                                                                                         */
/* 04/12/2019 jmyers   Changed code to function in the same manner as Reprint By Job, Print Fixture Label projects         */
/* 04/23/2019 jmyers   Added code to find Pick Codes for processing alternate messages, adjusted Watts/Hertz/IP Class      */
/*                     field positioning on large labels                                                                   */
/* 06/12/2019 jmyers   Commented out code that includes the phrase "in USA" per IT Request# 19216                          */
/* 10/09/2019 jmyers   Changed code to print only number of lamp codes indicated in UD100.Number02 field                   */
/* 10/21/2019 jmyers   Changed code to eliminate substring retrieval if the string is not as long as expected              */
/* 01/23/2020 jmyers   Changed Watts and IP Class positioning on Large labels to avoid overlap with Hertz label and value, */
/*                     per IT request# 20294                                                                               */
/* 10/02/2020 jmyers   Changed code in GetAddressInfo method to use the AddressCode variable value if it's not blank to    */
/*                     find AddrMst type label data entries instead of always using "Kenall"                               */
/* 01/07/2021 jmyers   Changed code in GetAddressInfo method to find and flag a part as an Indigo Clean part by data in the*/
/*                     Series parent or child entries, print Indigo Clean labels in a different format from other labels   */
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace FixtureLabelReprint_ByPart_WebApp
{
    public partial class _Default : Page
    {
        // Global Variables
        //public string studoCode;
        public bool hasF0862Material = false;
        public bool IndigoCleanPart = false;
        public bool isDownlight = false;
        public bool isSTUDO = false;
        public bool isSTUDX = false;
        public bool isSTUDZ = false;
        public bool multipleLabelsNeeded;
        public bool needsULLabel = false;
        public bool prtmsg12 = false;

        public int ampcode1 = 0;
        public int ampcode2 = 0;
        public int ampcode3 = 0;
        public int ampcode4 = 0;
        public int element = 0;
        public int msgCount;
        public int nbrLamps;
        public int partCount;
        public int printCounter;
        public int prodQty;
        public int ucbElementCounter = 2;
        public int userAmount = 0;
        public int wrkNbrLamps;

        public string addressCode;
        public string AES_PartNum = "";
        public string asemCode = "";
        public string currentID;
        public string F0826_DueDate;
        public string FixtureLabelFound;
        public string ipClass = "";
        public string jobDueDate;
        public string jobPartDesc;
        public string jobPartNum;
        public string jobValue;
        public string labelSize;
        public string labelTypeCheck;
        public string lampCode1 = "";
        public string lampCode2 = "";
        public string lampCode3 = "";
        public string materialFound;
        public string msgID;
        public string mtlPartNum;
        public string oldPart;
        public string part1 = "";
        public string part2 = "";
        public string part3 = "";
        public string part4 = "";
        public string part5 = "";
        public string part6 = "";
        public string part7 = "";
        public string part8 = "";
        public string part9 = "";
        public string part10 = "";
        public string part11 = "";
        public string part12 = "";
        public string part13 = "";
        public string part14 = "";
        public string part15 = "";
        public string part16 = "";
        public string part17 = "";
        public string part18 = "";
        public string part19 = "";
        public string part20 = "";
        public string partDesc = "";
        public string partNum = "";
        public string pickCode;
        public string plannerName = "";
        public string prtLamp;
        public string prtLamp1;
        public string prtLamp2;
        public string prtLamp3;
        public static string sqlConnString = "";
        public string SSCS_PartNum = "";
        public string studioText;
        public string ucbFound = "N";
        public string ud100Key = "";
        public string udTableFound = "";
        public string validRecords;
        public string voltkey1 = "";
        public string voltkey2 = "";
        public string voltkey3 = "";
        public string voltkey4 = "";
        public string voltsFound = "NO";
        //public string wrkADDR;
        public static string wrkADDR1;
        public static string wrkADDR2;
        public static string wrkADDR3;
        public string wrkhza = "";
        public string wrkhzb = " ";
        public string wrklamptype = " ";
        public string wrkvolts = "";
        public string wrkwatts = " ";

        public static string _ServerID;
        public static string DBConnection = "";
        public static string eServer;



        public List<string> elementList = new List<string>();
        public List<string> jobList = new List<string>();
        public List<string> MessagesList = new List<string>();

        /********** CHECK THESE VARIABLES BEFORE PUBLISHING TO PRODUCTION **********/
        public string LargeLabelPrint = @"\\KPrint01\LBL910";
        public string SmallLabelPrint = @"\\KPrint01\LBL909";
        //public string LargeLabelPrint = @"\\KPrint01\LBL900";   /* Test Printer */
        //public string SmallLabelPrint = @"\\KPrint01\LBL900";   /* Test Printer */

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                btnPrintLabels.Attributes.Add("onclick", "this.disabled=true;this.value='Please Wait...';" + ClientScript.GetPostBackEventReference(btnPrintLabels, string.Empty));
            }

            btnPrintLabels.Attributes.Add("onclick", "this.disabled=true;this.value='Please Wait...';" + ClientScript.GetPostBackEventReference(btnPrintLabels, string.Empty));
        }

        protected void btnPrintLabels_Click(object sender, EventArgs e)
        {
            // Variables
            lblError.Visible = false;
            jobValue = "";
            printCounter = 0;
            prodQty = 0;
            multipleLabelsNeeded = false;
            hasF0862Material = false;
            F0826_DueDate = "";
            prtmsg12 = false;

            _ServerID = "EpicorErpTest";
            SelectDatabase(_ServerID);


            int counter = 0;

            // Getting Op Code for Job, If this errors out the job is invalid for fixture label printing
            validRecords = "";
            //jobValue = tbPartNum.Text;
            //SqlCommand checkOpCMD = new SqlCommand();
            //checkOpCMD.Connection = new SqlConnection(sqlConnString);
            //String checkOp = "SELECT TOP 1 OpCode FROM EpicorERP.dbo.JobOper WHERE JobNum = '" + jobValue + "'";

            //// Asem Code
            //checkOpCMD.CommandText = checkOp;
            //checkOpCMD.Connection.Open();
            //asemCode = checkOpCMD.ExecuteScalar().ToString();
            //checkOpCMD.Connection.Close();

            jobPartNum = tbPartNum.Text;    /* JWM 04/08/2019 Added to feed Part Number to next routine */
            // Run Function that checks for F-0862 material
            CheckIfF0862();

            // If Job does not have an F-0862 Materal then process normally, Else Print special label for that job
            if (hasF0862Material == false)
            {
                // Create connections - State Database Connection, Main SQL SELECT, ExecuteScalars
                SqlCommand commPartDesc = new SqlCommand();
                // SqlCommand commDueDate = new SqlCommand();
                // SqlCommand commPartNum = new SqlCommand();
                // SqlCommand commName = new SqlCommand();
                // SqlCommand commProdQty = new SqlCommand();

                commPartDesc.Connection = new SqlConnection(sqlConnString);
                // commPartNum.Connection = new SqlConnection(sqlConnString);
                // commName.Connection = new SqlConnection(sqlConnString);
                // commDueDate.Connection = new SqlConnection(sqlConnString);
                // commProdQty.Connection = new SqlConnection(sqlConnString);

                try
                {
                    // MAIN SQL STATEMENT - SQL JOINS etc..
                    String sqlPartDesc = "Select PartDescription FROM " + eServer + ".dbo.Part WHERE Company = 'KEN' AND PartNum = '" + tbPartNum.Text + "' AND ProdCode NOT LIKE 'RP%'";
                    // String sqlPartNum = "Select PartNum FROM dbo.Part WHERE JobNum = '" + jobValue + "'";
                    // String sqlDueDate = "Select DueDate FROM dbo.JobHead WHERE JobNum = '" + jobValue + "'";
                    // String sqlName = "Select T2.Name FROM dbo.JobHead as T1 INNER JOIN Erp.Person as T2 ON T2.Company = T1.Company AND T2.PersonID = T1.PersonID  WHERE JobNum = '" + jobValue + "'";
                    // String sqlProdQty = "Select ProdQty FROM dbo.JobHead WHERE JobNum = '" + jobValue + "'";

                    //// Part Desc
                    commPartDesc.CommandText = sqlPartDesc;
                    commPartDesc.Connection.Open();
                    partDesc = commPartDesc.ExecuteScalar().ToString();
                    commPartDesc.Connection.Close();

                    //// Prod Qty
                    //commProdQty.CommandText = sqlProdQty;
                    //commProdQty.Connection.Open();
                    //prodQty = Convert.ToInt32(commProdQty.ExecuteScalar());
                    //commProdQty.Connection.Close();

                    // Part Num

                    partNum = tbPartNum.Text;

                    //// Name
                    //commName.CommandText = sqlName;
                    //commName.Connection.Open();
                    //plannerName = commName.ExecuteScalar().ToString();
                    //commName.Connection.Close();

                    //// Due Date
                    //commDueDate.CommandText = sqlDueDate;
                    //commDueDate.Connection.Open();
                    //jobDueDate = commDueDate.ExecuteScalar().ToString();
                    //commDueDate.Connection.Close();
                    
                }
                catch
                {
                    lblError.Visible = true;
                    lblError.Text = "Not a valid Fixture Label Part #";
                    return;
                  //  Environment.Exit(0);
                }
                // Setting Global Variables for this Label Job //
                jobPartNum = partNum;
                jobPartDesc = partDesc;

                // Split part description at the dash (-) sign
                string[] parts = partDesc.Split('-');
                int arrayLength = parts.Length;
                partCount = 0;

                // This will grab only the parts of the description that are available and catch out when theres none left
                try
                {
                    // Assign the different parts of string to variables, first part starts at part0
                    part1 = parts[0];
                    elementList.Add(part1.ToString());
                    partCount++;
                    part2 = parts[1];
                    elementList.Add(part2.ToString());
                    partCount++;
                    part3 = parts[2];
                    elementList.Add(part3.ToString());
                    partCount++;
                    part4 = parts[3];
                    elementList.Add(part4.ToString());
                    partCount++;
                    part5 = parts[4];
                    elementList.Add(part5.ToString());
                    partCount++;
                    part6 = parts[5];
                    elementList.Add(part6.ToString());
                    partCount++;
                    part7 = parts[6];
                    elementList.Add(part7.ToString());
                    partCount++;
                    part8 = parts[7];
                    elementList.Add(part8.ToString());
                    partCount++;
                    part9 = parts[8];
                    elementList.Add(part9.ToString());
                    partCount++;
                    part10 = parts[9];
                    elementList.Add(part10.ToString());
                    partCount++;
                    part11 = parts[10];
                    elementList.Add(part11.ToString());
                    partCount++;
                    part12 = parts[11];
                    elementList.Add(part12.ToString());
                    partCount++;
                    part13 = parts[12];
                    elementList.Add(part13.ToString());
                    partCount++;
                    part14 = parts[13];
                    elementList.Add(part14.ToString());
                    partCount++;
                    part15 = parts[14];
                    elementList.Add(part15.ToString());
                    partCount++;
                    part16 = parts[15];
                    elementList.Add(part16.ToString());
                    partCount++;
                    part17 = parts[16];
                    elementList.Add(part17.ToString());
                    partCount++;
                    part18 = parts[17];
                    elementList.Add(part18.ToString());
                    partCount++;
                    part19 = parts[18];
                    elementList.Add(part19.ToString());
                    partCount++;
                    part20 = parts[19];
                    elementList.Add(part20.ToString());
                }
                catch { } // Leave the block when there are no more objects in the array (End of part description string)
                nbrLamps = 0;
                string seriesNum = "Part: " + part1;

                // Check if part is in ud100, if it isnt print error label and continue to next CSV

                if (validRecords != "0")
                    CheckIfValidPart1();

                //if (validRecords == "0")
                //{
                    // continue;
                //}
                
                // **** GET ALL INFO FOR LABELS ***** //
                CheckForUL();
                CheckIfStudo();
                GetIpClass();
                GetSeriesInfo();
                GetAddressInfo();
                GetVoltsInfo();

                /* JWM 04/09/2019 Changed code to get just the Lamp text for under cabinet parts */
                if (ucbFound == "Y")
                {
                    GetLampInfo(lampCode1);
                    prtLamp1 = prtLamp;
                }
                else
                    GetLampCodes();

                GetWattsInfo();

                // We use the counter so that the csv file doesnt grab header column
                counter++;

                // PRINT LABELS IF THERE IS A VALID UD100 Job Associated
                if (validRecords != "0")
                {
                    // Print Header Label
                    //   PrintHeaderLabel();

                    // Print Material Label if large label
                    //if (labelSize == "1")
                    //    PrintMaterialLabel();

                    // Convert Amount to integer
                    userAmount = Convert.ToInt32(tbLabelCount.Text);
                  //  userAmount = 1;

                    // Print all labels required for jobs
                    for (int i = 1; i <= userAmount; i++)
                    {
                        PrintLabels();
                        // If this is a UL job then print the UL Label as well
                        //if (needsULLabel == true)     /* 04/17/2019 JWM - Created new method for Exit Labels */
                        //{
                        //    StringBuilder sbUL = new StringBuilder();
                        //    sbUL.AppendLine((char)(2) + "L"); // tells the printer to start executing the command
                        //    sbUL.AppendLine("Q0001");
                        //    sbUL.AppendLine("D11"); // Dot size
                        //    sbUL.AppendLine("1Y1100000300020UL150W");
                        //    sbUL.AppendLine("190200100750100" + "EXIT FIXTURE 531X - WALL OR");
                        //    sbUL.AppendLine("190200100600100" + "COVERED CEILING MOUNT ONLY - SUITABLE FOR");
                        //    sbUL.AppendLine("190200100450100" + "WET LOCATIONS - IEC 598 CLASSIFIED IP65");
                        //    sbUL.AppendLine("190200100300100" + "SUITABLE FOR FLOOR PROXIMITY INSTALLATION");
                        //    sbUL.AppendLine("1X1100000250000L395001");
                        //    sbUL.AppendLine("1911A0800110005" + jobPartDesc);
                        //    sbUL.AppendLine("1911A0800110170" + jobPartNum);
                        //   // sbUL.AppendLine("1911A0800110255" + "JOB: ");
                        //  //  sbUL.AppendLine("1911A0800110280" + jobValue.ToString());
                        //    sbUL.AppendLine("1911A0800000005" + "KENALL MFG CO KENOSHA,WI 53144 MADE IN USA");
                        //    //  sbUL.AppendLine("1911A0800000290" + mtlPartNum);
                        //    sbUL.AppendLine("E");
                        //    RawPrinterHelper.SendStringToPrinter(SmallLabelPrint, sbUL.ToString());
                        //}
                    }
                }

                // PrintLastLabel();
                tbPartNum.Text = "";
                // tbQuantity.Text = "";
            }
            else
            {
                SqlCommand commPartDesc = new SqlCommand();
                // SqlCommand commDueDate = new SqlCommand();
                // SqlCommand commPartNum = new SqlCommand();
                // SqlCommand commName = new SqlCommand();
                // SqlCommand commProdQty = new SqlCommand();

                commPartDesc.Connection = new SqlConnection(sqlConnString);
                // commPartNum.Connection = new SqlConnection(sqlConnString);
                // commName.Connection = new SqlConnection(sqlConnString);
                // commDueDate.Connection = new SqlConnection(sqlConnString);
                // commProdQty.Connection = new SqlConnection(sqlConnString);

                // MAIN SQL STATEMENT - SQL JOINS etc..
                String sqlPartDesc = "Select PartDescription FROM " + eServer + $@".dbo.Part WHERE Company = 'KEN' AND PartNum = '" + tbPartNum.Text + "'";
                // String sqlDueDate = "Select DueDate FROM dbo.JobHead WHERE JobNum = '" + jobValue + "'";
                // String sqlName = "Select T2.Name FROM dbo.JobHead as T1 INNER JOIN Erp.Person as T2 ON T2.Company = T1.Company AND T2.PersonID = T1.PersonID  WHERE JobNum = '" + jobValue + "'";
                // String sqlProdQty = "Select ProdQty FROM dbo.JobHead WHERE JobNum = '" + jobValue + "'";

                // Part Desc
                commPartDesc.CommandText = sqlPartDesc;
                commPartDesc.Connection.Open();
                partDesc = commPartDesc.ExecuteScalar().ToString();
                commPartDesc.Connection.Close();

                //// Prod Qty
                //commProdQty.CommandText = sqlProdQty;
                //commProdQty.Connection.Open();
                //prodQty = Convert.ToInt32(commProdQty.ExecuteScalar());
                //commProdQty.Connection.Close();

                // Part Num
                partNum = tbPartNum.Text;

                //// Name
                //commName.CommandText = sqlName;
                //commName.Connection.Open();
                //plannerName = commName.ExecuteScalar().ToString();
                //commName.Connection.Close();

                // Due Date
                //commDueDate.CommandText = sqlDueDate;
                //commDueDate.Connection.Open();
                //jobDueDate = commDueDate.ExecuteScalar().ToString();
                //commDueDate.Connection.Close();

                //GetUlmsgInfo();
                // Parse Quantity
                // Convert Amount to integer
                userAmount = Convert.ToInt32(tbLabelCount.Text);

                // Send Cover Label
                StringBuilder sbCover = new StringBuilder();
                sbCover.AppendLine((char)(2) + "L"); // tells the printer to start executing the command
                sbCover.AppendLine("Q0001");
                sbCover.AppendLine("D11"); // Dot size
                // sbCover.AppendLine("132100000700100" + "Job: " + jobValue.ToString());
                // sbCover.AppendLine("132100000500030" + "Due Date: ");
                sbCover.AppendLine("E");
                RawPrinterHelper.SendStringToPrinter(SmallLabelPrint, sbCover.ToString());

                for (int i = 1; i <= userAmount; i++)
                {
                    // Print Labels for this Job
                    StringBuilder sbLabels = new StringBuilder();
                    sbLabels.AppendLine((char)(2) + "L"); // tells the printer to start executing the command
                    sbLabels.AppendLine("Q0001");
                    sbLabels.AppendLine("D11"); // Dot size
                    sbLabels.AppendLine("193300000720005" + "KENALL: ");
                    sbLabels.AppendLine("193300000720095" + partNum.ToString());
                    sbLabels.AppendLine("192200000600005" + partDesc.ToString());
                    //sbLabels.AppendLine("193300000420005" + jobValue.ToString());
                    sbLabels.AppendLine("192200000050005" + "Kenall Mfg Co.  Kenosha, WI 53144  USA");
                    sbLabels.AppendLine("192200000050290" + "F-0862");
                    sbLabels.AppendLine("E");
                    RawPrinterHelper.SendStringToPrinter(SmallLabelPrint, sbLabels.ToString());
                }
            }
            tbLabelCount.Text = "1";
        }

        // MAIN LABEL PRINT - Small Labels
        public void PrintSmallLabel()
        {
            int MsgLine = 0;

            List<string> ULMessageList = new List<string>();
            StringBuilder sb = new StringBuilder();

            // ** PRINT JOB LABELS MAIN ** //
            sb.AppendLine((char)(2) + "L"); // tells the printer to start executing the command
            sb.AppendLine("Q0001");
            sb.AppendLine("D11"); // Dot size

            // Part Num
            if (IndigoCleanPart == true)
            {
                sb.AppendLine("192200000800005KENALL: " + jobPartNum);
                sb.AppendLine("192200000680005" + jobPartDesc);
                //sb.AppendLine("192200000000005" + "Job: " + jobValue);
            }
            else
            {
                sb.AppendLine("193300000720005" + "KENALL:");
                sb.AppendLine("193300000720095" + jobPartNum);

                // Part Desc
                sb.AppendLine("192200000600005" + jobPartDesc);

                // Address Info
                // sb.AppendLine("192200000000005" + "Job: " + jobValue);
            }

            if (isSTUDO == true || isSTUDX == true)
            {
                if (IndigoCleanPart == true)
                {
                    sb.AppendLine("192100000620005" + "Finishing Section For Use With Rough-In");
                    sb.AppendLine("192100000540005" + "Section" + studioText);
                    sb.AppendLine("192100000460005" + "Section De Finition A Utiliser Avec Section");
                    sb.AppendLine("192100000380005" + "Non Finie " + studioText);
                }
                else
                {
                    sb.AppendLine("192100000510005" + "Finishing Section For Use With Rough-In");
                    sb.AppendLine("192100000430005" + "Section" + studioText);
                    sb.AppendLine("192100000350005" + "Section De Finition A Utiliser Avec Section");
                    sb.AppendLine("192100000280005" + "Non Finie " + studioText);
                }
            }

            if (isSTUDZ == true)
            {
                if (SSCS_PartNum != "SSCS")
                {
                    // Volts
                    sb.AppendLine("192200000480005" + "Volts:");
                    sb.AppendLine("192200000480045" + wrkvolts);

                    // Watts
                    //if (wrklamptype == "C" && SSCS_PartNum != "SSCS")
                    //{
                    sb.AppendLine("192200000480110" + "Watts:");
                    sb.AppendLine("192200000480150" + wrkwatts);
                    //}

                    // Hertz
                    sb.AppendLine("192200000480190" + "Hz:");
                    sb.AppendLine("192200000480215" + wrkhza);

                    // ip class
                    if (ipClass != "")
                    {
                        sb.AppendLine("192200000480260" + "IP Class:");
                        sb.AppendLine("192200000480320" + ipClass);
                    }

                    // Lamp Code
                    if (lampCode1 != "")
                    {
                        sb.AppendLine("192200000350005" + "Lamp:");
                        sb.AppendLine("192200000350045" + prtLamp1);
                    }
                }
            }

            // Address Info
            if (IndigoCleanPart == true)
            {
                sb.AppendLine("192100000300005" + wrkADDR1);
                sb.AppendLine("192100000220005" + wrkADDR2);
                sb.AppendLine("192100000140005" + wrkADDR3);
                sb.AppendLine("192200000000280" + DateTime.Now.ToShortDateString());
            }
            else
            {
                sb.AppendLine("192200000140005" + wrkADDR1);
                //sb.AppendLine("191200000000124" + "MADE & ASSEMBLED IN USA");
                sb.AppendLine("192200000000285" + DateTime.Now.ToShortDateString());
            }

            sb.AppendLine("E");
            RawPrinterHelper.SendStringToPrinter(SmallLabelPrint, sb.ToString());

            if (isSTUDZ == true)
            {
                // ** PRINT JOB LABELS MAIN ** //
                //StringBuilder sb2 = new StringBuilder();
                sb.Clear();

                sb.AppendLine((char)(2) + "L"); // tells the printer to start executing the command
                sb.AppendLine("Q0001");
                sb.AppendLine("D11"); // Dot size

                // Part Num
                sb.AppendLine("192200000820005" + "KENALL:");
                sb.AppendLine("192200000820095" + jobPartNum);

                // Part Desc
                sb.AppendLine("192200000700005" + jobPartDesc);

                // Get Work Messages
                // Create connection - State Database Connection, Main SQL SELECT
                SqlCommand comm = new SqlCommand();
                comm.Connection = new SqlConnection(sqlConnString);

                // MAIN SQL SELECT STATEMENT
                String sql = "SELECT Key1, Key2, ShortChar01, ShortChar02 FROM " + eServer + $@".Ice.UD100A WHERE Company = 'KEN' AND Key1 = '" + part1 + "'";

                // Assign CommandText and Open DB Connections
                comm.CommandText = sql;
                comm.Connection.Open();

                // Reset Msg Count if printing new label
                //msgCount = 0;

                // Create Reader
                SqlDataReader sqlReader = comm.ExecuteReader();
                
                // While Reading from SQL - Write SQL data to Text File
                while (sqlReader.Read())
                {
                    currentID = sqlReader["ShortChar01"].ToString();
                    if (currentID != "")
                    {
                        SqlCommand comm2 = new SqlCommand();
                        comm2.Connection = new SqlConnection(sqlConnString);

                        // MAIN SQL SELECT STATEMENT
                        String sql2 = @"SELECT Key1,
                                               Key2,
                                               Character01,
                                               Character02,
                                               Character03,
                                               ShortChar03
                                          FROM " + eServer + $@".Ice.UD01
                                         WHERE Company = 'KEN'
                                           AND Key1 = 'ULMsgMst'
                                           AND Key2 = '" + currentID + "'";

                        // Assign CommandText and Open DB Connections
                        comm2.CommandText = sql2;
                        comm2.Connection.Open();
                        SqlDataReader sqlReader2 = comm2.ExecuteReader();

                        while (sqlReader2.Read())
                        {
                            //if (msgCount == 0)
                            //    sb2.AppendLine("111200000600005" + sqlReader2["Character01"].ToString() + " " + sqlReader2["Character02"].ToString());
                            //if (msgCount == 1)
                            //    sb2.AppendLine("111200000500005" + sqlReader2["Character01"].ToString() + " " + sqlReader2["Character02"].ToString());
                            //if (msgCount == 2)
                            //    sb2.AppendLine("111200000400005" + sqlReader2["Character01"].ToString() + " " + sqlReader2["Character02"].ToString());
                            //if (msgCount == 3)
                            //    sb2.AppendLine("111200000300005" + sqlReader2["Character01"].ToString() + " " + sqlReader2["Character02"].ToString());
                            //if (msgCount == 4)
                            //    sb2.AppendLine("111200000200005" + sqlReader2["Character01"].ToString() + " " + sqlReader2["Character02"].ToString());
                            //if (msgCount == 5)
                            //    sb2.AppendLine("111200000100005" + sqlReader2["Character01"].ToString() + " " + sqlReader2["Character02"].ToString());
                            //msgCount++;
                            ULMessageList.Add(sqlReader2["Character01"].ToString());

                            if (sqlReader2["Character02"].ToString() != "")
                                ULMessageList.Add(sqlReader2["Character02"].ToString());
                        }
                        sqlReader2.Close();
                        comm2.Connection.Close();
                    }
                }
                // Close Connections
                sqlReader.Close();
                comm.Connection.Close();

                // Print label if we have Message lines
                if (ULMessageList.Count > 0)
                {
                    // Set starting line for label messages
                    MsgLine = 60;

                    foreach (string MessageLine in ULMessageList)
                    {
                        sb.AppendLine($@"1211000{MsgLine.ToString("0000")}0005{MessageLine}");
                        MsgLine -= 10;
                    }
                }

                sb.AppendLine("E");
                RawPrinterHelper.SendStringToPrinter(SmallLabelPrint, sb.ToString());
            }
        }

        public void PrintExitLabel()
        {
            int MsgLine = 0;
            int MsgLineCount = 0;
            string LabelType = "";
            string sql = "";
            string sql2 = "";
            string sql3 = "";

            List<string> ULMessageList = new List<string>();

            StringBuilder sb = new StringBuilder();

            // Create connection - State Database Connection, Main SQL SELECT
            SqlCommand comm = new SqlCommand();
            SqlCommand comm2 = new SqlCommand();
            SqlCommand comm3 = new SqlCommand();
            comm.Connection = new SqlConnection(sqlConnString);
            comm2.Connection = new SqlConnection(sqlConnString);
            comm3.Connection = new SqlConnection(sqlConnString);

            // MAIN SQL SELECT STATEMENT
            sql = $@"Select T1.MtlPartNum FROM " + eServer + $@".Erp.PartMtl as T1 WHERE T1.Company = 'KEN' AND T1.PartNum = '{jobPartNum}' AND T1.MtlPartNum LIKE 'f-%'";

            // Assign CommandText and Open DB Connections
            comm.CommandText = sql;
            comm.Connection.Open();
            SqlDataReader sqlReader = comm.ExecuteReader();

            while (sqlReader.Read())
            {
                sql2 = $@"SELECT ShortChar01, ShortChar02 FROM " + eServer + $@".Ice.UD100A WHERE Company = 'KEN' AND Key1 = '{part1}' AND ShortChar01 LIKE '{sqlReader["MtlPartNum"].ToString()}%'";

                // Clear the Message List and the label string
                sb.Clear();
                ULMessageList.Clear();

                // Assign CommandText and Open DB Connections
                comm2.CommandText = sql2;
                comm2.Connection.Open();
                SqlDataReader sqlReader2 = comm2.ExecuteReader();

                while (sqlReader2.Read())
                {
                    LabelType = sqlReader2["ShortChar02"].ToString();

                    sql3 = $@"SELECT Character01, Character02 FROM " + eServer + $@".Ice.UD01 WHERE Company = 'KEN' AND Key1 = 'ULMsgMst' AND Key2 = '{sqlReader2["ShortChar01"].ToString()}'";

                    // Assign CommandText and Open DB Connections
                    comm3.CommandText = sql3;
                    comm3.Connection.Open();
                    SqlDataReader sqlReader3 = comm3.ExecuteReader();

                    while (sqlReader3.Read())
                    {
                        ULMessageList.Add(sqlReader3["Character01"].ToString());

                        if (sqlReader3["Character02"].ToString() != "")
                            ULMessageList.Add(sqlReader3["Character02"].ToString());
                    }
                    sqlReader3.Close();
                    comm3.Connection.Close();
                }
                sqlReader2.Close();
                comm2.Connection.Close();

                // Print label if we have Message lines
                if (ULMessageList.Count > 0)
                {
                    // Set starting line for label messages
                    MsgLine = 75;
                    MsgLineCount = 0;

                    sb.AppendLine((char)(2) + "L"); // tells the printer to start executing the command
                    sb.AppendLine("Q0001");
                    sb.AppendLine("D11"); // Dot size

                    foreach (string MessageLine in ULMessageList)
                    {
                        if (LabelType == "MSG")
                        {
                            MsgLineCount++;
                            if (MsgLineCount == 1 || MsgLineCount == 5)
                                sb.AppendLine($@"1902001{MsgLine.ToString("0000")}0005{MessageLine}");
                            else
                                sb.AppendLine($@"1902001{MsgLine.ToString("0000")}0050{MessageLine}");
                        }

                        if (LabelType == "UL")
                        {
                            sb.AppendLine($@"1902001{MsgLine.ToString("0000")}0100{MessageLine}");
                        }
                        MsgLine -= 15;
                    }

                    if (LabelType == "UL")
                    {
                        sb.AppendLine(@"1Y1100000300020UL150W");    /* UL Logo */
                        sb.AppendLine($@"1911A0800110005{jobPartDesc}");
                        sb.AppendLine($@"1911A0800110170{jobPartNum}");
                        // sb.AppendLine($@"1911A0800110255JOB: {jobValue}");
                        sb.AppendLine($@"1911A0800000005{wrkADDR1}");
                        sb.AppendLine($@"1911A0800000235{DateTime.Now.ToString("MM/dd/yyyy")}");
                        sb.AppendLine($@"1911A0800000295{sqlReader["MtlPartNum"].ToString()}");
                    }

                    sb.AppendLine("E");
                    RawPrinterHelper.SendStringToPrinter(SmallLabelPrint, sb.ToString());
                }
            }

            sqlReader.Close();
            comm.Connection.Close();
        }

        // MAIN LABEL PRINT - Large Labels
        public void PrintLargeLabel()
        {
            //string matLoc1 = "121100001160024";
            //string matLoc2 = "121100001040024";
            //string matLoc3 = "121100000920024";
            //string matLoc4 = "121100000800024";
            //string matLoc5 = "121100000680024";
            //string matLoc6 = "121100000560024";
            //string matLoc7 = "121100000440024";
            //string matLoc8 = "121100000320024";

            //bool slotDbl1 = false;
            //bool slotDbl2 = false;
            //bool slotDbl3 = false;
            //bool slotDbl4 = false;
            //bool slotDbl5 = false;
            //bool slotDbl6 = false;
            //bool slotDbl7 = false;
            //bool slotDbl8 = false;

            // ** PRINT JOB LABELS MAIN ** //
            printCounter = 1;
            msgCount = 0;

            // Get messages for the middle section of label
            GetWorkMessages();

            StringBuilder sb = new StringBuilder();

            // Append the top part of the label
            sb = PrintLargeLabelTop(sb);

            /* 04/18/2019 JWM - Section moved to PrintLargeLabelTop */
            //sb.AppendLine((char)(2) + "L"); // tells the printer to start executing the command
            //sb.AppendLine("Q0001");
            //sb.AppendLine("D11"); // Dot size

            //// Part Num
            //sb.AppendLine("122200002150005" + "Part:");
            //sb.AppendLine("122200002150072" + jobPartNum);

            //// Part Desc
            //sb.AppendLine("121100002010005" + jobPartDesc);

            //// Volts
            //sb.AppendLine("192200001850005" + "Volts:");
            //sb.AppendLine("192200001850049" + wrkvolts);

            //// Watts
            ////if (wrklamptype == "C" && SSCS_PartNum != "SSCS")
            ////{
            //sb.AppendLine("192200001850259" + "Watts:");
            //sb.AppendLine("192200001850301" + wrkwatts);
            ////}

            //// Hertz
            //sb.AppendLine("192200001850330" + "Hz:");
            //sb.AppendLine("192200001850354" + wrkhza);

            ////IpClass
            //if (ipClass != "")
            //{
            //    sb.AppendLine("192200001730315" + "Ip Class:");
            //    sb.AppendLine("192200001730373" + ipClass);
            //}

            //// Lamp Codes - May be more than 1??
            //sb.AppendLine("1911A0801590005" + "Lamp:");
            //sb.AppendLine("192200001590043" + prtLamp1);

            /* 04/18/2019 JWM - Section moved to PrintLargeLabelBottom */
            //// Address Info Line 1
            //if (AES_PartNum != "AES")
            //    sb.AppendLine("1911A0800180010" + wrkADDR);
            //else if (AES_PartNum == "AES")
            //{
            //    wrkADDR = "AES Clean Technology, Inc.";
            //    sb.AppendLine("1911A0800180130" + wrkADDR);
            //}
            //// Address Info Line 2
            //if (AES_PartNum != "AES")
            //{
            //    sb.AppendLine("1911A0800050010" + "Job: " + jobValue);
            //    sb.AppendLine("1911A0800050150" + "MADE IN USA");
            //    sb.AppendLine("1911A0800050300" + DateTime.Now.ToShortDateString());
            //}
            //else if (AES_PartNum == "AES")
            //{
            //    sb.AppendLine("1911A0800050125" + "AES WOLF TOP Access Light");
            //}
            //// Write Label Count Number
            ////if (multipleLabelsNeeded == true)
            //sb.AppendLine("122200002150360" + "1");

            /* 04/18/2019 JWM - Section moved to GetWorkMessages */
            //// Get Work Messages
            //// Create connection - State Database Connection, Main SQL SELECT
            //SqlCommand comm = new SqlCommand();
            //comm.Connection = new SqlConnection(sqlConnString);

            //// MAIN SQL SELECT STATEMENT
            //String sql = @"SELECT Key1,
            //                      Key2,
            //                      ShortChar01,
            //                      ShortChar02,
            //                      ShortChar03
            //                 FROM Ice.UD100A
            //                WHERE Company = 'KEN'
            //                  AND Key1 = '" + part1 + "'";

            //// Assign CommandText and Open DB Connections
            //comm.CommandText = sql;
            //comm.Connection.Open();

            //// Create Reader
            //SqlDataReader sqlReader = comm.ExecuteReader();
            //msgCount = 0;

            //// While Reading from SQL - Write SQL data to Text File
            //while (sqlReader.Read())
            //{
            //    // SEE if there is an Optional UL Msg ID for this, if there is use it instead of ShortChar01 field
            //    string optionalID = sqlReader["ShortChar03"].ToString();

            //    if (optionalID == "")
            //        currentID = sqlReader["ShortChar01"].ToString();
            //    else
            //        currentID = optionalID;

            //    // If valid ID
            //    if (currentID != "")
            //    {
            //        SqlCommand comm2 = new SqlCommand();
            //        comm2.Connection = new SqlConnection(sqlConnString);

            //        // Get the verbage data from ULMsgMst table (Ice.UD01)
            //        String sql2 = @"SELECT Key1,
            //                               Key2,
            //                               Character01,
            //                               Character02,
            //                               Character03,
            //                               ShortChar03
            //                          FROM Ice.UD01
            //                         WHERE Company = 'KEN'
            //                           AND Key1 = 'ULMsgMst'
            //                           AND Key2 = '" + currentID + "'";

            //        // Assign CommandText and Open DB Connections
            //        comm2.CommandText = sql2;
            //        comm2.Connection.Open();
            //        SqlDataReader sqlReader2 = comm2.ExecuteReader();

            //        while (sqlReader2.Read())
            //        {
            //            string message = "";
            //            string char1 = sqlReader2["Character01"].ToString();
            //            string char2 = sqlReader2["Character02"].ToString();

            //            // Get message verbage from UD01 character01 + character02 fields
            //            message = char1 + " " + char2;

            //            if (msgCount < 8)
            //            {
            //                if (msgCount == 0)
            //                {
            //                    sb.AppendLine(matLoc1 + message);
            //                }
            //                if (msgCount == 1)
            //                {
            //                    sb.AppendLine(matLoc2 + message);
            //                }
            //                if (msgCount == 2)
            //                {
            //                    sb.AppendLine(matLoc3 + message);
            //                }
            //                if (msgCount == 3)
            //                {
            //                    sb.AppendLine(matLoc4 + message);
            //                }
            //                if (msgCount == 4)
            //                {
            //                    sb.AppendLine(matLoc5 + message);
            //                }
            //                if (msgCount == 5)
            //                {
            //                    sb.AppendLine(matLoc6 + message);
            //                }
            //                if (msgCount == 6)
            //                {
            //                    sb.AppendLine(matLoc7 + message);
            //                }
            //                if (msgCount == 7)
            //                {
            //                    sb.AppendLine(matLoc8 + message);
            //                }
            //                msgCount++;
            //            }
            //        }
            //        sqlReader2.Close();
            //        comm2.Connection.Close();
            //    }
            //}
            //// Close Connections
            //sqlReader.Close();
            //comm.Connection.Close();

            //sb.AppendLine("E");
            //RawPrinterHelper.SendStringToPrinter(LargeLabelPrint, sb.ToString());

            // Set beginning vertical position of middle section
            //int MsgLine = 116;
            int MsgLine = 100;

            /* Go through the MessagesList, appending up to 8 lines at a time to the middle section. */
            /* If more than 8 lines exist in the list, append the bottom label section and send the label to the printer, */
            /* then start a new */
            foreach (string MsgRow in MessagesList)
            {
                sb.AppendLine($@"1211000{MsgLine.ToString("0000")}0024{MsgRow}");
                //MsgLine -= 12;
                MsgLine -= 10;
                msgCount++;

                /* If the msgCount is an even multiple of 8, print the bottom portion of the label and reset for the next label */
                if (msgCount % 8 == 0)
                {
                    PrintLargeLabelBottom(sb);
                    printCounter++;
                    //MsgLine = 116;
                    MsgLine = 100;

                    /* If we are not at the end of the messages list, print the top portion of the label again */
                    if (msgCount < MessagesList.Count)
                        sb = PrintLargeLabelTop(sb);
                }
            }

            PrintLargeLabelBottom(sb);

            /* 04/18/2019 JWM - Section not needed, multiple labels handled in foreach loop above */
            // If message count is greater than 7 lines then print a second label with the rest
            //if (msgCount == 8)
            //{
            //    multipleLabelsNeeded = true;
            //    // ** PRINT JOB LABELS MAIN ** //
            //    StringBuilder sb3 = new StringBuilder();
            //    sb3.AppendLine((char)(2) + "L"); // tells the printer to start executing the command
            //    sb3.AppendLine("Q0001");
            //    sb3.AppendLine("D11"); // Dot size

            //    // Label Count
            //    sb3.AppendLine("122200002150360" + "2");

            //    // Part Num
            //    sb3.AppendLine("122200002150005" + "Part:");
            //    sb3.AppendLine("122200002150072" + jobPartNum);

            //    // Part Desc
            //    sb3.AppendLine("121100002010005" + jobPartDesc);

            //    // Volts
            //    sb3.AppendLine("192200001850005" + "Volts:");
            //    sb3.AppendLine("192200001850049" + wrkvolts);

            //    // Watts
            //    //if (wrklamptype == "C" && SSCS_PartNum != "SSCS")
            //    //{
            //    sb3.AppendLine("192200001850259" + "Watts:");
            //    sb3.AppendLine("192200001850301" + wrkwatts);
            //    //}

            //    // Hertz
            //    sb3.AppendLine("192200001850330" + "Hz:");
            //    sb3.AppendLine("192200001850354" + wrkhza);

            //    //IpClass
            //    if (ipClass != "")
            //    {
            //        sb3.AppendLine("192200001730315" + "Ip Class:");
            //        sb3.AppendLine("192200001730373" + ipClass);
            //    }

            //    // Lamp Codes - May be more than 1??
            //    sb3.AppendLine("1911A0801590005" + "Lamp:");
            //    sb3.AppendLine("192200001590043" + prtLamp1);

            //    // Address Info Line 1
            //    if (AES_PartNum != "AES")
            //        sb3.AppendLine("1911A0800180010" + wrkADDR);
            //    else if (AES_PartNum == "AES")
            //    {
            //        wrkADDR = "AES Clean Technology, Inc.";
            //        sb3.AppendLine("1911A0800180130" + wrkADDR);
            //    }
            //    // Address Info Line 2
            //    if (AES_PartNum != "AES")
            //    {
            //        sb3.AppendLine("1911A0800050010" + "Job: " + jobValue);
            //        sb3.AppendLine("1911A0800050150" + "MADE IN USA");
            //        sb3.AppendLine("1911A0800050300" + DateTime.Now.ToShortDateString());
            //    }
            //    else if (AES_PartNum == "AES")
            //    {
            //        sb3.AppendLine("1911A0800050125" + "AES WOLF TOP Access Light");
            //    }

            //    // Get Work Messages
            //    // Create connection - State Database Connection, Main SQL SELECT
            //    SqlCommand comm3 = new SqlCommand();
            //    comm3.Connection = new SqlConnection(sqlConnString);

            //    // MAIN SQL SELECT STATEMENT
            //    String sql3 = @"SELECT Key1,
            //                           Key2,
            //                           ShortChar01,
            //                           ShortChar02,
            //                           ShortChar03
            //                      FROM Ice.UD100A
            //                     WHERE Company = 'KEN'
            //                       AND Key1 = '" + part1 + 
            //              @"' ORDER BY Key1 OFFSET 7 ROWS FETCH NEXT 20 ROWS ONLY;";

            //    // Assign CommandText and Open DB Connections
            //    comm3.CommandText = sql3;
            //    comm3.Connection.Open();

            //    // Create Reader
            //    SqlDataReader sqlReader3 = comm3.ExecuteReader();
            //    // While Reading from SQL - Write SQL data to Text File
            //    while (sqlReader3.Read())
            //    {
            //        if (currentID != "")
            //        {
            //            SqlCommand comm4 = new SqlCommand();
            //            comm4.Connection = new SqlConnection(sqlConnString);

            //            // MAIN SQL SELECT STATEMENT
            //            String sql4 = @"SELECT Key1,
            //                                   Key2,
            //                                   Character01,
            //                                   Character02,
            //                                   Character03,
            //                                   ShortChar03
            //                              FROM Ice.UD01
            //                             WHERE Company = 'KEN'
            //                               AND Key1 = 'ULMsgMst'
            //                               AND Key2 = '" + currentID + "'";

            //            // Assign CommandText and Open DB Connections
            //            comm4.CommandText = sql4;
            //            comm4.Connection.Open();
            //            SqlDataReader sqlReader4 = comm4.ExecuteReader();
            //            //int currentCount = 0;

            //            while (sqlReader4.Read())
            //            {
            //                // currentCount++;
            //                //if (currentCount > msgCount)
            //                //{
            //                if (msgCount == 7)
            //                    sb3.AppendLine("121100001160024" + sqlReader4["Character01"].ToString() + " " + sqlReader4["Character02"].ToString());
            //                if (msgCount == 8)
            //                    sb3.AppendLine("121100001040024" + sqlReader4["Character01"].ToString() + " " + sqlReader4["Character02"].ToString());
            //                if (msgCount == 9)
            //                    sb3.AppendLine("121100000920024" + sqlReader4["Character01"].ToString() + " " + sqlReader4["Character02"].ToString());
            //                if (msgCount == 10)
            //                    sb3.AppendLine("121100000800024" + sqlReader4["Character01"].ToString() + " " + sqlReader4["Character02"].ToString());
            //                if (msgCount == 11)
            //                    sb3.AppendLine("121100000680024" + sqlReader4["Character01"].ToString() + " " + sqlReader4["Character02"].ToString());
            //                if (msgCount == 12)
            //                    sb3.AppendLine("121100000560024" + sqlReader4["Character01"].ToString() + " " + sqlReader4["Character02"].ToString());
            //                if (msgCount == 13)
            //                    sb3.AppendLine("121100000440024" + sqlReader4["Character01"].ToString() + " " + sqlReader4["Character02"].ToString());
            //                if (msgCount == 14)
            //                    sb3.AppendLine("121100000320024" + sqlReader4["Character01"].ToString() + " " + sqlReader4["Character02"].ToString());
            //                //}
            //                msgCount++;
            //            }
            //            sqlReader4.Close();
            //            comm4.Connection.Close();
            //        }
            //    }
            //    // Close Connections
            //    sqlReader3.Close();
            //    comm3.Connection.Close();

            //    sb3.AppendLine("E");
            //    RawPrinterHelper.SendStringToPrinter(LargeLabelPrint, sb3.ToString());
            //}
        }

        public StringBuilder PrintLargeLabelTop(StringBuilder sb)
        {
            sb.AppendLine((char)(2) + "L"); // tells the printer to start executing the command
            sb.AppendLine("Q0001");
            sb.AppendLine("D11"); // Dot size

            // Write Label Count Number
            if (MessagesList.Count() > 8)
                sb.AppendLine("122200002150360" + printCounter.ToString("0"));

            // Part Num
            sb.AppendLine("122200002150005" + "Part:");
            sb.AppendLine("122200002150072" + jobPartNum);

            // Part Desc
            sb.AppendLine("121100002010005" + jobPartDesc);

            // Volts
            sb.AppendLine("192200001850005" + "Volts:");
            sb.AppendLine("192200001850049" + wrkvolts);

            // Watts
            //sb.AppendLine("192200001850259" + "Watts:");
            //sb.AppendLine("192200001850301" + wrkwatts);
            sb.AppendLine($@"192200001850200Watts: {wrkwatts}");

            // Hertz
            //sb.AppendLine("192200001850330" + "Hz:");
            //sb.AppendLine("192200001850354" + wrkhza);
            sb.AppendLine($@"192200001850310Hz: {wrkhza}");

            //IpClass
            if (ipClass != "")
            {
                //sb.AppendLine("192200001730315" + "Ip Class:");
                //sb.AppendLine("192200001730373" + ipClass);
                sb.AppendLine($@"192200001730200Ip Class: {ipClass}");
            }

            // Lamp Codes - May be more than 1??
            //sb.AppendLine("1911A0801590005" + "Lamp:");
            //sb.AppendLine("192200001590043" + prtLamp1);
            sb.AppendLine("1911A0801730005" + "Lamp: " + prtLamp1);
            //sb.AppendLine("1911A0801630042Lamp2");  /* Test */
            //sb.AppendLine("1911A0801530042Lamp3");  /* Test */

            if (wrkNbrLamps > 1)
                //sb.AppendLine("192200001460043" + prtLamp2);
                sb.AppendLine("1911A0801630042" + prtLamp2);

            if (wrkNbrLamps > 2)
                //sb.AppendLine("192200001330043" + prtLamp3);
                sb.AppendLine("191100001530042" + prtLamp3);

            return sb;
        }

        public void PrintLargeLabelBottom(StringBuilder sb)
        {
            // Address Info Line 1
            if (AES_PartNum != "AES")
            {
                if (IndigoCleanPart == true)
                {
                    sb.AppendLine("1911A0801370005" + wrkADDR1);
                    sb.AppendLine("1911A0801270005" + wrkADDR2);
                    sb.AppendLine("1911A0801170005" + wrkADDR3);
                }
                else
                    sb.AppendLine("1911A0801370010" + wrkADDR1);
            }
            else if (AES_PartNum == "AES")
            {
                wrkADDR1 = "AES Clean Technology, Inc.";
                sb.AppendLine("1911A0800180130" + wrkADDR1);
            }
            // Address Info Line 2
            if (AES_PartNum != "AES")
            {
                //sb.AppendLine("1911A0800050010" + "Job: " + jobValue);
                //sb.AppendLine("1911A0800050150" + "MADE IN USA");
                sb.AppendLine("1911A0800050295" + DateTime.Now.ToShortDateString());
            }
            else if (AES_PartNum == "AES")
            {
                sb.AppendLine("1911A0800050125" + "AES WOLF TOP Access Light");
            }

            sb.AppendLine("E");
            RawPrinterHelper.SendStringToPrinter(LargeLabelPrint, sb.ToString());

            // Clear the StringBuilder contents
            sb.Clear();
        }

        // Get Series Info - Done **
        public void GetSeriesInfo()
        {
            string sql = "";
            string ud01Key3 = "";
            string UD100AIndigoType = "";

            element = 0;
            addressCode = "";
            oldPart = "";
            lampCode1 = "";
            lampCode2 = "";
            lampCode3 = "";
            AES_PartNum = "";
            SSCS_PartNum = "";
            ud100Key = "";

            nbrLamps = 0;
            //wrkNbrLamps = 0;
            ucbElementCounter = 0;

            IndigoCleanPart = false;

            // Create connection - State Database Connection, Main SQL SELECT
            SqlCommand comm = new SqlCommand();
            comm.Connection = new SqlConnection(sqlConnString);

            // Main SQL SELECT STATEMENT - SQL JOINS ETC..
            //String sql = "SELECT * TOP 1 FROM dbo.CS_FixtureLookup WHERE ud100Key1 = '" + part1 + "'";

            sql = $@"SELECT T1.Key3 UD01Key3
                           ,T1.ShortChar01 ud01ShortChar01
                            FROM " + eServer + $@".Ice.UD01 as T1
                           WHERE T1.Company = 'KEN'
                             AND T1.Key1 = 'UCBMst'
                             AND T1.Key2 = '{part1}'";

            // Assign CommandText and Open DB Connections
            comm.CommandText = sql;
            comm.Connection.Open();

            // Create Reader
            SqlDataReader sqlReader = comm.ExecuteReader();

            ucbFound = "N";

            // While Reading from SQL - Write SQL data to Text File
            while (sqlReader.Read())
            {
                ud01Key3 = sqlReader["ud01Key3"].ToString();

                // See if this job has a UCBmst Record
                foreach (string item in elementList)
                {
                    if (ud01Key3 == item)
                    {
                        ucbFound = "Y";
                        nbrLamps = nbrLamps + 1;
                        lampCode1 = sqlReader["ud01ShortChar01"].ToString();
                        break;
                    }
                }
            }

            sqlReader.Close();

            if (ucbFound == "N")
            {
                SSCS_PartNum = part1;
                //SSCS_PartNum = part1.Substring(0, 4);
                ud100Key = "";

                if (SSCS_PartNum == "SSCS" && part2.Length >= 3)
                    ud100Key = SSCS_PartNum + part2.Substring(2, 1);
                else
                    ud100Key = part1;
            }

            sql = $@"SELECT T2.Key1 as ud100Key1
                           ,T2.ShortChar03 as ud100ShortChar03
                           ,T2.ShortChar01 as ud100ShortChar01
                           ,T2.Number01 as ud100Number01
                           ,T2.Number02 as ud100Number02
                           ,T2.ShortChar02 as ud100ShortChar02
                           ,T2.ShortChar04 as ud100ShortChar04
                            FROM " + eServer + $@".Ice.UD100 as T2
                           WHERE T2.Company = 'KEN'
                             AND T2.Key1 = '{ud100Key}'";

            // Assign CommandText and Open DB Connections
            comm.CommandText = sql;

            // Create Reader
            sqlReader = comm.ExecuteReader();

            while (sqlReader.Read())
            {
                ipClass = sqlReader["ud100ShortChar01"].ToString();
                element = Convert.ToInt32(sqlReader["ud100Number01"]);
                wrkNbrLamps = Convert.ToInt32(sqlReader["ud100Number02"]);
                oldPart = sqlReader["ud100ShortChar02"].ToString();
                addressCode = sqlReader["ud100ShortChar04"].ToString();
            }
            sqlReader.Close();

            if (addressCode == "INDIGO-CLEAN")
                IndigoCleanPart = true;
            else
            {
                sql = $@"SELECT T2.Key1 as ud100AKey1
                           ,T2.ShortChar02 as ud100AShortChar02
                           ,T2.ShortChar03 as ud100AShortChar03
                            FROM " + eServer + $@".Ice.UD100A as T2
                           WHERE T2.Company = 'KEN'
                             AND T2.Key1 = '{ud100Key}'
                             AND T2.ShortChar03 = 'INDIGO'";

                comm.CommandText = sql;
                // Create Reader
                sqlReader = comm.ExecuteReader();

                while (sqlReader.Read())
                {
                    UD100AIndigoType = sqlReader["ud100AShortChar02"].ToString().ToUpper();

                    if (partDesc.IndexOf(UD100AIndigoType, StringComparison.OrdinalIgnoreCase) > 0)
                        IndigoCleanPart = true;

                }
            }

            // FIND LAMP CODE - Go through all positions and identify if it is a lamp code  /* 04/18/2019 JWM - Not used */
            //foreach (string item in elementList)
            //{
            //    string elementItem = item.ToString();
            //    // Look for lamp code in UD01 Table
            //    if (ud01Key1 == "lampmst" && ud01Key2 == elementItem)
            //    {
            //        nbrLamps = nbrLamps + 1;
            //        if (lampCode1 == "")
            //            lampCode1 = item.ToString();
            //        else if (lampCode2 == "")
            //            lampCode2 = item.ToString();
            //        else if (lampCode3 == "")
            //            lampCode3 = item.ToString();
            //    }
            //}
            sqlReader.Close();
            comm.Connection.Close();
        }

        // Get Info for volts - Done **
        public void GetVoltsInfo()
        {
            wrkvolts = "";
            wrkhza = "";
            voltkey1 = "";
            voltkey2 = "";
            voltkey3 = "";
            voltkey4 = "";
            voltsFound = "NO";
            ucbElementCounter = 0;
            int listcount = elementList.Count;

            // Create connection - State Database Connection, Main SQL SELECT
            SqlCommand comm = new SqlCommand();
            comm.Connection = new SqlConnection(sqlConnString);

            // MAIN SQL SELECT STATEMENT
            String sql = @"SELECT T1.Key1 as ud01Key1,
                                  T1.Key2 as ud01Key2,
                                  T1.ShortChar01 as sc01,
                                  T1.ShortChar02 as sc02,
                                  T1.ShortChar03 as sc03,
                                  T1.ShortChar04 as sc04,
                                  T1.ShortChar05 as sc05,
                                  T1.ShortChar06 as sc06
                             FROM " + eServer + $@".Ice.UD01 as T1
                            WHERE T1.Company = 'KEN'
                              AND T1.Key1 = 'voltsmst'";

            // Assign CommandText and Open DB Connections
            comm.CommandText = sql;
            comm.Connection.Open();

            // Create Reader
            SqlDataReader sqlReader = comm.ExecuteReader();

            // While Reading from SQL - Write SQL data to Text File
            while (sqlReader.Read())
            {
                int numberOfVoltRecords = 0;
                // Variables
                string ud01Key1 = sqlReader["ud01Key1"].ToString();
                string ud01Key2 = sqlReader["ud01Key2"].ToString();
                string shortChar01 = sqlReader["sc01"].ToString();
                string shortChar02 = sqlReader["sc02"].ToString();
                string shortChar03 = sqlReader["sc03"].ToString();
                string shortChar04 = sqlReader["sc04"].ToString();
                string shortChar05 = sqlReader["sc05"].ToString();
                string shortChar06 = sqlReader["sc06"].ToString();

                // Get Volts
                foreach (string item in elementList)
                {
                    string elementItem = item.ToString();
                    if (ud01Key1 == "VoltsMst" && ud01Key2 == elementItem)
                    {
                        voltsFound = "YES";
                        wrkvolts = shortChar01;
                        wrkhza = shortChar02;
                        voltkey1 = shortChar03;
                        voltkey2 = shortChar04;
                        voltkey3 = shortChar05;
                        voltkey4 = shortChar06;
                        numberOfVoltRecords++;
                    }
                    ucbElementCounter++;
                }
            }
            comm.Connection.Close();
            sqlReader.Close();
        }

        // Get Lamp Code Info
        public void GetLampCodes()
        {
            SqlCommand comm2 = new SqlCommand();
            comm2.Connection = new SqlConnection(sqlConnString);

            // MAIN SQL SELECT STATEMENT
            String sql2 = "SELECT Number01 FROM " + eServer + $@".Ice.UD100 as T1 WHERE T1.Company = 'KEN' AND T1.Key1 = '" + part1 + "'";

            comm2.CommandText = sql2;
            comm2.Connection.Open();
            int startIndex = Convert.ToInt32(comm2.ExecuteScalar());
            comm2.Connection.Close();

            IEnumerable<string> lampList = elementList.Skip(startIndex - 1);

            lampCode1 = "";
            lampCode2 = "";
            lampCode3 = "";
            prtLamp1 = "";
            prtLamp2 = "";
            prtLamp3 = "";

            if (lampList.ToString() != "")
            {
                // Create connection - State Database Connection, Main SQL SELECT
                SqlCommand comm = new SqlCommand();
                comm.Connection = new SqlConnection(sqlConnString);

                // MAIN SQL SELECT STATEMENT
                String sql = "SELECT Key1, Key2 FROM " + eServer + $@".Ice.UD01 as T1 WHERE T1.Company = 'KEN' AND T1.Key1 = 'LampMst' ORDER BY Key2";

                // Assign CommandText and Open DB Connections
                comm.CommandText = sql;
                comm.Connection.Open();

                // Create Reader
                SqlDataReader sqlReader = comm.ExecuteReader();
                // While Reading from SQL - Write SQL data to Text File
                while (sqlReader.Read())
                {
                    //if (prtLamp1 != "")
                    //{
                    //    break;
                    //}

                    // Variables
                    string Key1 = sqlReader["Key1"].ToString();
                    string Key2 = sqlReader["Key2"].ToString();

                    // Lamp Code
                    foreach (string item2 in lampList)
                    {
                        // Check Number01 from ud100 to see the starting point in the string to find the lamp code // 01/29/2019 DTF

                        if (Key2 == item2)
                        {
                            nbrLamps = nbrLamps + 1;

                            if (lampCode1 == "")
                                lampCode1 = Key2;
                            else
                            if (lampCode2 == "")
                                lampCode2 = Key2;
                            else
                                lampCode3 = Key2;

                            /* MAIN SQL SELECT STATEMENT - JWM 04/09/2019 Commented out to split logic for use in under-counter labels */
                            //String sql3 = "SELECT ShortChar01 FROM Ice.UD01 as T1 WHERE T1.Key2 = '" + lampCode1 + "' AND Key1 = 'LampMst'";

                            //comm2.CommandText = sql3;
                            //comm2.Connection.Open();
                            //prtLamp1 = comm2.ExecuteScalar().ToString();
                            //comm2.Connection.Close();
                            GetLampInfo(Key2);

                            if (Key2 == lampCode1)
                                prtLamp1 = prtLamp;

                            if (Key2 == lampCode2 && wrkNbrLamps > 1)
                                prtLamp2 = prtLamp;

                            if (Key2 == lampCode3 && wrkNbrLamps > 2)
                                prtLamp3 = prtLamp;
                        }

                        ucbElementCounter++;
                    }
                }
                comm.Connection.Close();
                sqlReader.Close();
            }
        }

        public void GetLampInfo(string LampCode)
        {
            /* JWM 04/09/2019 This method gets the Lamp text to put on the label */
            // MAIN SQL SELECT STATEMENT
            String sql = $@"SELECT ShortChar01, ShortChar03 FROM " + eServer + $@".Ice.UD01 as T1 WHERE T1.Company = 'KEN' AND T1.Key1 = 'LampMst' AND T1.Key2 = '{LampCode}'";
            prtLamp = "";

            SqlCommand comm = new SqlCommand();
            comm.Connection = new SqlConnection(sqlConnString);
            comm.CommandText = sql;
            comm.Connection.Open();
            // Create Reader
            SqlDataReader sqlReader = comm.ExecuteReader();

            // While Reading from SQL - Write SQL data to Text File
            while (sqlReader.Read())
            {
                prtLamp = sqlReader["ShortChar01"].ToString();
                if (sqlReader["ShortChar03"].ToString() != "B" && sqlReader["ShortChar03"].ToString() != "C")
                   prtmsg12 = true;
            }

            comm.Connection.Close();
            sqlReader.Close();
        }

        public void GetWattsInfo()
        {
            wrkwatts = "";

            //SqlCommand comm2 = new SqlCommand();
            //comm2.Connection = new SqlConnection(sqlConnString);

            // MAIN SQL SELECT STATEMENT
            // String sql2 = "SELECT Number01 FROM Ice.UD100 as T1 WHERE T1.Company = 'KEN' AND T1.Key1 = '" + part1 + "'";

            //comm2.CommandText = sql2;
            //comm2.Connection.Open();
            //int startIndex = Convert.ToInt32(comm2.ExecuteScalar());
            //comm2.Connection.Close();

            //IEnumerable<string> lampList = elementList.Skip(startIndex - 1);

            // Create connection - State Database Connection, Main SQL SELECT
            SqlCommand comm = new SqlCommand();
            comm.Connection = new SqlConnection(sqlConnString);

            // MAIN SQL SELECT STATEMENT
            //            String sql = "SELECT * FROM Ice.UD01 as T1 WHERE T1.Key1 = 'LampMst'";
            String sql = $@"SELECT ShortChar02 FROM " + eServer + $@".Ice.UD01 as T1 WHERE T1.Company = 'KEN' and T1.Key1 = 'LampMst' AND T1.Key2 = '{lampCode1}'";

            // Assign CommandText and Open DB Connections
            comm.CommandText = sql;
            comm.Connection.Open();

            // Create Reader
            SqlDataReader sqlReader = comm.ExecuteReader();

            // While Reading from SQL - Write SQL data to Text File
            while (sqlReader.Read())
            {
                //// Get Volts
                //foreach (string item in lampList)
                //{
                //    string key2 = sqlReader["Key2"].ToString().Replace(" ", "");
                //    string elementItem = item.ToString();
                //    if (key2 == elementItem)
                //    {
                //        wrkwatts = sqlReader["ShortChar02"].ToString();
                //    }
                //    ucbElementCounter++;
                //}
                wrkwatts = sqlReader["ShortChar02"].ToString();
            }
            comm.Connection.Close();
            sqlReader.Close();
        }

        // Get Pick Code and msgID      /* 04/18/2019 JWM - Not used */
        //public void GetUlmsgInfo()
        //{

        //    // Create connection - State Database Connection, Main SQL SELECT
        //    SqlCommand comm = new SqlCommand();
        //    comm.Connection = new SqlConnection(sqlConnString);

        //    // MAIN SQL SELECT STATEMENT
        //    String sql = "SELECT Key1, ShortChar01, ShortChar02, ShortChar03 FROM Ice.UD100A WHERE Company = 'KEN' AND Key1 = '" + part1 + "'";

        //    // Assign CommandText and Open DB Connections
        //    comm.CommandText = sql;
        //    comm.Connection.Open();

        //    // Create Reader
        //    SqlDataReader sqlReader = comm.ExecuteReader();

        //    // While Reading from SQL - Write SQL data to Text File
        //    while (sqlReader.Read())
        //    {
        //        pickCode = "";
        //        msgID = "";
        //        // Get Volts
        //        foreach (string item in elementList)
        //        {
        //            string elementItem = item.ToString();
        //            if (sqlReader["ShortChar02"].ToString().StartsWith(elementItem))
        //            {
        //                labelTypeCheck = sqlReader["ShortChar02"].ToString();
        //                pickCode = sqlReader["ShortChar02"].ToString();
        //                msgID = sqlReader["ShortChar03"].ToString();
        //                continue;
        //            }
        //            ucbElementCounter++;
        //        }
        //    }
        //    comm.Connection.Close();
        //    sqlReader.Close();
        //}

        public void GetWorkMessages()
        {
            MessagesList.Clear();

            // Create connection - State Database Connection, Main SQL SELECT
            SqlCommand comm = new SqlCommand();
            comm.Connection = new SqlConnection(sqlConnString);

            // MAIN SQL SELECT STATEMENT
            String sql = "SELECT Key1, Key2, ShortChar01, ShortChar02, ShortChar03 FROM " + eServer + $@".Ice.UD100A WHERE Company = 'KEN' AND Key1 = '" + part1 + "'";

            // Assign CommandText and Open DB Connections
            comm.CommandText = sql;
            comm.Connection.Open();

            // Create Reader
            SqlDataReader sqlReader = comm.ExecuteReader();

            // Start stringbuilder
            //StringBuilder sb = new StringBuilder();
            //sb.AppendLine((char)(2) + "L"); // tells the printer to start executing the command
            //sb.AppendLine("Q0001");
            //sb.AppendLine("D11"); // Dot size
            //int msgCount = 0;
            // While Reading from SQL - Write SQL data to Text File
            while (sqlReader.Read())
            {
                // SEE if there is an Optional UL Msg ID for this, if there is use it instead of ShortChar01 field
                // string optionalID = sqlReader["ShortChar03"].ToString();
                pickCode = "";
                currentID = "";

                if (sqlReader["ShortChar02"].ToString() != "")
                {
                    if (partDesc.IndexOf(sqlReader["ShortChar02"].ToString(), StringComparison.CurrentCultureIgnoreCase) > 0)
                        pickCode = sqlReader["ShortChar02"].ToString();
                }

                if (pickCode != "")
                    currentID = sqlReader["ShortChar03"].ToString();
                else
                    currentID = sqlReader["ShortChar01"].ToString();

                // string currentID = sqlReader["ShortChar01"].ToString();
                if (currentID != "")
                {
                    SqlCommand comm2 = new SqlCommand();
                    comm2.Connection = new SqlConnection(sqlConnString);

                    // MAIN SQL SELECT STATEMENT
                    String sql2 = @"SELECT Key1,
                                           Key2,
                                           Character01,
                                           Character02,
                                           Character03,
                                           ShortChar03
                                      FROM " + eServer + $@".Ice.UD01
                                     WHERE Company = 'KEN'
                                       AND Key1 = 'ULMsgMst'
                                       AND Key2 = '" + currentID + "'";

                    // Assign CommandText and Open DB Connections
                    comm2.CommandText = sql2;
                    comm2.Connection.Open();
                    //if (msgCount == 0)
                    //    sb.AppendLine("121100001160024" + sqlReader["Character01"].ToString());
                    //if (msgCount == 1)
                    //    sb.AppendLine("121100001040024" + sqlReader["Character01"].ToString());
                    //if (msgCount == 2)
                    //    sb.AppendLine("121100000920024" + sqlReader["Character01"].ToString());
                    //if (msgCount == 3)
                    //    sb.AppendLine("121100000800024" + sqlReader["Character01"].ToString());
                    //if (msgCount == 4)
                    //    sb.AppendLine("121100000680024" + sqlReader["Character01"].ToString());
                    //if (msgCount == 5)
                    //    sb.AppendLine("121100000560024" + sqlReader["Character01"].ToString());
                    //if (msgCount == 6)
                    //    sb.AppendLine("121100000440024" + sqlReader["Character01"].ToString());
                    //if (msgCount == 7)
                    //    sb.AppendLine("121100000320024" + sqlReader["Character01"].ToString());
                    //msgCount++;
                    SqlDataReader sqlReader2 = comm2.ExecuteReader();

                    while (sqlReader2.Read())
                    {
                        MessagesList.Add(sqlReader2["Character01"].ToString());

                        if (sqlReader2["Character02"].ToString() != "")
                            MessagesList.Add(sqlReader2["Character02"].ToString());
                    }
                    // Close Connections
                    sqlReader2.Close();
                    comm2.Connection.Close();
                }
            }
            // Close Connections
            sqlReader.Close();
            comm.Connection.Close();

            if (prtmsg12 == true)
               AddMessage12();

            // End Print Function and Print Material Label If any Materials Available
            //if (msgCount > 0)
            //{
            //    sb.AppendLine("E");
            //    // PRINT MATERIAL RECORD LABEL
            //    RawPrinterHelper.SendStringToPrinter(LargeLabelPrint, sb.ToString());
            //}
        }

        public void AddMessage12()
        {
            // Create connection - State Database Connection, Main SQL SELECT
            SqlCommand comm = new SqlCommand();
            comm.Connection = new SqlConnection(sqlConnString);

            // MAIN SQL SELECT STATEMENT
            String sql = @"SELECT Character01,
                                  Character02
                             FROM " + eServer + $@".Ice.UD01
                            WHERE Company = 'KEN'
                              AND Key1 = 'ULMsgMst'
                              AND Key2 = '12'";

            // Assign CommandText and Open DB Connections
            comm.CommandText = sql;
            comm.Connection.Open();

            // Create Reader
            SqlDataReader sqlReader = comm.ExecuteReader();

            while (sqlReader.Read())
            {
                MessagesList.Add(sqlReader["Character01"].ToString());

                if (sqlReader["Character02"].ToString() != "")
                    MessagesList.Add(sqlReader["Character02"].ToString());
            }
            // Close Connections
            sqlReader.Close();
            comm.Connection.Close();
        }

        // GET IP CLASS AND LABEL SIZE
        public void GetIpClass()
        {
            // Create connection - State Database Connection, Main SQL SELECT
            SqlCommand comm = new SqlCommand();
            comm.Connection = new SqlConnection(sqlConnString);

            // MAIN SQL SELECT STATEMENT
            String sql = "SELECT ShortChar01 FROM " + eServer + $@".Ice.UD100 as T1 WHERE T1.Company = 'KEN' AND T1.Key1 = '" + part1 + "'";

            comm.CommandText = sql;
            comm.Connection.Open();
            ipClass = comm.ExecuteScalar().ToString();
            comm.Connection.Close();

            SqlCommand comm2 = new SqlCommand();
            comm2.Connection = new SqlConnection(sqlConnString);

            // MAIN SQL SELECT STATEMENT
            String sql2 = "SELECT ShortChar03 FROM " + eServer + $@".Ice.UD100 as T1 WHERE T1.Company = 'KEN' AND T1.Key1 = '" + part1 + "'";

            comm2.CommandText = sql2;
            comm2.Connection.Open();
            labelSize = comm2.ExecuteScalar().ToString();
            comm2.Connection.Close();
        }

        // Call Print Functions Based on Label Size
        public void PrintLabels()
        {
            // Print Large Labels
            if (labelSize == "1")
            {
                PrintLargeLabel();
            }
            // Print Small Lables
            if (labelSize == "2")
            {
                if (isDownlight == true)
                    PrintSmallLabel();
                else
                    PrintExitLabel();
            }
        }

        // Check if fixture is one of the STUD series fixtures
        public void CheckIfStudo()
        {
            isDownlight = false;
            isSTUDO = false;
            isSTUDX = false;
            isSTUDZ = false;
            studioText = "";
            //studoCode = "";
            string sql = "";

            // Create connection - State Database Connection, Main SQL SELECT
            SqlCommand comm = new SqlCommand();
            comm.Connection = new SqlConnection(sqlConnString);

            /* JWM 04/12/2019 Commented out - Logic incorrect */
            //SqlCommand comm2 = new SqlCommand();
            //comm2.Connection = new SqlConnection(sqlConnString);

            //// MAIN SQL SELECT STATEMENT
            //String sql = "SELECT TOP 1 ShortChar02 Company FROM Ice.UD100 WHERE Company = 'KEN' AND Key1 = '" + part1 + "' AND (Shortchar02 = 'STUDO' OR Shortchar02 = 'STUDX' OR Shortchar02 = 'STUDZ')";
            //String sql2 = "SELECT TOP 1 T3.Character02 FROM Ice.UD100 as T1 LEFT OUTER JOIN Ice.UD100A as T2 ON T2.Company = T1.Company LEFT OUTER JOIN Ice.UD01 as T3 ON T3.Company = T2.Company AND T3.Key1 = 'ULMsgMst' WHERE T1.Company = 'KEN' AND T1.Key1 = '" + part1 + "' AND T3.Key2 = T2.ShortChar01 AND T3.Character01 = '" + part2 + "'";

            //// If no Studo code for this part then CATCH out and close connection
            //try
            //{
            //    // Studo Code
            //    comm.CommandText = sql;
            //    comm.Connection.Open();
            //    studoCode = comm.ExecuteScalar().ToString();
            //    comm.Connection.Close();
            //}
            //catch { comm.Connection.Close(); }

            //// If no Studo Text for this part then CATCH out and close connection
            //try
            //{
            //    // StudoText
            //    comm2.CommandText = sql2;
            //    comm2.Connection.Open();
            //    studioText = comm2.ExecuteScalar().ToString();
            //    comm2.Connection.Close();
            //}
            //catch { comm2.Connection.Close(); }

            // Set sql to get the studo code
            sql = @"SELECT Key1,
                           ShortChar02
                      FROM " + eServer + $@".Ice.UD100
                     WHERE Company = 'KEN'
                       AND Key1 = '" + part1 +
                   @"' AND (Shortchar02 = 'STUDO' OR Shortchar02 = 'STUDX' OR Shortchar02 = 'STUDZ')";

            // If no Studo code for this part then CATCH out and close connection
            try
            {
                // Studo Code
                comm.CommandText = sql;
                comm.Connection.Open();

                // Create Reader
                SqlDataReader sqlReader = comm.ExecuteReader();

                while (sqlReader.Read())
                {
                    // Check if STUDO Lights
                    if (sqlReader["ShortChar02"].ToString() == "STUDO")
                    {
                        isDownlight = true;
                        isSTUDO = true;
                    }
                    // Check if STUDX Lights
                    if (sqlReader["ShortChar02"].ToString() == "STUDX")
                    {
                        isDownlight = true;
                        isSTUDX = true;
                    }
                    // Check if STUDZ Lights
                    if (sqlReader["ShortChar02"].ToString() == "STUDZ")
                    {
                        isDownlight = true;
                        isSTUDZ = true;
                    }
                }
                comm.Connection.Close();
            }
            catch { comm.Connection.Close(); }

            // studioText = "";
            // ucbFound = "N";

            // set sql to get the studo text message
            if (isSTUDO == true)
            {
                sql = @"SELECT TOP 1 UD01.Character02
                               FROM " + eServer + $@".Ice.UD100A as UD100A
                    LEFT OUTER JOIN " + eServer + $@".Ice.UD01 as UD01 ON UD01.Company = UD100A.Company AND UD01.Key1 = 'ULMsgMst' AND UD01.Key2 = UD100A.ShortChar01
                              WHERE UD100A.Company = 'KEN'
                                AND UD100A.Key1 = '" + part1 +
                            @"' AND UD01.Character01 = '" + part2 + "'";

                // If no Studo code for this part then CATCH out and close connection
                try
                {
                    // Studo Code
                    comm.CommandText = sql;
                    comm.Connection.Open();
                    studioText = comm.ExecuteScalar().ToString();
                    comm.Connection.Close();
                }
                catch { comm.Connection.Close(); }
            }
        }

        // Header Label Logic
        public void PrintHeaderLabel()
        {
            if (labelSize == "1")
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendLine((char)(2) + "L"); // tells the printer to start executing the command
                sb.AppendLine("Q0001");
                sb.AppendLine("D11"); // Dot size
                // sb.AppendLine("196600002000005" + "Job: " + jobValue.ToString());
                // sb.AppendLine("194400001600005" + "Due Date: " + Convert.ToDateTime(jobDueDate).ToString("MM/dd/yyyy"));
                sb.AppendLine("E");
                RawPrinterHelper.SendStringToPrinter(LargeLabelPrint, sb.ToString());
            }

            if (labelSize == "2")
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendLine((char)(2) + "L"); // tells the printer to start executing the command
                sb.AppendLine("Q0001");
                sb.AppendLine("D11"); // Dot size
                // sb.AppendLine("132100000700100" + "Job: " + jobValue.ToString());
                // sb.AppendLine("132100000500030" + "Due Date: " + Convert.ToDateTime(jobDueDate).ToString("MM/dd/yyyy"));
                sb.AppendLine("E");
                RawPrinterHelper.SendStringToPrinter(SmallLabelPrint, sb.ToString());
            }
        }

        // Material Label Logic
        public void PrintMaterialLabel()
        {
            // ** CREATE MATERIAL RECORD LABEL ** //
            int matCount = 0;
            int MtlLine = 215;

            FixtureLabelFound = "N";

            // Create connection - State Database Connection, Main SQL SELECT
            SqlCommand comm = new SqlCommand();
            comm.Connection = new SqlConnection(sqlConnString);

            // MAIN SQL SELECT STATEMENT
            String sql = $@"Select PartMtl.MtlPartNum as PartNum,
                                   Part.PartDescription as PartDescription,
	                               PartMtl.QtyPer as RequiredQty
                              FROM " + eServer + $@".Erp.PartMtl
                        INNER JOIN " + eServer + $@".dbo.Part as Part on Part.Company = PartMtl.Company AND Part.PartNum = PartMtl.MtlPartNum
                             WHERE PartMtl.Company = 'KEN'
                               AND PartMtl.PartNum = '{jobPartNum}'";

            // Assign CommandText and Open DB Connections
            comm.CommandText = sql;
            comm.Connection.Open();
            SqlDataReader sqlReader = comm.ExecuteReader();

            // Start stringbuilder
            StringBuilder sbMats = new StringBuilder();

            sbMats.AppendLine((char)(2) + "L"); // tells the printer to start executing the command
            sbMats.AppendLine("Q0001");
            sbMats.AppendLine("D11"); // Dot size
            sbMats.AppendLine("192200002300005" + "Part: " + jobPartNum.ToString());
            //materialFound = "N";

            // Loop through record returned in above sql statement
            while (sqlReader.Read())
            {
                string matPartNum = sqlReader["PartNum"].ToString();
                string matDesc = sqlReader["Description"].ToString();

                if (matDesc.Length > 40)    /* 04/24/2019 JWMJ - Limit the Part Description to first 40 characters */
                    matDesc = matDesc.Substring(0, 40); /* 04/24/2019 JWMJ - Limit the Part Description to first 40 characters */

                decimal matQtyReq = Convert.ToDecimal(sqlReader["RequiredQty"]);
                // Loop through all possible materials and add them to label
                if ((sqlReader["Description"].ToString().StartsWith("LBL") || sqlReader["Description"].ToString().StartsWith("INSTR") || sqlReader["Description"].ToString().StartsWith("CTN")) && matQtyReq > 0)
                {
                    // Check for Fixture Label in material list
                    if (matPartNum == "F-2200")
                        FixtureLabelFound = "Y";

                    //materialFound = "Y";
                    //if (matCount == 0)
                    //    sbMats.AppendLine("1922F0002150005" + matPartNum + "   " + matDesc + "    " + matQtyReq.ToString("N2"));
                    //if (matCount == 1)
                    //    sbMats.AppendLine("1922F0002000005" + matPartNum + "   " + matDesc + "    " + matQtyReq.ToString("N2"));
                    //if (matCount == 2)
                    //    sbMats.AppendLine("1922F0001850005" + matPartNum + "   " + matDesc + "    " + matQtyReq.ToString("N2"));
                    //if (matCount == 3)
                    //    sbMats.AppendLine("1922F0001700005" + matPartNum + "   " + matDesc + "    " + matQtyReq.ToString("N2"));
                    //if (matCount == 4)
                    //    sbMats.AppendLine("1922F0001550005" + matPartNum + "   " + matDesc + "    " + matQtyReq.ToString("N2"));
                    //if (matCount == 5)
                    //    sbMats.AppendLine("1922F0001400005" + matPartNum + "   " + matDesc + "    " + matQtyReq.ToString("N2"));
                    //if (matCount == 6)
                    //    sbMats.AppendLine("1922F0001250005" + matPartNum + "   " + matDesc + "    " + matQtyReq.ToString("N2"));
                    //if (matCount == 7)
                    //    sbMats.AppendLine("1922F0001100005" + matPartNum + "   " + matDesc + "    " + matQtyReq.ToString("N2"));
                    //if (matCount == 8)
                    //    sbMats.AppendLine("1922F0000970005" + matPartNum + "   " + matDesc + "    " + matQtyReq.ToString("N2"));
                    //if (matCount == 9)
                    //    sbMats.AppendLine("1922F0000820005" + matPartNum + "   " + matDesc + "    " + matQtyReq.ToString("N2"));
                    //if (matCount == 10)
                    //    sbMats.AppendLine("1922F0000680005" + matPartNum + "   " + matDesc + "    " + matQtyReq.ToString("N2"));
                    //if (matCount == 11)
                    //    sbMats.AppendLine("1922F0000550005" + matPartNum + "   " + matDesc + "    " + matQtyReq.ToString("N2"));
                    //if (matCount == 12)
                    //    sbMats.AppendLine("1922F0000400005" + matPartNum + "   " + matDesc + "    " + matQtyReq.ToString("N2"));
                    //if (matCount == 13)
                    //    sbMats.AppendLine("1922F0000270005" + matPartNum + "   " + matDesc + "    " + matQtyReq.ToString("N2"));
                    //if (matCount == 14)
                    //    sbMats.AppendLine("1922F0000130005" + matPartNum + "   " + matDesc + "    " + matQtyReq.ToString("N2"));

                    if (matCount == 14)
                    {
                        // Finish material label and send to printer
                        sbMats.AppendLine("E");
                        // PRINT MATERIAL RECORD LABEL
                        RawPrinterHelper.SendStringToPrinter(LargeLabelPrint, sbMats.ToString());

                        // Set up for a new material label
                        sbMats.Clear();
                        matCount = 0;
                        MtlLine = 215;

                        // Print the header for the next material label
                        sbMats.AppendLine((char)(2) + "L"); // tells the printer to start executing the command
                        sbMats.AppendLine("Q0001");
                        sbMats.AppendLine("D11"); // Dot size
                        sbMats.AppendLine("192200002300005" + "Part: " + jobPartNum.ToString());
                    }
                    sbMats.AppendLine($@"1922F00{MtlLine.ToString("0000")}0005{matPartNum}");
                    sbMats.AppendLine($@"1922F00{MtlLine.ToString("0000")}0070{matDesc}");
                    sbMats.AppendLine($@"1922F00{MtlLine.ToString("0000")}0355{matQtyReq.ToString("N2")}");

                    matCount++;
                    MtlLine -= 15;
                }
            }
            // Close Connections
            sqlReader.Close();
            comm.Connection.Close();
            // End Print Function and Print Material Label If any Materials Available
            if (matCount > 0)
            {
                sbMats.AppendLine("E");
                // PRINT MATERIAL RECORD LABEL
                RawPrinterHelper.SendStringToPrinter(LargeLabelPrint, sbMats.ToString());
            }
        }

        // Get Address Info
        public void GetAddressInfo()
        {
            wrkADDR1 = "";
            wrkADDR2 = "";
            wrkADDR3 = "";

            if (addressCode == "")
               addressCode = "Kenall";

            // Create connection - State Database Connection, Main SQL SELECT
            SqlCommand comm = new SqlCommand();
            comm.Connection = new SqlConnection(sqlConnString);

            // MAIN SQL SELECT STATEMENT
            //String sql = "SELECT ShortChar01 FROM Ice.UD01 WHERE Company = 'KEN' AND Key1 = 'ADDRMST' AND Key2 = '" + addressCode + "' AND Key3 = '" + labelSize + "'";
            String sql = "SELECT ShortChar01, ShortChar02, ShortChar03 FROM " + eServer + $@".Ice.UD01 WHERE Company = 'KEN' AND Key1 = 'ADDRMST' AND Key2 = '" + addressCode + "' AND Key3 = '" + labelSize + "'";

            comm.CommandText = sql;
            comm.Connection.Open();
            //wrkADDR = comm.ExecuteScalar().ToString();
            SqlDataReader sqlReader = comm.ExecuteReader();
            while (sqlReader.Read())
            {
                if (IndigoCleanPart == true)
                {
                    //wrkADDR1 = "Indigo Clean by Kenall Manufacturing";  /* Test */
                    wrkADDR1 = sqlReader["ShortChar01"].ToString();
                    wrkADDR2 = sqlReader["ShortChar02"].ToString();
                    wrkADDR3 = sqlReader["ShortChar03"].ToString();
                }
                else
                    wrkADDR1 = sqlReader["ShortChar01"].ToString();
            }
            sqlReader.Close();
            comm.Connection.Close();
        }

        // Print Last Label
        public void PrintLastLabel()
        {
            // Large Label
            if (labelSize == "1")
            {
                // Print last label on Large label
                StringBuilder sb = new StringBuilder();
                sb.AppendLine((char)(2) + "L"); // tells the printer to start executing the command
                sb.AppendLine("Q0001");
                sb.AppendLine("D11"); // Dot size
                sb.AppendLine("196600002000005" + "LAST LABEL");
                sb.AppendLine("E");
                RawPrinterHelper.SendStringToPrinter(LargeLabelPrint, sb.ToString());
            }
            // Small Label
            if (labelSize == "2")
            {
                // Print last label on small label
                StringBuilder sb = new StringBuilder();
                sb.AppendLine((char)(2) + "L"); // tells the printer to start executing the command
                sb.AppendLine("Q0001");
                sb.AppendLine("D11"); // Dot size
                sb.AppendLine("132100000700100" + "LAST LABEL");
                sb.AppendLine("E");
                RawPrinterHelper.SendStringToPrinter(SmallLabelPrint, sb.ToString());
            }
        }

        // Generate Error Labels
        public void CheckIfValidPart1()
        {
            // Check if Part1 is valid
            SqlCommand comm = new SqlCommand();
            comm.Connection = new SqlConnection(sqlConnString);

            // MAIN SQL SELECT STATEMENT
            String sql = "SELECT COUNT(*) From " + eServer + $@".Ice.UD100 WHERE Company = 'KEN' AND Key1 = '" + part1 + "'";

            comm.CommandText = sql;
            comm.Connection.Open();
            validRecords = comm.ExecuteScalar().ToString();
            comm.Connection.Close();

            if (validRecords == "0")
            {
                // If part1 is not a valid part id, then engineering needs to fix the description
                // Send Error Label
                StringBuilder sb = new StringBuilder();
                sb.AppendLine((char)(2) + "L"); // tells the printer to start executing the command
                sb.AppendLine("Q0001");
                sb.AppendLine("D11"); // Dot size
                sb.AppendLine("196600002000005" + "Part: " + jobPartNum.ToString());
                sb.AppendLine("194400001600005" + "Part Not Found in Element 1");
                sb.AppendLine("121100001390005" + "Fix data in element1 of Part Desc and Reprint");
                sb.AppendLine("121100001000005" + "STOP AND READ!");
                sb.AppendLine("E");
                RawPrinterHelper.SendStringToPrinter(LargeLabelPrint, sb.ToString());
            }
        }

        //public string ToSafeFileName(string s)
        //{
        //    return s
        //        .Replace("\\", "")
        //        .Replace("/", "")
        //        .Replace("\"", "")
        //        .Replace("*", "")
        //        .Replace(":", "")
        //        .Replace("?", "")
        //        .Replace("<", "")
        //        .Replace(">", "")
        //        .Replace("|", "");
        //}

        public void CheckIfF0862()
        {
            //string jobNum = jobValue;
            int recordCount = 0;
            // Check if job has a part material of F-0862
            SqlCommand comm = new SqlCommand();
            comm.Connection = new SqlConnection(sqlConnString);

            // MAIN SQL SELECT STATEMENT
            /* JWM 04/08/2019 Commented out due to being job oriented, put in statement below to find F-0862 part on approved revision for part */
            // String sql = "SELECT COUNT(*) From dbo.Part as T1 INNER JOIN EpicorERP.Erp.JobMtl as T2 ON T2.Company = T1.Company AND T2.PartNum = T1.PartNum AND AssemblySeq = 0 AND T2.PartNum = 'F-0862' WHERE T1.PartNum = '" + jobPartNum + "'";
            String sql = $@"select COUNT(*)
                              from " + eServer + $@".dbo.PartRev as PartRev
                        inner join " + eServer + $@".Erp.PartMtl as PartMtl ON PartMtl.Company = PartRev.Company and PartMtl.PartNum = PartRev.PartNum and PartMtl.RevisionNum = PartRev.RevisionNum
                             where PartRev.Company = 'KEN'
                               and PartRev.PartNum = '{jobPartNum}'
                               and PartRev.Approved = 1
                               and PartMtl.MtlPartNum = 'F-0862'";

            comm.CommandText = sql;
            comm.Connection.Open();
            recordCount = Convert.ToInt32(comm.ExecuteScalar().ToString());
            comm.Connection.Close();

            if (recordCount > 0)
            {
                hasF0862Material = true;

                //// Check if Part1 is valid
                //SqlCommand comm2 = new SqlCommand();
                //comm2.Connection = new SqlConnection(sqlConnString);

                //// MAIN SQL SELECT STATEMENT
                //String sql2 = "SELECT DueDate From dbo.JobHead where JobNum = '" + jobNum + "'";

                //comm2.CommandText = sql2;
                //comm2.Connection.Open();
                //F0826_DueDate = Convert.ToDateTime(comm2.ExecuteScalar().ToString()).ToString("MM/dd/yyyy");
                //comm2.Connection.Close();
            }
        }

        //public void GetExitInfo()
        //{
        //    string msg1, msg2, msg3, msg4, msg5, msg6;
        //}

        public void CheckForUL()
        {
            // LOOP THROUGH MATERIALS FOR JOB
            //string jobNum = jobValue;

            needsULLabel = false;
            //// Check if job has a part material of F-0862
            //SqlCommand comm = new SqlCommand();
            //comm.Connection = new SqlConnection(sqlConnString);

            //// MAIN SQL SELECT STATEMENT
            //String sql = "SELECT T2.JobNum, T2.PartNum, T3.Key1, T3.ShortChar01, T3.Shortchar02, T3.shortchar03 From dbo.JobHead as T1 INNER JOIN EpicorERP.Erp.JobMtl as T2 ON T2.Company = T1.Company AND T2.JobNum = T1.JobNum AND AssemblySeq = 0 INNER JOIN EpicorERP.Ice.UD100A as T3 ON T3.Company = T2.Company AND T3.ShortChar01 LIKE CONCAT(T2.PartNum, '%') WHERE T1.JobNum = '" + jobNum + "'";

            // Create connection - State Database Connection, Main SQL SELECT
            SqlCommand comm = new SqlCommand();
            comm.Connection = new SqlConnection(sqlConnString);

            // MAIN SQL SELECT STATEMENT
            String sql = $@"SELECT T1.PartNum as PartNum,
                                  T2.MtlPartNum as mtlPartNum,
                                  T3.Key1,
                                  T3.ShortChar01,
                                  T3.Shortchar02,
                                  T3.shortchar03
                             From " + eServer + $@".dbo.Part as T1
                       INNER JOIN " + eServer + $@".Erp.PartMtl as T2 ON T2.Company = T1.Company AND T2.PartNum = T1.PartNum
                       INNER JOIN " + eServer + $@".Ice.UD100A as T3 ON T3.Company = T2.Company AND T3.Key1 = '{part1}' AND T3.ShortChar01 LIKE CONCAT(T2.MtlPartNum, '%')
                            WHERE T1.Company = 'KEN'
							  and T1.PartNum = '{jobPartNum}'";

            // Assign CommandText and Open DB Connections
            comm.CommandText = sql;
            comm.Connection.Open();

            // Create Reader
            SqlDataReader sqlReader = comm.ExecuteReader();

            // While Reading from SQL - Write SQL data to Text File
            while (sqlReader.Read())
            {
                mtlPartNum = "";
                string shortChar02 = "";
                mtlPartNum = sqlReader["mtlPartNum"].ToString();
                shortChar02 = sqlReader["ShortChar02"].ToString();

                if (shortChar02 == "UL")
                {
                    needsULLabel = true;
                }
            }
            comm.Connection.Close();
        }





        private static void SelectDatabase(string ServerID)
        {
            /* This routine sets the connection string for the database based on what database the Epicor user is logged into */
            DBConnection = "";

            if (ServerID == FixtureLabelReprint_ByPart_WebApp.Properties.Settings.Default.TestServerID)
            {
                sqlConnString = FixtureLabelReprint_ByPart_WebApp.Properties.Settings.Default.DBConnectionTest;
                eServer = FixtureLabelReprint_ByPart_WebApp.Properties.Settings.Default.TestServerID;


            }
            else
                if (ServerID == FixtureLabelReprint_ByPart_WebApp.Properties.Settings.Default.PilotServerID)
            {
                sqlConnString = FixtureLabelReprint_ByPart_WebApp.Properties.Settings.Default.DBConnectionPilot;
                eServer = FixtureLabelReprint_ByPart_WebApp.Properties.Settings.Default.PilotServerID;

            }
            else
            {
                sqlConnString = FixtureLabelReprint_ByPart_WebApp.Properties.Settings.Default.DBConnectionLive;
                eServer = "EpicorErp";

            }
        }






    }
}