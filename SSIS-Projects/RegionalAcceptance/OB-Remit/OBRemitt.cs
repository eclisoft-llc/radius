using System;
using System.Data;
using System.Data.Odbc;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using CommonClasses;
using WinSCP;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Text;

namespace SMC_OBRemit
{
    public partial class Form1 : Form
    {
        public static string _delimiter = ",";
        public static string _environment = "THIRDPROD";
        public static string _clientPath = @"M:\RegionalAcceptance\";
        public static string _fileLocation = _clientPath + @"Outbound\";
        public static string _archivePath = _clientPath + @"Archive\";
        public static string _logPath = _clientPath + @"Logs\";
        public static string _appDirectory = @"R:\Program Files (x86)\WinSCP\";
        public string _approvalEmailsTo = "nnaemeka.okeke@radiusgs.com;pat.danner@radiusgs.com;LatitudeAccounting@radiusgs.com;Laura.Dailey@radiusgs.com;CRSArtivaAccounting@radiusgs.com;Beverly.WellensRudolph@radiusgs.com;clientservices@radiusgs.com;reports@radiusgs.com";//"nnaemeka.okeke@radiusgs.com;corisa.dcarlin@radiusgs.com;steve.palopoli@radiusgs.com";//;//
        public static List<string> CloseAccounts = new List<string>();

        public Form1()
        {
            InitializeComponent();
        }

        private void btnSubmit_Click(object sender, EventArgs e)
        {
            try
            {
                string statementNumber = txtStatementNumber.Text;
                string[] statementNumbers = null;
                string statementSQL = string.Empty;
                lblError.Text = string.Empty;

                var table = new DataTable();

                if (statementNumber.Contains(';'))
                    statementNumbers = statementNumber.Split(';');

                if (statementNumbers != null)
                {
                    statementSQL += " (";
                    int stmtCount = 0;

                    foreach (string stmtNum in statementNumbers)
                    {
                        if (stmtCount > 0)
                            statementSQL += " OR ";

                        statementSQL += "statementrecipienttable.AFSTRERUID = '" + stmtNum + @"'";
                        stmtCount++;
                    }

                    statementSQL += ") ";
                }
                else
                {
                    statementSQL += " statementrecipienttable.AFSTRERUID = '" + statementNumber + @"' ";
                    statementNumbers = new string[] { statementNumber };
                }

                string SQL = @"
SELECT DISTINCT fintransactiontable.AFTREFFDTE,
    account.ARACCLACCT,
   clientinfo.ARCLID,
   CASE WHEN fintranstype.AFTTCATEGORY = 'M' AND fintransactiontable.AFTRSTATUS='U' then 'AP'   
   ELSE fintransactiontable.AFTRTYP END AS TransactionCode,    
fintransactiontable.AFTRAMT,
   CASE WHEN fintranstype.AFTTCATEGORY = 'M' AND fintransactiontable.AFTRSTATUS='U' then 'Payment'   
   ELSE fintransactiontable.AFTRTYP END AS TransactionDescription,
account.ZZACBNKLNCODE,
account.ZZACBNKMIOCODE,
account.ZZACBATCHID,
coll.AFSCOVERAMT,
AFTTHOWRECEIVED,
AFTRREFERENCE
FROM
                                AFSPLIT splittable
                                INNER JOIN AFSTDETAIL splitdetailtable on splittable.AFSPKEY = splitdetailtable.AFSTDESPID
                                INNER JOIN AFSTRECIPIENT statementrecipienttable on splitdetailtable.AFSTDERECID = statementrecipienttable.AFSTREID
                                INNER JOIN AFSTRUN statementruntable on splitdetailtable.AFSTDERUID = statementruntable.AFSTRUID
                                INNER JOIN AFTRANSACTION fintransactiontable on splittable.AFSPTRNID = fintransactiontable.AFTRKEY
                                INNER JOIN AFTRANSTYPE fintranstype on fintransactiontable.AFTRTYP = fintranstype.AFTTKEY
                                INNER JOIN AFACCOUNT finaccount on splittable.AFSPACCTID  = finaccount.AFACKEY
                                INNER JOIN ARACCOUNT account on finaccount.AFACKEY = account.ARACFINACCTID
                                INNER JOIN ARCLIENT clientinfo on account.ARACCLTID = clientinfo.ARCLID
                                INNER JOIN ARENTITY entity on account.ARACRPID = entity.ARENID
                                INNER JOIN AFAPPLY ON AFAPPLY.AFAPSPLID = splittable.AFSPKEY
                                LEFT JOIN AFSCOLLECT coll ON splittable.AFSPCOLID = coll.AFSCKEY


WHERE fintranstype.AFTTHOWRECEIVED!='C' AND 
    statementruntable.AFSTRUSTATUS = 'Update Complete'  
AND" + statementSQL + @"
AND clientinfo.ARCLID IN ('RAC100','RAC505')
AND splitdetailtable.AFSTDETYID IS NOT NULL " +
"AND NOT (fintranstype.AFTTCATEGORY = 'R' AND fintransactiontable.AFTRSTATUS='U' AND fintranstype.ZZTTISMONEY='Y') ORDER BY clientinfo.ARCLID "; 


                string conString = "Dsn=" + _environment;

                using (OdbcConnection con = new OdbcConnection(conString))
                {
                    using (OdbcCommand cmd = new OdbcCommand(SQL, con))
                    {
                        try
                        {
                            con.Open();
                            OdbcDataAdapter da = new OdbcDataAdapter(cmd);
                            da.Fill(table);
                            da.Dispose();
                        }
                        catch (Exception ex)
                        {
                            lblError.ForeColor = System.Drawing.Color.Red;
                            lblError.Text = "An error occurred: " + ex.Message + (ex.InnerException != null ? ex.InnerException.Message : "");
                        }
                        finally
                        {
                            if (con.State == ConnectionState.Open)
                                con.Close();
                        }
                    }
                }

                if (!string.IsNullOrEmpty(lblError.Text))
                {
                    return;
                }

                string body = "";
                decimal totalCollected = 0.00m;
                InvoiceReturnRecord invRec = null;

                if (!(table == null || table.Columns.Count == 0))
                {
                    invRec = CreateInvoice(table);

                    if (!(invRec.EmailBody == "" && invRec.Log == ""))
                    {
                        body += invRec.EmailBody + ",\r\n,\r\n";
                        totalCollected += invRec.Total;
                    }

                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }



                body = @"Attached is the file for Regional Acceptance Corp's Payments and NSFs" + "<br>" +
                    @"Created By: " + Environment.UserName + "<br>" + " <br> " +
                    @"Total Collected: -$" + totalCollected.ToString("0.00").Replace("-","") + "<br>" +
                    @"For Statement(s): " + string.Join(",", statementNumbers) + "<br>" + " <br> " + body;


                string subject = "Regional Acceptance Corp Payments NSF Tracker - Total payments - -$" + totalCollected.ToString("0.00").Replace("-", "");
                DirectoryInfo di = new DirectoryInfo(_fileLocation);
                FileInfo[] fi = di.GetFiles("REGRMT." + DateTime.Today.ToString("MMdd"));
                FileInfo[] fi2 = di.GetFiles("CCSFIL." + DateTime.Today.ToString("MMdd"));

                //var sendfile = new FileInfo(_fileLocation + "RadiusPayments_" + DateTime.Today.ToString("MMddyyyy") + ".csv");
                //FileInfo[] fi2 = di.GetFiles("VVGMC_payments*.txt");
                //FileInfo[] fi = fi1.Concat(fi2).ToArray();

                if (fi.Length > 0)
                {
                    var attachments = new List<string>();
                    attachments.Add(fi[0].Directory + @"\" + fi[0].Name);
                    attachments.Add(fi2[0].Directory + @"\" + fi2[0].Name);
                    new EmailSender().SendMail(_approvalEmailsTo, "", attachments,
                        subject, body.Replace(",",""));

                    DialogResult dr = MessageBox.Show("Regional Acceptance Corp OB Payment Files created successfully\r\n\r\nAn email has been sent to client accounting.  Please wait until an approval email has been recieved before submitting the files to the client.\r\n\r\nIf approved, click OK to transfer files to Client.", "Files Created.", MessageBoxButtons.OKCancel);
                    if (dr == DialogResult.OK)
                    {
                        // OK given by Client Accounting.  Send the file to the client
                        /*string ftpURL = "";
                        string ftpUser = "";
                        string ftpPass = "";
                        int ftpPort = 0;
                        string ftpDir = "/Home/eft_RegionalAcceptance/From Agency/";
                        string ftpKey = "";

                        using (var connection = new SqlConnection("Server=dfw2-etl-001;Database=ArtivaJobEngine;User Id=WAREHOUSEMAN;Password=p3RV@$1V3;"))
                        {
                            using (var command = new SqlCommand("SELECT [Host],[Username],[Password],[Port],[HostKey]  FROM [ArtivaJobEngine].[dbo].[FTPInformation] WHERE ID = 14", connection))
                            {
                                connection.Open();
                                using (var reader = command.ExecuteReader())
                                {
                                    reader.Read();
                                    ftpURL = reader.GetString(0);
                                    ftpUser = reader.GetString(1);
                                    ftpPass = reader.GetString(2);
                                    ftpPort = reader.GetInt32(3);
                                    ftpKey = reader.GetString(4);
                                }
                            }
                        }

                        Session session = new Session();
                        TransferOptions transferOptions = new TransferOptions();
                        transferOptions.TransferMode = TransferMode.Binary;
                        transferOptions.ResumeSupport.State = TransferResumeSupportState.Off;
                        transferOptions.FilePermissions = null;
                        transferOptions.PreserveTimestamp = false;

                        SessionOptions sessionOptions = new SessionOptions();
                        sessionOptions.Protocol = Protocol.Sftp;
                        sessionOptions.HostName = ftpURL;
                        sessionOptions.UserName = ftpUser;
                        sessionOptions.Password = ftpPass;
                        sessionOptions.PortNumber = ftpPort;
                        sessionOptions.SshHostKeyFingerprint = ftpKey;

                        session.ExecutablePath = _appDirectory + "winscp.exe";
                        session.SessionLogPath = _logPath + "PutOBPaymentFiles.log";
                        session.Open(sessionOptions);
                        */
                        try
                        {
                            //MessageBox.Show(file.Directory + @"\" + file.Name);

                            File.Move(fi[0].FullName, _archivePath + fi[0].Name);
                            File.Move(fi2[0].FullName, _archivePath + fi2[0].Name);
                            //TransferOperationResult transResult = session.PutFiles(_archivePath + fi[0].Name, ftpDir, false, transferOptions);
                            //transResult.Check();

                            new EmailSender().SendMail("reports@radiusgs.com", "", null,
                            "Regional Acceptance Corp's Remit File Ready for Upload", "Regional Acceptance Corp's Remit File for today is ready to be uploaded. It is currently available in \"M:\\RegionalAcceptance\\Archive\" folder");


                            MessageBox.Show("Files have been transferred to M:\\RegionalAcceptance\\Archive!");
                        }
                        catch (Exception ex)
                        {
                            lblError.ForeColor = System.Drawing.Color.Red;
                            lblError.Text = "An error occurred: " + ex.Message + (ex.InnerException != null ? ex.InnerException.Message : "");
                        }

                        //session.Dispose();
                    }
                    else
                    {
                        MessageBox.Show("You have chosen to Cancel the transfer process.  All files that were created will now be deleted.");

                        try
                        {
                            fi[0].Delete();
                            fi2[0].Delete();
                            MessageBox.Show("The files have been deleted successfully.");
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("There was an issue removing the files.  Please contact IT for assistance.");
                        }
                    }
                }
                else
                    MessageBox.Show("No files created");
            }
            catch (Exception ex)
            {
                lblError.ForeColor = System.Drawing.Color.Red;
                lblError.Text = "An error occurred Hey!! : " + ex.Message + (ex.InnerException != null ? ex.InnerException.Message : "");
            }
        }

        private InvoiceReturnRecord CreateInvoice(DataTable tbl)
        {
            InvoiceReturnRecord retRec = new InvoiceReturnRecord();
            DateTime fileDate = DateTime.Today;
            string exportDirectory = _fileLocation;
            string fileName1 = "REGRMT." + fileDate.ToString("MMdd");
            string fileName2 = "CCSFIL." + fileDate.ToString("MMdd");
            //string fileName2 = "VVGMC_payments " + fileDate.ToString("MMdd") + ".txt";
            string fullFilePath1 = exportDirectory + fileName1;
            string fullFilePath2 = exportDirectory + fileName2;
            //string fullFilePath2 = exportDirectory + fileName2;
            //bool isfilename1 = false;
            //bool isfilename2 = false;
            decimal totalAmount1 = 0.00m;
            string body = "";
            string returnLog = "";
            lblError.Text = "";
            var regrmt = new StringBuilder();
            var ccsfil = new StringBuilder();
            var dt = DateTime.Today;


            string batchid = "";
            string currdate = dt.ToString("yyyyMMdd");
            string currtime = dt.ToString("HH:mm");
            string accid = "";
            string miocode = "";
            string amt = "";

            if (File.Exists(fullFilePath1))
                File.Delete(fullFilePath1);
            if (File.Exists(fullFilePath2))
                File.Delete(fullFilePath2);

            int tblcnt = tbl.Rows.Count;
            try
            {

                
                if (!(tbl == null || tblcnt == 0))
                {
                    try
                    {
                        string c = "";
                        string h = "";
                        string j = "";
                        foreach (DataRow row in tbl.Rows)
                        {
                            j = row["ARCLID"].ToString();                      
                            batchid = row["ZZACBATCHID"].ToString();
                            miocode = "5030";// row["ZZACBNKMIOCODE"].ToString();
                            accid = row["ARACCLACCT"].ToString();


                            if (c != j)
                            {
                                regrmt.AppendLine((currdate.PadRight(32) + "HD" + "000000000{" + "000000000{" + tbl.Select("ARCLID = '" + j + "'").Length.ToString().PadLeft(5,'0') + batchid + miocode).PadRight(100));
                                c = j;
                            }


                            //double overpayment = Math.Abs(double.Parse(CheckForNullField(row["AFSCOVERAMT"].ToString(), "", "0")));
                            //double payment = Math.Abs(double.Parse(CheckForNullField(row["AFTRAMT"].ToString(), "", "0")));
                            decimal overpayment = Math.Abs((decimal)row["AFSCOVERAMT"]);
                            decimal payment = Math.Abs((decimal)row["AFTRAMT"]);

                            if (row["AFTTHOWRECEIVED"].ToString() != "C")
                            {

                                payment -= overpayment;
                            }


                            totalAmount1 += payment;
                            amt = getNum(payment.ToString(), row["TransactionDescription"].ToString());
                            
                            regrmt.AppendLine((Convert.ToDateTime(row["AFTREFFDTE"].ToString()).ToString("yyyyMMdd") + "0000" + accid.PadRight(20) + row["TransactionCode"].ToString() + amt + 
                                "N".PadRight(2) + row["TransactionDescription"].ToString().PadRight(20) + amt.PadLeft(10,'0') + "0000" + "X" +
                       batchid + batchid.PadRight(8) + miocode).PadRight(100));

                            ccsfil.AppendLine("D" + accid.PadLeft(10) + getTotalNum(payment.ToString()).PadLeft(10, '0') + "0000000000");




                        }

                        File.AppendAllText(fullFilePath1, regrmt.ToString());
                        ccsfil.AppendLine("T" + tblcnt.ToString().PadLeft(10,'0') + getTotalNum(totalAmount1.ToString().Replace("-","")).PadLeft(10, '0'));
                        File.AppendAllText(fullFilePath2, ccsfil.ToString());

                        int rac100cnt = tbl.Select("ARCLID = 'RAC100'").Length;
                        int rac505cnt = tbl.Select("ARCLID = 'RAC505'").Length;

                        if (rac100cnt == 0)
                        {
                            regrmt = new StringBuilder();
                            regrmt.AppendLine((currdate.PadRight(32) + "HD" + "000000000{" + "000000000{" + "00000" + "C100" + "5030").PadRight(100));
                            File.AppendAllText(fullFilePath1, regrmt.ToString());
                        }
                        if (rac505cnt == 0)
                        {
                            regrmt = new StringBuilder();
                            regrmt.AppendLine((currdate.PadRight(32) + "HD" + "000000000{" + "000000000{" + "00000" + "C505" + "5030").PadRight(100));
                            File.AppendAllText(fullFilePath1, regrmt.ToString());
                        }

                        }
                    catch (Exception ex)
                    {

                        lblError.ForeColor = System.Drawing.Color.Red;
                        lblError.Text = "An error occurred: " + ex.Message + "\n\n" + ex.StackTrace + (ex.InnerException != null ? ex.InnerException.Message : "");
                    }
                }

                try
                {

                        returnLog += fileName1 + " Created Successfully!\r\n";
                        body +=
                        @"Target File: " + fileName1 + "<br>" +
                        @"Created Time: " + fileDate.ToString("MM/dd/yyyy HH:mm:ss") + "<br>" +
                        @"Amount Reported: -$" + totalAmount1.ToString("0.00").Replace("-", "") + "<br>" + "<br>";

                    


                }
                catch (Exception ex)
                {
                    lblError.ForeColor = System.Drawing.Color.Red;
                    lblError.Text = "An error occurred: " + ex.Message + (ex.InnerException != null ? ex.InnerException.Message : "");
                }
            }
            catch (Exception ex)
            {
                lblError.ForeColor = System.Drawing.Color.Red;
                lblError.Text = "An error occurred: " + ex.Message + (ex.InnerException != null ? ex.InnerException.Message : "");
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            retRec.Log = returnLog;
            retRec.EmailBody = body;
            retRec.Total += totalAmount1;
            return retRec;
        }
        private string getNum(string v, string typ)
        {

                
            bool iszero = Convert.ToDecimal(v)==0;            
            string strv = v.Replace("-", "").Replace(".","");

            var ibm = new Dictionary<string, string>() {
            { "0", "{" },{ "-0", "}" },
            { "1", "A" },{ "-1", "J" },
            { "2", "B" },{ "-2", "K" },
            { "3", "C" },{ "-3", "L" },
            { "4", "D" },{ "-4", "M" },
            { "5", "E" },{ "-5", "N" },
            { "6", "F" },{ "-6", "O" },
            { "7", "G" },{ "-7", "P" },
            { "8", "H" },{ "-8", "Q" },
            { "9", "I" },{ "-9", "R" }};

            string pad = strv.PadLeft(10, '0');           

            if(typ == "Reversal")
            {
                if (iszero)
                    return "000000000}";
                else
                    return pad.Remove(pad.Length - 1, 1) + ibm["-"+pad[pad.Length - 1].ToString()];
            }
            else
            {
                if (iszero)
                    return "000000000{";
                else
                    return pad.Remove(pad.Length - 1, 1) + ibm[pad[pad.Length - 1].ToString()];
            }




        }

        private string getTotalNum(string v)
        {
            
            bool iszero = Convert.ToDecimal(v) == 0;
            string strv = v.Replace("-", "").Replace(".", "");

            var ibm = new Dictionary<string, string>() {
            { "0", "{" },{ "-0", "}" },
            { "1", "A" },{ "-1", "J" },
            { "2", "B" },{ "-2", "K" },
            { "3", "C" },{ "-3", "L" },
            { "4", "D" },{ "-4", "M" },
            { "5", "E" },{ "-5", "N" },
            { "6", "F" },{ "-6", "O" },
            { "7", "G" },{ "-7", "P" },
            { "8", "H" },{ "-8", "Q" },
            { "9", "I" },{ "-9", "R" }};

            string pad = strv.PadLeft(10, '0');


            if (iszero)
                return "0000000000";// {";
            else
                return pad;//.Remove(pad.Length - 1, 1) + ibm[pad[pad.Length - 1].ToString()];
            




        }
        private string GetCommissionRate(string artivaCode)
        {
            if (artivaCode == "B020PCT") return "20";
            else if (artivaCode == "B025PCT") return "25";
            else if (artivaCode == "B030PCT") return "30";
            else if (artivaCode == "B040PCT") return "40";
            else if (artivaCode == "B045PCT") return "45";
            else if (artivaCode == "B050PCT") return "50";
            else if (artivaCode == "B055PCT") return "55";
            else if (artivaCode == "B060PCT") return "60";
            else return "";
        }

        private string GetTransactionCode(string artivaCode)
        {
            if ("CA, CC, CK, CCC, ECC, ECK, VI, MC, DSC, WAC, WU, MG, MO".Contains(artivaCode))
                return "PMT";
            else if (artivaCode == "NSF")
                return "RET";
            else if ("COR|DBJ".Contains(artivaCode))
                return "ERR";
            else return "";
        }

        private string GetPaymentID(string artivaCode)
        {
            artivaCode = artivaCode.ToUpper();

            if (artivaCode == "CHECK")
                return "1";
            else if (artivaCode == "MONEY ORDER")
                return "2";
            else if ("VISA, MASTER CARD,CREDIT CARD, DISCOVER".Contains(artivaCode))
                return "3";
            else if ("MONEY GRAM, WESTERN UNION".Contains(artivaCode))
                return "4";
            else if ("ELECTORONIC CHECK, ELECTRONIC CREDIT CARD".Contains(artivaCode))
                return "5";
            else return "";
        }

        private string GetStatus(string cancelID, string prevCancelCode, double curBalDbl)
        {
            string status = "";

            if (cancelID == "SETTLE")
                status = "Settlement";
            else if ((cancelID == "" || cancelID == "PIF") && curBalDbl == 0.00)
                status = "Paid in full";
            else if (cancelID == "RETURN")
            {
                if (prevCancelCode == "SETTLE")
                    status = "Settlement";
                else if ((prevCancelCode == "" || prevCancelCode == "PIF") && curBalDbl == 0.00)
                    status = "Paid in full";
            }

            return status;
        }

        private string CheckForNullField(object field, string format, string defaultIfNull = "")
        {
            string returnField = "";

            if (field != DBNull.Value)
            {
                if (format == "")
                    returnField = field.ToString().Trim();
                else
                    returnField = Convert.ToDateTime(field).ToString(format);
            }
            else
                returnField = defaultIfNull;

            return returnField;
        }
    }

    public class InvoiceReturnRecord
    {
        public string Log = "";
        public string EmailBody = "";
        public decimal Total = 0.00m;
        public bool isfilename1 = false;
        public bool isfilename2 = false;
    }
}