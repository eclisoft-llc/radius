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

namespace SMC_OBRemit
{
    public partial class Form1 : Form
    {
        public static string _delimiter = ",";
        public static string _environment = "THIRDPROD";
        public static string _clientPath = @"M:\integracredit\";
        public static string _fileLocation = _clientPath + @"Outbound\";
        public static string _archivePath = _clientPath + @"Archive\";
        public static string _logPath = _clientPath + @"Logs\";
        public static string _appDirectory = @"R:\Program Files (x86)\WinSCP\";
        public string _approvalEmailsTo = "nnaemeka.okeke@radiusgs.com;corisa.dcarlin@radiusgs.com;CRSArtivaAccounting@radiusgs.com;Beverly.WellensRudolph@radiusgs.com;clientservices@radiusgs.com;reports@radiusgs.com";//"nnaemeka.okeke@radiusgs.com;corisa.dcarlin@radiusgs.com;steve.palopoli@radiusgs.com";//;//
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

                string SQL = @"SELECT Account#,
SecondaryAccount#,
TransactionTypeCode,     
TenderDate,
TransactionReference,
PaymentAmount,   
PrincipalPmtAmount,
InterestPmtAmount,
OtherPmtAmount,
AdditionalChargesPmtAmount,
PaymentStatuscode,
SUM(Commission) AS Commission
 FROM(
SELECT  
    account.ARACCLACCT AS Account#,
    '' AS SecondaryAccount#,
   CASE WHEN fintransactiontable.AFTRTYP NOT IN ('CA','CC','CCC','CK','CRJ','DBJ','ECC','ECK','MC','MG','MO','NSF','CW','VI','ACH','OTH','COR')
   THEN case when fintranstype.AFTTCATEGORY = 'M' AND fintranstype.AFTTHOWRECEIVED!='C' then 'OTH'
   when fintranstype.AFTTCATEGORY = 'M' AND fintranstype.AFTTHOWRECEIVED='C' then ''
   when fintranstype.AFTTCATEGORY = 'R' AND fintranstype.AFTTHOWRECEIVED='C' then ''
   when fintranstype.AFTTCATEGORY = 'R' AND fintranstype.AFTTHOWRECEIVED!='C' then 'COR'
   when fintranstype.AFTTCATEGORY = 'A' AND fintranstype.AFTTINCDEC='I' then 'DBJ'
   when fintranstype.AFTTCATEGORY = 'A' AND fintranstype.AFTTINCDEC='D' then 'CRJ'
   end WHEN fintransactiontable.AFTRTYP='CW' THEN 'ECK' ELSE fintransactiontable.AFTRTYP END AS TransactionTypeCode,     
fintransactiontable.AFTRACCTDTE AS TenderDate,
fintransactiontable.AFTRREFERENCE AS TransactionReference,
fintransactiontable.AFTRAMT AS PaymentAmount,   
'' AS PrincipalPmtAmount,
'' AS InterestPmtAmount,
'' AS OtherPmtAmount,
'' AS AdditionalChargesPmtAmount,
CASE WHEN account.ARACCANCID='SETTLE' THEN 'SIF' WHEN (account.ARACCANCID='' OR account.ARACCANCID IS NULL) AND finaccount.AFACCURBALINT<=0 THEN 'PIF'  ELSE '' END AS PaymentStatuscode,
AFAFEE.AFAFDUEAGY AS Commission
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
                                LEFT JOIN AFAFEE ON AFAFEE.afafkey = AFAPPLY.afapfeeid

WHERE
    statementruntable.AFSTRUSTATUS = 'Update Complete'  
AND" + statementSQL + @"
AND clientinfo.ARCLID IN ('INTEG1','INTEG2','INTEG3','INTEG4','INTEG5')
AND splitdetailtable.AFSTDETYID IS NOT NULL) AS G
GROUP BY Account#,
SecondaryAccount#,
TransactionTypeCode,     
TenderDate,
TransactionReference,
PaymentAmount,   
PrincipalPmtAmount,
InterestPmtAmount,
OtherPmtAmount,
AdditionalChargesPmtAmount,
PaymentStatuscode";


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
                double totalCollected = 0.00;
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



                body = @"Attached is the file for Integra Credit's Payments and NSFs" + "<br>" +
                    @"Created By: " + Environment.UserName + "<br>" + " <br> " +
                    @"Total Collected: -$" + totalCollected.ToString("0.00").Replace("-","") + "<br>" +
                    @"For Statement(s): " + string.Join(",", statementNumbers) + "<br>" + " <br> " + body;


                string subject = "Integra Credit Payments NSF Tracker - Total payments - -$" + totalCollected.ToString("0.00").Replace("-", "");
                DirectoryInfo di = new DirectoryInfo(_fileLocation);
                FileInfo[] fi = di.GetFiles("RadiusPayments_" + DateTime.Today.ToString("MMddyyyy") + "*.csv");
                
                //var sendfile = new FileInfo(_fileLocation + "RadiusPayments_" + DateTime.Today.ToString("MMddyyyy") + ".csv");
                //FileInfo[] fi2 = di.GetFiles("VVGMC_payments*.txt");
                //FileInfo[] fi = fi1.Concat(fi2).ToArray();
                
                if (fi.Length > 0)
                {
                    var attachments = new List<string>();
                    attachments.Add(fi[0].Directory + @"\" + fi[0].Name);
                    new EmailSender().SendMail(_approvalEmailsTo, "", attachments,
                        subject, body.Replace(",",""));

                    DialogResult dr = MessageBox.Show("Integra Credit OB Payment Files created successfully\r\n\r\nAn email has been sent to client accounting.  Please wait until an approval email has been recieved before submitting the files to the client.\r\n\r\nIf approved, click OK to transfer files to Client.", "Files Created.", MessageBoxButtons.OKCancel);
                    if (dr == DialogResult.OK)
                    {
                        // OK given by Client Accounting.  Send the file to the client
                        string ftpURL = "";
                        string ftpUser = "";
                        string ftpPass = "";
                        int ftpPort = 0;
                        string ftpDir = "/Home/eft_integracredit/From Agency/";
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

                        try
                        {
                            //MessageBox.Show(file.Directory + @"\" + file.Name);

                            File.Move(fi[0].FullName, _archivePath + fi[0].Name);
                            TransferOperationResult transResult = session.PutFiles(_archivePath + fi[0].Name, ftpDir, false, transferOptions);
                            transResult.Check();


                            MessageBox.Show("Files have been transferred to the client successfully!");
                        }
                        catch (Exception ex)
                        {
                            lblError.ForeColor = System.Drawing.Color.Red;
                            lblError.Text = "An error occurred: " + ex.Message + (ex.InnerException != null ? ex.InnerException.Message : "");
                        }

                        session.Dispose();
                    }
                    else
                    {
                        MessageBox.Show("You have chosen to Cancel the transfer process.  All files that were created will now be deleted.");

                        try
                        {
                            fi[0].Delete();
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
            string fileName1 = "RadiusPayments_" + fileDate.ToString("MMddyyyy") + ".csv";
            //string fileName2 = "VVGMC_payments " + fileDate.ToString("MMdd") + ".txt";
            string fullFilePath1 = exportDirectory + fileName1;
            //string fullFilePath2 = exportDirectory + fileName2;
            //bool isfilename1 = false;
            //bool isfilename2 = false;
            double totalAmount1 = 0.00;
            //double totalAmount2 = 0.00;
            string body = "";
            string returnLog = "";
            lblError.Text = "";

            if (File.Exists(fullFilePath1))
                File.Delete(fullFilePath1);

            try
            {


                if (!(tbl == null || tbl.Columns.Count == 0))
                {
                    try
                    {
                        foreach (DataRow row in tbl.Rows)
                        {


                            totalAmount1 += double.Parse(CheckForNullField(row["PaymentAmount"], "", "0"));

                            string[] transactionref = CheckForNullField(row["TransactionReference"], "").Split(' ');
                            string tranref = "";

                            if (transactionref.Length > 1)
                                tranref = transactionref[transactionref.Length - 1];
                            else
                                tranref = CheckForNullField(row["TransactionReference"], "");

                            File.AppendAllText(fullFilePath1,
                                CheckForNullField(row["Account#"], "") + "," +
                                CheckForNullField(row["SecondaryAccount#"], "") + "," +
                                CheckForNullField(row["TransactionTypeCode"], "") + "," +  
                                "" + "," +                                                         
                                CheckForNullField(row["TenderDate"], "MMddyyyy") + "," +
                                tranref + "," +
                                (0 + double.Parse(CheckForNullField(row["PaymentAmount"], ""))).ToString("0.00").Replace("-", "") + "," +
                                CheckForNullField(row["PrincipalPmtAmount"], "") + "," +
                                CheckForNullField(row["InterestPmtAmount"], "") + "," +                                
                                CheckForNullField(row["OtherPmtAmount"], "") + "," +
                                CheckForNullField(row["AdditionalChargesPmtAmount"], "") + "," +                                
                                CheckForNullField(row["PaymentStatuscode"], "") + "," +
                                (double.Parse(CheckForNullField(row["Commission"], ""))).ToString("0.00").Replace("-", "") + "\r\n");
                            
                            
                           




                        }
                    }
                    catch (Exception ex)
                    {

                        lblError.ForeColor = System.Drawing.Color.Red;
                        lblError.Text = "An error occurred: " + ex.Message + (ex.InnerException != null ? ex.InnerException.Message : "");
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
        public double Total = 0.00;
        public bool isfilename1 = false;
        public bool isfilename2 = false;
    }
}