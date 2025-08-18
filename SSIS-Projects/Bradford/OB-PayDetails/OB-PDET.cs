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
using System.Diagnostics;
using System.IO.Compression;

namespace SMC_OBRemit
{
    public partial class Form1 : Form
    {
        public static string _delimiter = ",";
        public static string _environment = "THIRDPROD";
        public static string _clientPath = @"M:\BradfordGroup\";
        public static string _fileLocation = _clientPath + @"Outbound\";
        public static string _archivePath = _clientPath + @"Archive\";
        public static string _logPath = _clientPath + @"Logs\";
        public static string _appDirectory = @"R:\Program Files (x86)\WinSCP\";
        public string _approvalEmailsTo = "nnaemeka.okeke@radiusgs.com;corisa.dcarlin@radiusgs.com;CRSArtivaAccounting@radiusgs.com;Beverly.WellensRudolph@radiusgs.com;ClientServices-All@radiusgs.com;reports@radiusgs.com;processing@radiusgs.com;Travis.Lane@Radiusgs.com;jodi.berman@radiusgs.com;TJ.Apfel@radiusgs.com;donald.schnabel@radiusgs.com";
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

                createReport(statementSQL);


                if (!string.IsNullOrEmpty(lblError.Text))
                {
                    return;
                }

                string body = "";

                body = @"Attached is the file for Bradford's Payment Detail Report" + "<br>" +
                    @"Created By: " + Environment.UserName + "<br>" + " <br> " +
                    @"For Statement(s): " + string.Join(",", statementNumbers) + "<br>" + " <br> " + body;

                string subject = "Attached is the file for Bradford's Payment Detail Report";
                DirectoryInfo di = new DirectoryInfo(_fileLocation);
                FileInfo[] fi = di.GetFiles("BradfordPortfolioObReport_" + DateTime.Today.ToString("MMddyyyy") + "*.xlsx");


                if (fi.Length > 0)
                {
                    var attachments = new List<string>();
                    attachments.Add(fi[0].Directory + @"\" + fi[0].Name);
                    new EmailSender().SendMail(_approvalEmailsTo, "", attachments,
                        subject, body.Replace(",", ""));

                    DialogResult dr = MessageBox.Show("Bradford OB Payment Report created successfully\r\n\r\nAn email has been sent to client accounting.  Click OK to exit window.", "Files Created.", MessageBoxButtons.OKCancel);
                    File.Move(fi[0].FullName, _archivePath + fi[0].Name);
                    /*
                    if (dr == DialogResult.OK)
                    {
                        // OK given by Client Accounting.  Send the file to the client
                        string ftpURL = "";
                        string ftpUser = "";
                        string ftpPass = "";
                        int ftpPort = 0;
                        string ftpDir = "/PRDP/0314/";
                        string ftpKey = "";

                        using (var connection = new SqlConnection("Server=dfw2-etl-001;Database=ArtivaJobEngine;User Id=WAREHOUSEMAN;Password=p3RV@$1V3;"))
                        {
                            using (var command = new SqlCommand("SELECT [Host],[Username],[Password],[Port],[HostKey]  FROM [ArtivaJobEngine].[dbo].[FTPInformation] WHERE ID = 49", connection))
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
                            //File.Move(fi[0].FullName.Replace(".gz",""), _archivePath + fi[0].Name.Replace(".gz", ""));
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

                    */
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

        object missing = Type.Missing;
        private string[] letterphase = { "BBEL01", "BAGL03", "BASL07", "BHCL09", "BHAL10", "BVSL11", "BBPL12", "BHAL14", "BBEL15", "BCNL16", "BBAL17", "BBCL18", "BSHL19", "BBOL40", "BAOL41", "BHOL42" };
        private string[] collectionphase = { "BBEC01", "BAGC03", "BASC07", "BHCC09", "BHAC10", "BVSC11", "BBPC12", "BHAC14", "BBEC15", "BCNC16", "BBAC17", "BBCC18", "BSHC19", "BBOC40", "BAOC41", "BHOC42" };

        private void createReport(string statementSQL)
        {

            var acc = new DataTable();
            var pmt = new DataTable();
            var lmtd = new DataTable();
            var bal = new DataTable();
            var plc = new DataTable();
            var ABlcnt = new DataTable();

            HashSet<string> lmtdData = new HashSet<string>();
            HashSet<string> allrooms = new HashSet<string>();
            HashSet<string> lstmthrooms = new HashSet<string>();

            string[] s = DateTime.Today.AddMonths(-1).ToString("MM-yyyy").Split('-');
            string[] t = DateTime.Today.ToString("MM-yyyy").Split('-');

            string start = s[1] + "-" + s[0] + "-06";
            string end = t[1] + "-" + t[0] + "-05";

            string accQ = string.Format(@"SELECT DISTINCT zz.ZZRCRM, ent.ARENLNM, ent.ARENFNM,zz.ZZRCCUSTID, account.ARACCLACCT,account.ARACLSTDTE,  
                            arl.ARLHUPDDTE,arl.ARLHLTR,ent.ARENST 
,ARLHREQDTE
,ARLHPL95
FROM %STARTTABLE ARCLIENT INNER JOIN ARACCOUNT account on ARCLID = account.ARACCLTID AND ARCLID IN  
                             ('BBEL01', 'BAGL03', 'BASL07', 'BHCL09', 'BHAL10', 'BVSL11','BBPL12', 'BHAL14', 'BBEL15', 'BCNL16', 'BBAL17', 'BBCL18','BSHL19', 'BBOL40', 'BAOL41', 'BHOL42')  
                             JOIN ARENTITY ent ON account.ARACRPID = ent.ARENID  
                             JOIN ZZRETAILCLUB zz ON zz.ZZRCACID = account.ARACID  
                             LEFT JOIN ARLTRHIS arl ON arl.ARLHACID = account.ARACID  
            WHERE arl.ARLHREQDTE >= CAST('{0}' AS DATE) AND arl.ARLHREQDTE <= CAST('{1}' AS DATE) ORDER BY arl.ARLHREQDTE", start, end);


            string plcQ = string.Format(@"SELECT ZZRCRM,CONVERT(VARCHAR(3),ARAHLSTDTE,100) AS M,RIGHT(DATEPART(YEAR,ARAHLSTDTE),2) AS Y,COUNT(*) AS CNT,
			SUM(ARAHINILSTAMT) AS BAL,ARAHLSTDTE
FROM(SELECT ZZRCRM,ARACCTRH.ARAHLSTDTE,ARACCTRH.ARAHINILSTAMT	FROM %STARTTABLE ARCLIENT INNER JOIN ARACCTRH on ARCLID = ARACCTRH.ARAHCLTID AND ARCLID IN  
                             ('BBEC01','BAGC03','BASC07','BHCC09','BHAC10','BVSC11','BBPC12','BHAC14','BBEC15','BCNC16','BBAC17','BBCC18','BSHC19','BBOC40','BAOC41','BHOC42',
							 'BBEL01', 'BAGL03', 'BASL07', 'BHCL09', 'BHAL10', 'BVSL11','BBPL12', 'BHAL14', 'BBEL15', 'BCNL16', 'BBAL17', 'BBCL18','BSHL19', 'BBOL40', 'BAOL41', 'BHOL42')
INNER JOIN ZZRETAILCLUB zz ON zz.ZZRCACID = ARACCTRH.ARAHACID
UNION ALL 
SELECT ZZRCRM,ARACLSTDTE,ZZACPRINATPLC	FROM %STARTTABLE ARCLIENT INNER JOIN ARACCOUNT account on ARCLID = account.ARACCLTID AND ARCLID IN  
                             ('BBEL01', 'BAGL03', 'BASL07', 'BHCL09', 'BHAL10', 'BVSL11','BBPL12', 'BHAL14', 'BBEL15', 'BCNL16', 'BBAL17', 'BBCL18','BSHL19', 'BBOL40', 'BAOL41', 'BHOL42')  
                             JOIN ARENTITY ent ON account.ARACRPID = ent.ARENID  
                             LEFT JOIN ZZRETAILCLUB zz ON zz.ZZRCACID = account.ARACID  
			  WHERE ARACLSTDTE <= CAST('{0}' AS DATE)	
)
GROUP BY ZZRCRM,CONVERT(VARCHAR(3),ARAHLSTDTE,100),RIGHT(DATEPART(YEAR,ARAHLSTDTE),2)	
ORDER BY ARAHLSTDTE", end);

            string ABlcntQ = string.Format(@"SELECT ZZRCRM,CONVERT(VARCHAR(3),ARAHLSTDTE,100) AS M,RIGHT(DATEPART(YEAR,ARAHLSTDTE),2) AS Y,COUNT(*) AS CNT
FROM(SELECT ZZRCRM,ARACCTRH.ARAHLSTDTE FROM %STARTTABLE ARCLIENT INNER JOIN ARACCTRH on ARCLID = ARACCTRH.ARAHCLTID AND ARCLID IN  
                             ('BBEC01','BAGC03','BASC07','BHCC09','BHAC10','BVSC11','BBPC12','BHAC14','BBEC15','BCNC16','BBAC17','BBCC18','BSHC19','BBOC40','BAOC41','BHOC42',
							 'BBEL01', 'BAGL03', 'BASL07', 'BHCL09', 'BHAL10', 'BVSL11','BBPL12', 'BHAL14', 'BBEL15', 'BCNL16', 'BBAL17', 'BBCL18','BSHL19', 'BBOL40', 'BAOL41', 'BHOL42')
INNER JOIN ZZRETAILCLUB zz ON zz.ZZRCACID = ARACCTRH.ARAHACID
WHERE ARACCTRH.ARAHLSTLTR='IDN00004'
UNION ALL 
SELECT ZZRCRM,ARACLSTDTE	FROM %STARTTABLE ARCLIENT INNER JOIN ARACCOUNT account on ARCLID = account.ARACCLTID AND ARCLID IN  
                             ('BBEL01', 'BAGL03', 'BASL07', 'BHCL09', 'BHAL10', 'BVSL11','BBPL12', 'BHAL14', 'BBEL15', 'BCNL16', 'BBAL17', 'BBCL18','BSHL19', 'BBOL40', 'BAOL41', 'BHOL42')  
                             JOIN ARENTITY ent ON account.ARACRPID = ent.ARENID  
                             LEFT JOIN ZZRETAILCLUB zz ON zz.ZZRCACID = account.ARACID  
							 JOIN ARLTRHIS arl ON arl.ARLHACID = account.ARACID 
			  WHERE arl.ARLHREQDTE <= CAST('{0}' AS DATE) AND arl.ARLHREQDTE IS NOT NULL AND ARLHPL95 IS NOT NULL
			AND ARLHLTR='IDN00004'
)
GROUP BY ZZRCRM,CONVERT(VARCHAR(3),ARAHLSTDTE,100),RIGHT(DATEPART(YEAR,ARAHLSTDTE),2)", end);

            string pmtQ = @"SELECT fintransactiontable.AFTRACCTDTE,
    account.ARACCLACCT,
zz.ZZRCRM, entity.ARENLNM, entity.ARENFNM,zz.ZZRCCUSTID,
CASE  
   WHEN accounthistory.arahlstdte is null THEN account.araclstdte
   ELSE accounthistory.arahlstdte
END as ARACLSTDTE, 
accounthistory.arahlstdte,
ISNULL(coll.AFSCAMT,0) AS AFSCAMT,
ISNULL(fintransactiontable.AFTRAMT,0) AS AFTRAMT,
ISNULL(coll.AFSCOVERAMT,0) AS AFSCOVERAMT,
ISNULL(AFAFEE.AFAFDUEAGY,0) AS Commission,
entity.ARENST,
CASE WHEN fintransactiontable.AFTRTYP NOT IN ('CA','CC','CCC','CK','CRJ','DBJ','ECC','ECK','MC','MG','MO','NSF','CW','VI','ACH','OTH','COR','DP','DPCOR')
   THEN case when fintranstype.AFTTCATEGORY = 'M' AND fintranstype.AFTTHOWRECEIVED!='C' then 'OTH'
   when fintranstype.AFTTCATEGORY = 'M' AND fintranstype.AFTTHOWRECEIVED='C' then ''
   when fintranstype.AFTTCATEGORY = 'R' AND fintranstype.AFTTHOWRECEIVED='C' then ''
   when fintranstype.AFTTCATEGORY = 'R' AND fintranstype.AFTTHOWRECEIVED!='C' then 'COR'
   when fintranstype.AFTTCATEGORY = 'A' AND fintranstype.AFTTINCDEC='I' then 'DBJ'
   when fintranstype.AFTTCATEGORY = 'A' AND fintranstype.AFTTINCDEC='D' then 'CRJ'
   end ELSE fintransactiontable.AFTRTYP END AS TransactionTypeCode,
AFTRREFERENCE,
clientinfo.ARCLID,
((splitdetailtable.AFSTDEDUEAGY/splitdetailtable.AFSTDESPLAMT)*-1) * 100 AS Comm_Rate
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
LEFT JOIN ZZRETAILCLUB zz ON zz.ZZRCACID = account.ARACID
                                INNER JOIN AFAPPLY ON AFAPPLY.AFAPSPLID = splittable.AFSPKEY
LEFT JOIN AFAFEE ON AFAFEE.afafkey = AFAPPLY.afapfeeid 
                                LEFT JOIN AFSCOLLECT coll ON splittable.AFSPCOLID = coll.AFSCKEY
                                LEFT JOIN ARACCTRH accounthistory on accounthistory.arahacid = account.aracid and accounthistory.arahlstdte = (select min(arahlstdte) from aracctrh where arahacid = account.aracid)
 
 
WHERE clientinfo.ARCLID IN ('BBEL01','BAGL03','BASL07','BHCL09','BHAL10','BVSL11',
'BBPL12','BHAL14','BBEL15','BCNL16','BBAL17','BBCL18',
'BSHL19','BBOL40','BAOL41','BHOL42','BBEC01','BAGC03','BASC07','BHCC09','BHAC10','BVSC11','BBPC12','BHAC14','BBEC15','BCNC16','BBAC17','BBCC18','BSHC19','BBOC40','BAOC41','BHOC42')
AND " + statementSQL + @"ORDER BY ZZRCRM"; ;

            string lmtdQ = string.Format(@"SELECT fintransactiontable.AFTRENTDTE,
    account.ARACCLACCT,
zz.ZZRCRM, entity.ARENLNM, entity.ARENFNM,zz.ZZRCCUSTID,
CASE  
   WHEN accounthistory.arahlstdte is null THEN account.araclstdte
   ELSE accounthistory.arahlstdte
END as ARACLSTDTE, 
accounthistory.arahlstdte,
ISNULL(coll.AFSCAMT,0) AS AFSCAMT,
ISNULL(fintransactiontable.AFTRAMT,0) AS AFTRAMT,
ISNULL(coll.AFSCOVERAMT,0) AS AFSCOVERAMT,
ISNULL(AFAFEE.AFAFDUEAGY,0) AS Commission,
entity.ARENST,
CASE WHEN fintransactiontable.AFTRTYP NOT IN ('CA','CC','CCC','CK','CRJ','DBJ','ECC','ECK','MC','MG','MO','NSF','CW','VI','ACH','OTH','COR','DP','DPCOR')
   THEN case when fintranstype.AFTTCATEGORY = 'M' AND fintranstype.AFTTHOWRECEIVED!='C' then 'OTH'
   when fintranstype.AFTTCATEGORY = 'M' AND fintranstype.AFTTHOWRECEIVED='C' then ''
   when fintranstype.AFTTCATEGORY = 'R' AND fintranstype.AFTTHOWRECEIVED='C' then ''
   when fintranstype.AFTTCATEGORY = 'R' AND fintranstype.AFTTHOWRECEIVED!='C' then 'COR'
   when fintranstype.AFTTCATEGORY = 'A' AND fintranstype.AFTTINCDEC='I' then 'DBJ'
   when fintranstype.AFTTCATEGORY = 'A' AND fintranstype.AFTTINCDEC='D' then 'CRJ'
   end ELSE fintransactiontable.AFTRTYP END AS TransactionTypeCode,
AFTRREFERENCE,
clientinfo.ARCLID,
((splitdetailtable.AFSTDEDUEAGY/splitdetailtable.AFSTDESPLAMT)*-1) * 100 AS Comm_Rate
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
LEFT JOIN ZZRETAILCLUB zz ON zz.ZZRCACID = account.ARACID
                                INNER JOIN AFAPPLY ON AFAPPLY.AFAPSPLID = splittable.AFSPKEY
LEFT JOIN AFAFEE ON AFAFEE.afafkey = AFAPPLY.afapfeeid 
                                LEFT JOIN AFSCOLLECT coll ON splittable.AFSPCOLID = coll.AFSCKEY
                                LEFT JOIN ARACCTRH accounthistory on accounthistory.arahacid = account.aracid and accounthistory.arahlstdte = (select min(arahlstdte) from aracctrh where arahacid = account.aracid)
 
 
WHERE clientinfo.ARCLID IN ('BBEL01','BAGL03','BASL07','BHCL09','BHAL10','BVSL11',
'BBPL12','BHAL14','BBEL15','BCNL16','BBAL17','BBCL18',
'BSHL19','BBOL40','BAOL41','BHOL42','BBEC01','BAGC03','BASC07','BHCC09','BHAC10','BVSC11','BBPC12','BHAC14','BBEC15','BCNC16','BBAC17','BBCC18','BSHC19','BBOC40','BAOC41','BHOC42')
AND AFTRENTDTE <= CAST('{0}' AS DATE)
ORDER BY ZZRCRM",end);

            string conString = "Dsn=" + _environment;

            using (OdbcConnection conn = new OdbcConnection(conString))
            {
                //ACC
                using (OdbcCommand cmd = new OdbcCommand(accQ, conn, null))
                {

                    if (conn.State != ConnectionState.Open)
                        conn.Open();

                    cmd.CommandTimeout = 300000000;


                    using (var reader = cmd.ExecuteReader())
                    {
                        acc.Load(reader);
                    }




                }

                //AB Letter Count
                using (OdbcCommand cmd = new OdbcCommand(ABlcntQ, conn, null))
                {

                    if (conn.State != ConnectionState.Open)
                        conn.Open();

                    cmd.CommandTimeout = 300000000;


                    using (var reader = cmd.ExecuteReader())
                    {
                        ABlcnt.Load(reader);
                    }




                }

                //PLC
                using (OdbcCommand cmd = new OdbcCommand(plcQ, conn, null))
                {

                    if (conn.State != ConnectionState.Open)
                        conn.Open();

                    cmd.CommandTimeout = 300000000;


                    using (var reader = cmd.ExecuteReader())
                    {
                        plc.Load(reader);
                    }




                }


                


                //PMT
                using (OdbcCommand cmd = new OdbcCommand(pmtQ, conn, null))
                {

                    if (conn.State != ConnectionState.Open)
                        conn.Open();

                    cmd.CommandTimeout = 300000000;
                    using (var reader = cmd.ExecuteReader())
                    {
                        pmt.Load(reader);
                    }


                }

                                


                //ALLPAY
                using (OdbcCommand cmd = new OdbcCommand(lmtdQ, conn, null))
                {
                    if (conn.State != ConnectionState.Open)
                        conn.Open();

                    cmd.CommandTimeout = 300000000;
                    using (var reader = cmd.ExecuteReader())
                    {
                        lmtd.Load(reader);
                    }


                }

            }

            Dictionary<string, string> lettercnts = new Dictionary<string, string>();
            Dictionary<string, string> lettercnts_gprdt = new Dictionary<string, string>();
            Dictionary<string, string> lettercnts_mtsry = new Dictionary<string, string>();
            HashSet<string> rooms = new HashSet<string>();


            try
            {
                var query = from row in acc.AsEnumerable()
                            group row by row.Field<string>("ARACCLACCT") into accs
                            orderby accs.Key
                            select new
                            {
                                ID = accs.Key,
                                cnt = accs.Count(x => x.Field<string>("ARLHLTR") == "IDN00004"
                                    && x.Field<string>("ARLHPL95") != null && x.Field<DateTime?>("ARLHREQDTE") != null
                                    )
                            };

                var query2 = from row in acc.AsEnumerable()
                             group row by row.Field<string>("ZZRCRM") into accs
                             orderby accs.Key
                             select new
                             {
                                 ID = accs.Key,
                                 cnt = accs.Count(x => x.Field<string>("ARLHLTR") == "IDN00004"
                                 && x.Field<string>("ARLHPL95") != null && x.Field<DateTime?>("ARLHREQDTE") != null)
                             };


                foreach (var id in query)
                {

                    lettercnts.Add(id.ID, id.cnt.ToString() + "-" + (id.cnt * 0.50).ToString());

                }

                foreach (var id in query2)
                {

                    lettercnts_mtsry.Add(id.ID, id.cnt.ToString() + "-" + (id.cnt * 0.50).ToString());

                }
            }
            catch (Exception ex)
            {

                Console.WriteLine("@@BLOCK-1@@" + ex.Message + "@@" + ex.StackTrace);



            }

            var ids = new StringBuilder();
            int tot = 4;
            string dlabel = DateTime.Parse(start).ToString("MMM") + "-" + DateTime.Parse(start).ToString("yyyy");

            Microsoft.Office.Interop.Excel.Application oXL = new Microsoft.Office.Interop.Excel.Application();
            oXL.Visible = false;
            oXL.DisplayAlerts = false;
            oXL.StandardFont = "Calibri";
            oXL.StandardFontSize = 11;
            oXL.ScreenUpdating = false;
            Microsoft.Office.Interop.Excel.Workbook oWB = oXL.Workbooks.Add(missing);
            Microsoft.Office.Interop.Excel.Worksheet oSheet = oWB.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;


            oSheet.Name = "Report F - Letterdetail";

            oSheet.Cells[1, 1] = "Report F - Month To Date Letters Sent Detail: ";
            oSheet.Cells[1, 2] = dlabel;


            oSheet.Cells[2, 1] = "NOTE: Info pulled for accounts while in the 45day Letter phase";


            oSheet.Cells[4, 1] = "Division";
            oSheet.Cells[4, 2] = "Customer Name";
            oSheet.Cells[4, 3] = "Customer Number";
            oSheet.Cells[4, 4] = "Order ID #";
            oSheet.Cells[4, 5] = "Place Date";
            oSheet.Cells[4, 6] = "Letter Date";
            oSheet.Cells[4, 7] = "# of Letters sent";
            oSheet.Cells[4, 8] = "Pay Amount to Agency";
            oSheet.Cells[4, 9] = "Amount Per Letter";
            oSheet.Cells[4, 10] = "State";

            tot++;

            string[] a = null;
            string c = "";
            string j = "";

            string ARENFNM = "", ZZRCCUSTID = "", ARACCLACCT = "", ARENLNM = "", ZZRCRM = "";
            int at = 0;
            int ta = 0;
            decimal bt = 0;
            int d = 0;
            decimal tb = 0, tc = 0;
            object[,] arr = new object[acc.Rows.Count, 10];
            try
            {
                foreach (DataRow r in acc.Rows)
                {


                    try
                    {
                        ARENFNM = r["ARENFNM"] == DBNull.Value || r["ARENFNM"] == null || r["ARENFNM"].ToString() == "" ? "" : r["ARENFNM"].ToString();
                        ZZRCCUSTID = r["ZZRCCUSTID"] == DBNull.Value || r["ZZRCCUSTID"] == null || r["ZZRCCUSTID"].ToString() == "" ? "" : r["ZZRCCUSTID"].ToString();
                        ARACCLACCT = r["ARACCLACCT"] == DBNull.Value || r["ARACCLACCT"] == null || r["ARACCLACCT"].ToString() == "" ? "" : r["ARACCLACCT"].ToString();
                        ARENLNM = r["ARENLNM"] == DBNull.Value || r["ARENLNM"] == null || r["ARENLNM"].ToString() == "" ? "" : r["ARENLNM"].ToString();
                        ZZRCRM = r["ZZRCRM"] == DBNull.Value || r["ZZRCRM"] == null || r["ZZRCRM"].ToString() == "" ? "" : r["ZZRCRM"].ToString();
                    }
                    catch (Exception ex)
                    {

                        Console.WriteLine("@@BLOCK-2-REPORT-F@@" + ex.Message + "@@" + ex.StackTrace);
                        
                    }

                    
                   

                    if (!string.IsNullOrEmpty(ZZRCRM))
                    {

                        a = lettercnts[ARACCLACCT].Split('-');
                        ta = Convert.ToInt32(a[0]);
                        tb = Convert.ToDecimal(a[1]);
                        at += ta;
                        bt += tb;
                        tc += 0.50m;
                    }


                    try
                    {
                        oSheet.Range["A" + tot.ToString(), "J" + tot.ToString()].Value = new Object[,] { { ZZRCRM
                        , ARENLNM + "," + ARENFNM
                        ,"'"+r["ZZRCCUSTID"]
                        ,"'"+r["ARACCLACCT"]
                        ,(DateTime)r["ARACLSTDTE"]
                        ,(DateTime)r["ARLHUPDDTE"]
                        ,ta
                        ,tb
                        ,0.50
                        ,r["ARENST"] }};

                        formatrDF(oSheet, tot, "H", "I");

                    }
                    catch (Exception ex)
                    {

                        Console.WriteLine("@@BLOCK-3-REPORT-F@@" + ex.Message + "@@" + ex.StackTrace);
                    }



                    tot++;



                }



                //oSheet.Range["A5", "J" + acc.Rows.Count.ToString()].set_Value(missing,arr);


                tot++;
                oSheet.Cells[tot, 1] = "Totals";
                oSheet.Cells[tot, 7] = at;
                oSheet.Cells[tot, 8] = bt;
                oSheet.Cells[tot, 9] = tc;

                formatrDF(oSheet, tot, "H", "I");

                oSheet.Cells.Range["A" + tot.ToString(), "A" + tot.ToString()].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbSkyBlue;
                Microsoft.Office.Interop.Excel.Range cell42 = oSheet.Cells.Range["A" + tot.ToString(), "A" + tot.ToString()].Cells;
                Microsoft.Office.Interop.Excel.Borders border42 = cell42.Borders;

                border42.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                border42.Weight = 2d;




            }
            catch (Exception ex)
            {

                Console.WriteLine("@@BLOCK-4-REPORT-F@@" + ex.Message + "@@" + ex.StackTrace);
            }



            //Report E
            Microsoft.Office.Interop.Excel.Worksheet oSheet2 = oWB.Sheets.Add(missing, missing, 1, missing)
                            as Microsoft.Office.Interop.Excel.Worksheet;
            oSheet2.Name = "Report E - OverPaymentdetail";

            oSheet2.Cells[1, 1] = "Report E -  OverPayment Detail: ";
            oSheet2.Cells[1, 2] = dlabel;


            decimal rate = 0.0m;
            string rt = "";
            string pt = "";
            decimal drt = 0.0m;
            decimal comovg = 0.0m;
            decimal sumcomovg = 0.0m;


            Dictionary<string, decimal> pagency35 = new Dictionary<string, decimal>();
            Dictionary<string, decimal> pagency55 = new Dictionary<string, decimal>();
            Dictionary<string, decimal> pbrad35 = new Dictionary<string, decimal>();
            Dictionary<string, decimal> pbrad55 = new Dictionary<string, decimal>();

            Dictionary<string, decimal> mpagency35 = new Dictionary<string, decimal>();
            Dictionary<string, decimal> mpagency55 = new Dictionary<string, decimal>();
            Dictionary<string, decimal> mpbrad35 = new Dictionary<string, decimal>();
            Dictionary<string, decimal> mpbrad55 = new Dictionary<string, decimal>();

            Dictionary<string, decimal> nocomm = new Dictionary<string, decimal>();
            decimal overpmttot = 0;





            tot = 4;


            oSheet2.Cells[tot, 1] = "Division";
            oSheet2.Cells[tot, 2] = "Customer Name";
            oSheet2.Cells[tot, 3] = "Customer Number";
            oSheet2.Cells[tot, 4] = "Order ID #";
            oSheet2.Cells[tot, 5] = "Place Date";
            oSheet2.Cells[tot, 6] = "Pay Date";
            oSheet2.Cells[tot, 7] = "Amount Due";
            oSheet2.Cells[tot, 8] = "Payment Amount";
            oSheet2.Cells[tot, 9] = "OvrPymnt Amount";
            oSheet2.Cells[tot, 10] = "Comm. Amt";
            oSheet2.Cells[tot, 11] = "Rate";
            oSheet2.Cells[tot, 12] = "Comm.Amt Overage";
            oSheet2.Cells[tot, 13] = "Payment Type";
            oSheet2.Cells[tot, 14] = "State";

            tot++;

            //report C group by rooms
            string[] dps = { "DP", "DPCOR" };

            var pagency = pmt.AsEnumerable()
    .Select(x => new
    {
        room = x.Field<string>("ZZRCRM"),
        listdate = x.Field<DateTime>("ARACLSTDTE").ToString("MMM-yyyy"),
        Commission = x.Field<Decimal>("Commission"),
        AFTRAMT = x.Field<Decimal>("AFTRAMT"),
        AFSCAMT = x.Field<Decimal>("AFSCAMT"),
        CLIENTID = x.Field<string>("ARCLID"),
        trantype = x.Field<string>("TransactionTypeCode"),
        commrate = x.Field<Decimal>("Comm_Rate")

    })
    .GroupBy(x => new { x.room })
    .Select(g => new
    {

        rm = g.Key.room,
        sum = g.Sum(x => x.commrate >= 10.0m &&
    x.commrate <= 53.0m && x.trantype != "DP" && x.trantype != "DPCOR" ? x.AFSCAMT : 0)


    });

            foreach (var i in pagency)
            {
                pagency35.Add(i.rm.Trim(), i.sum);

            }

            pagency = pmt.AsEnumerable()
    .Select(x => new
    {
        room = x.Field<string>("ZZRCRM"),
        listdate = x.Field<DateTime>("ARACLSTDTE").ToString("MMM-yyyy"),
        Commission = x.Field<Decimal>("Commission"),
        AFTRAMT = x.Field<Decimal>("AFTRAMT"),
        AFSCAMT = x.Field<Decimal>("AFSCAMT"),
        CLIENTID = x.Field<string>("ARCLID"),
        trantype = x.Field<string>("TransactionTypeCode"),
        commrate = x.Field<Decimal>("Comm_Rate")

    })
    .GroupBy(x => new { x.room })
    .Select(g => new
    {

        rm = g.Key.room,
        sum = g.Sum(x => x.commrate >= 54.8m &&
        x.commrate <= 55.5m && x.trantype != "DP" && x.trantype != "DPCOR" ? x.AFSCAMT : 0)


    });

            foreach (var i in pagency)
            {
                pagency55.Add(i.rm.Trim(), i.sum);

            }

            //paid to bradford
            var pbrad = pmt.AsEnumerable()
    .Select(x => new
    {
        room = x.Field<string>("ZZRCRM"),
        listdate = x.Field<DateTime>("ARACLSTDTE").ToString("MMM-yyyy"),
        Commission = x.Field<Decimal>("Commission"),
        AFTRAMT = x.Field<Decimal>("AFTRAMT"),
        AFSCAMT = x.Field<Decimal>("AFSCAMT"),
        CLIENTID = x.Field<string>("ARCLID"),
        trantype = x.Field<string>("TransactionTypeCode"),
        commrate = x.Field<Decimal>("Comm_Rate")

    })
    .GroupBy(x => new { x.room })
    .Select(g => new
    {

        rm = g.Key.room,
        sum = g.Sum(x => x.commrate >= 10.0m &&
        x.commrate <= 53.0m && dps.Contains(x.trantype) && collectionphase.Contains(x.CLIENTID) ? x.AFSCAMT : 0)


    });


            foreach (var i in pbrad)
            {
                pbrad35.Add(i.rm.Trim(), i.sum);

            }


            pbrad = pmt.AsEnumerable()
    .Select(x => new
    {
        room = x.Field<string>("ZZRCRM"),
        listdate = x.Field<DateTime>("ARACLSTDTE").ToString("MMM-yyyy"),
        Commission = x.Field<Decimal>("Commission"),
        AFTRAMT = x.Field<Decimal>("AFTRAMT"),
        AFSCAMT = x.Field<Decimal>("AFSCAMT"),
        CLIENTID = x.Field<string>("ARCLID"),
        trantype = x.Field<string>("TransactionTypeCode"),
        commrate = x.Field<Decimal>("Comm_Rate")

    })
    .GroupBy(x => new { x.room })
    .Select(g => new
    {

        rm = g.Key.room,
        sum = g.Sum(x => x.commrate >= 54.8m &&
        x.commrate <= 55.5m && dps.Contains(x.trantype) && collectionphase.Contains(x.CLIENTID) ? x.AFSCAMT : 0)


    });


            foreach (var i in pbrad)
            {
                pbrad55.Add(i.rm.Trim(), i.sum);

            }


            // for report B group by listdate and room
            var pagency2 = pmt.AsEnumerable()
    .Select(x => new
    {
        room = x.Field<string>("ZZRCRM"),
        listdate = x.Field<DateTime>("ARACLSTDTE").ToString("MMM-yyyy"),
        Commission = x.Field<Decimal>("Commission"),
        AFTRAMT = x.Field<Decimal>("AFTRAMT"),
        AFSCAMT = x.Field<Decimal>("AFSCAMT"),
        CLIENTID = x.Field<string>("ARCLID"),
        trantype = x.Field<string>("TransactionTypeCode"),
        commrate = x.Field<Decimal>("Comm_Rate")

    })
    .GroupBy(x => new { x.room, x.listdate })
    .Select(g => new
    {

        rm = g.Key.room,
        dt = g.Key.listdate,
        sum = g.Sum(x => x.commrate >= 10.0m &&
        x.commrate <= 53.0m && x.trantype != "DP" && x.trantype != "DPCOR" ? x.AFSCAMT : 0)


    });

            foreach (var i in pagency2)
            {
                mpagency35.Add(i.rm.Trim() + "." + i.dt.Remove(i.dt.Length - 4, 2), i.sum);

            }

            var pagency3 = pmt.AsEnumerable()
    .Select(x => new
    {
        room = x.Field<string>("ZZRCRM"),
        listdate = x.Field<DateTime>("ARACLSTDTE").ToString("MMM-yyyy"),
        Commission = x.Field<Decimal>("Commission"),
        AFTRAMT = x.Field<Decimal>("AFTRAMT"),
        AFSCAMT = x.Field<Decimal>("AFSCAMT"),
        CLIENTID = x.Field<string>("ARCLID"),
        trantype = x.Field<string>("TransactionTypeCode"),
        commrate = x.Field<Decimal>("Comm_Rate")

    })
    .GroupBy(x => new { x.room, x.listdate })
    .Select(g => new
    {

        rm = g.Key.room,
        dt = g.Key.listdate,
        sum = g.Sum(x => x.commrate >= 54.8m &&
        x.commrate <= 55.5m && x.trantype != "DP" && x.trantype != "DPCOR" ? x.AFSCAMT : 0)


    });

            foreach (var i in pagency3)
            {
                mpagency55.Add(i.rm.Trim() + "." + i.dt.Remove(i.dt.Length - 4, 2), i.sum);

            }

            //paid to bradford
            var pbrad2 = pmt.AsEnumerable()
    .Select(x => new
    {
        room = x.Field<string>("ZZRCRM"),
        listdate = x.Field<DateTime>("ARACLSTDTE").ToString("MMM-yyyy"),
        Commission = x.Field<Decimal>("Commission"),
        AFTRAMT = x.Field<Decimal>("AFTRAMT"),
        AFSCAMT = x.Field<Decimal>("AFSCAMT"),
        CLIENTID = x.Field<string>("ARCLID"),
        trantype = x.Field<string>("TransactionTypeCode"),
        commrate = x.Field<Decimal>("Comm_Rate")

    })
    .GroupBy(x => new { x.room, x.listdate })
    .Select(g => new
    {

        rm = g.Key.room,
        dt = g.Key.listdate,
        sum = g.Sum(x => x.commrate >= 10.0m &&
        x.commrate <= 53.0m && dps.Contains(x.trantype) && collectionphase.Contains(x.CLIENTID) ? x.AFSCAMT : 0)


    });


            foreach (var i in pbrad2)
            {
                mpbrad35.Add(i.rm.Trim() + "." + i.dt.Remove(i.dt.Length - 4, 2), i.sum);

            }


            var pbrad3 = pmt.AsEnumerable()
    .Select(x => new
    {
        room = x.Field<string>("ZZRCRM"),
        listdate = x.Field<DateTime>("ARACLSTDTE").ToString("MMM-yyyy"),
        Commission = x.Field<Decimal>("Commission"),
        AFTRAMT = x.Field<Decimal>("AFTRAMT"),
        AFSCAMT = x.Field<Decimal>("AFSCAMT"),
        CLIENTID = x.Field<string>("ARCLID"),
        trantype = x.Field<string>("TransactionTypeCode"),
        commrate = x.Field<Decimal>("Comm_Rate")

    })
    .GroupBy(x => new { x.room, x.listdate })
    .Select(g => new
    {

        rm = g.Key.room,
        dt = g.Key.listdate,
        sum = g.Sum(x => x.commrate >= 54.8m &&
        x.commrate <= 55.5m && dps.Contains(x.trantype) && collectionphase.Contains(x.CLIENTID) ? x.AFSCAMT : 0)


    });


            foreach (var i in pbrad3)
            {
                mpbrad55.Add(i.rm.Trim() + "." + i.dt.Remove(i.dt.Length - 4, 2), i.sum);

            }



            decimal Commission = 0, AFTRAMT = 0, AFSCAMT = 0, AFSCOVERAMT = 0;
            string trantype = "", CLIENTID = "";
            bool isdp = false;

            try
            {
                foreach (DataRow r in pmt.Rows)
                {
                    trantype = r["TransactionTypeCode"] == DBNull.Value || r["TransactionTypeCode"] == null || r["TransactionTypeCode"].ToString() == "" ? "" : r["TransactionTypeCode"].ToString();


                    ARENFNM = r["ARENFNM"] == DBNull.Value || r["ARENFNM"] == null || r["ARENFNM"].ToString() == "" ? "" : r["ARENFNM"].ToString();
                    ARENLNM = r["ARENLNM"] == DBNull.Value || r["ARENLNM"] == null || r["ARENLNM"].ToString() == "" ? "" : r["ARENLNM"].ToString();
                    ZZRCRM = r["ZZRCRM"] == DBNull.Value || r["ZZRCRM"] == null || r["ZZRCRM"].ToString() == "" ? "" : r["ZZRCRM"].ToString();

                    rate = r["Comm_Rate"] == DBNull.Value || r["Comm_Rate"] == null || r["Comm_Rate"].ToString() == "" ? 0 : (decimal)r["Comm_Rate"];
                    Commission = r["Commission"] == DBNull.Value || r["Commission"] == null || r["Commission"].ToString() == "" ? 0 : (decimal)r["Commission"];
                    AFTRAMT = r["AFTRAMT"] == DBNull.Value || r["AFTRAMT"] == null || r["AFTRAMT"].ToString() == "" ? 0 : (decimal)r["AFTRAMT"];
                    AFSCAMT = r["AFSCAMT"] == DBNull.Value || r["AFSCAMT"] == null || r["AFSCAMT"].ToString() == "" ? 0 : (decimal)r["AFSCAMT"];
                    AFSCOVERAMT = r["AFSCOVERAMT"] == DBNull.Value || r["AFSCOVERAMT"] == null || r["AFSCOVERAMT"].ToString() == "" ? 0 : (decimal)r["AFSCOVERAMT"];
                    CLIENTID = r["ARCLID"] == DBNull.Value || r["ARCLID"] == null || r["ARCLID"].ToString() == "" ? "" : r["ARCLID"].ToString();

                    
                    DateTime dt = (DateTime)r["ARACLSTDTE"];
                    string key = r["ZZRCRM"].ToString() + "." + string.Format("{0}-{1}", dt.ToString("MMM"), dt.ToString("yyyy").Substring(dt.ToString("yyyy").Length - 2, 2));

                    if (trantype != "DP" && trantype != "DPCOR" && Commission == 0)
                    {
                        if (nocomm.ContainsKey(ZZRCRM))
                            nocomm[ZZRCRM] += AFSCAMT;// AFTRAMT;
                        else
                            nocomm[ZZRCRM] = AFSCAMT;// AFTRAMT;
                    }
                    rt = "";

                    if (Commission > 0 && rate >= 10.0m && rate <= 53.0m)
                    {
                        rt = "35%";
                        drt = 0.35m;

                    }

                    if (Commission > 0 && rate >= 54.8m && rate <= 55.5m)
                    {
                        rt = "55%";
                        drt = 0.55m;

                    }

                    if (r["AFSCOVERAMT"] != DBNull.Value && r["AFSCOVERAMT"] != null && Convert.ToDecimal(r["AFSCOVERAMT"]) > 0 && trantype != "DP" && trantype != "DPCOR")
                    {

                        isdp = true;
                        decimal ovpmt = AFSCOVERAMT;

                        overpmttot += ovpmt;

                        if(Commission>0)
                            comovg = ovpmt * drt;
                        else
                            comovg = 0;

                        sumcomovg += comovg;

                        oSheet2.Range["A" + tot.ToString(), "N" + tot.ToString()].Value = new Object[,] { { ZZRCRM
                        , ARENLNM + "," + ARENFNM
                        ,"'"+r["ZZRCCUSTID"].ToString()
                        ,"'"+r["ARACCLACCT"].ToString()
                        ,(DateTime)r["ARACLSTDTE"]
                        ,(DateTime)r["AFTRACCTDTE"]
                        ,AFSCAMT
                        ,AFTRAMT
                        ,ovpmt
                        ,Commission
                        ,rt
                        ,comovg
                        ,r["TransactionTypeCode"]
                        ,r["ARENST"] }};





                        formatsheet23(oSheet2, tot, true);

                        tot++;

                    }




                }



                if (isdp)
                {
                    tot++;
                    oSheet2.Cells[tot, 1] = "Totals";
                    oSheet2.Cells[tot, 8] = Convert.ToDecimal(pmt.Compute("SUM(AFTRAMT)", "AFSCOVERAMT IS NOT NULL AND AFSCOVERAMT>0 AND TransactionTypeCode<>'DP'"));
                    oSheet2.Cells[tot, 9] = Convert.ToDecimal(pmt.Compute("SUM(AFSCOVERAMT)", "AFSCOVERAMT IS NOT NULL AND AFSCOVERAMT>0 AND TransactionTypeCode<>'DP'"));
                    oSheet2.Cells[tot, 10] = Convert.ToDecimal(pmt.Compute("SUM(Commission)", "AFSCOVERAMT IS NOT NULL AND AFSCOVERAMT>0 AND TransactionTypeCode<>'DP'"));
                    oSheet2.Cells[tot, 12] = sumcomovg;

                    oSheet2.Cells.Range["A" + tot.ToString(), "A" + tot.ToString()].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbSkyBlue;
                    Microsoft.Office.Interop.Excel.Range cell4 = oSheet2.Cells.Range["A" + tot.ToString(), "A" + tot.ToString()].Cells;
                    Microsoft.Office.Interop.Excel.Borders border4 = cell4.Borders;

                    border4.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    border4.Weight = 2d;

                    formatsheet23(oSheet2, tot, true);

                }
            }

            catch (Exception ex)
            {

                Console.WriteLine("@@BLOCK-5-REPORT-E@@" + ex.Message + "@@" + ex.StackTrace);
            }

            //sumcomovg = 0.0m;

            //Report D
            Microsoft.Office.Interop.Excel.Worksheet oSheet3 = oWB.Sheets.Add(missing, missing, 1, missing)
                    as Microsoft.Office.Interop.Excel.Worksheet;
            oSheet3.Name = "Report D - Paymentdetail";

            oSheet3.Cells[1, 1] = "Report D - Payment Detail: ";
            oSheet3.Cells[1, 2] = dlabel;

            rate = 0.0m;
            rt = "";
            pt = "";
            comovg = 0.0m;
            //sumcomovg = 0.0m;
            tot = 4;
            c = "";
            j = "";



            oSheet3.Cells[tot, 1] = "Division";
            oSheet3.Cells[tot, 2] = "Customer Name";
            oSheet3.Cells[tot, 3] = "Customer Number";
            oSheet3.Cells[tot, 4] = "Order ID #";
            oSheet3.Cells[tot, 5] = "Place Date";
            oSheet3.Cells[tot, 6] = "Pay Date";
            oSheet3.Cells[tot, 7] = "Pay Amount";
            oSheet3.Cells[tot, 8] = "Comm. Amt";
            oSheet3.Cells[tot, 9] = "Rate";
            oSheet3.Cells[tot, 10] = "Payment Type";
            oSheet3.Cells[tot, 11] = "State";

            tot++;


            try
            {
                foreach (DataRow r in pmt.Rows)
                {

                    ARENFNM = r["ARENFNM"] == DBNull.Value || r["ARENFNM"] == null || r["ARENFNM"].ToString() == "" ? "" : r["ARENFNM"].ToString();
                    ARENLNM = r["ARENLNM"] == DBNull.Value || r["ARENLNM"] == null || r["ARENLNM"].ToString() == "" ? "" : r["ARENLNM"].ToString();
                    ZZRCRM = r["ZZRCRM"] == DBNull.Value || r["ZZRCRM"] == null || r["ZZRCRM"].ToString() == "" ? "" : r["ZZRCRM"].ToString();
                    trantype = r["TransactionTypeCode"] == DBNull.Value || r["TransactionTypeCode"] == null || r["TransactionTypeCode"].ToString() == "" ? "" : r["TransactionTypeCode"].ToString();
                    rate = r["Comm_Rate"] == DBNull.Value || r["Comm_Rate"] == null || r["Comm_Rate"].ToString() == "" ? 0 : (decimal)r["Comm_Rate"];
                    Commission = r["Commission"] == DBNull.Value || r["Commission"] == null || r["Commission"].ToString() == "" ? 0 : (decimal)r["Commission"];
                    AFTRAMT = r["AFTRAMT"] == DBNull.Value || r["AFTRAMT"] == null || r["AFTRAMT"].ToString() == "" ? 0 : (decimal)r["AFTRAMT"];
                    AFSCAMT = r["AFSCAMT"] == DBNull.Value || r["AFSCAMT"] == null || r["AFSCAMT"].ToString() == "" ? 0 : (decimal)r["AFSCAMT"];
                    AFSCOVERAMT = r["AFSCOVERAMT"] == DBNull.Value || r["AFSCOVERAMT"] == null || r["AFSCOVERAMT"].ToString() == "" ? 0 : (decimal)r["AFSCOVERAMT"];

                    rooms.Add(ZZRCRM);

                    rt = "";

                    if (Commission > 0 && rate >= 10.0m && rate <= 53.0m)
                    {
                        rt = "35%";


                    }

                    if (Commission > 0 && rate >= 54.8m && rate <= 55.5m)
                    {
                        rt = "55%";

                    }


                    try
                    {
                        
                        oSheet3.Range["A" + tot.ToString(), "K" + tot.ToString()].Value = new Object[,] { { ZZRCRM
                        ,ARENLNM + "," + ARENFNM
                        ,"'"+r["ZZRCCUSTID"]
                        ,"'"+r["ARACCLACCT"]
                        ,(DateTime)r["ARACLSTDTE"]
                        ,(DateTime)r["AFTRACCTDTE"]
                        ,AFSCAMT
                        ,Commission
                        ,rt
                        ,r["TransactionTypeCode"]
                        ,r["ARENST"] }};

                    }
                    catch (Exception ex)
                    {

                        Console.WriteLine("@@BLOCK-6-REPORT-D@@" + ex.Message + "@@" + ex.StackTrace);
                    }

                    formatsheet23(oSheet3, tot, false);

                    tot++;

                }

                tot++;

                oSheet3.Cells[tot, 1] = "Totals";
                oSheet3.Cells[tot, 7] = Convert.ToDecimal(pmt.Compute("SUM(AFSCAMT)", "AFTRAMT IS NOT NULL"));
                oSheet3.Cells[tot, 8] = Convert.ToDecimal(pmt.Compute("SUM(Commission)", "Commission IS NOT NULL"));

                oSheet3.Cells.Range["A" + tot.ToString(), "A" + tot.ToString()].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbSkyBlue;
                Microsoft.Office.Interop.Excel.Range cell4 = oSheet3.Cells.Range["A" + tot.ToString(), "A" + tot.ToString()].Cells;
                Microsoft.Office.Interop.Excel.Borders border4 = cell4.Borders;

                border4.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                border4.Weight = 2d;

                formatsheet23(oSheet3, tot, false);

            }

            catch (Exception ex)
            {

                Console.WriteLine("@@BLOCK-7-REPORT-D@@" + ex.Message + "@@" + ex.StackTrace);
            }
            


            //Report C
            Microsoft.Office.Interop.Excel.Worksheet oSheet4 = oWB.Sheets.Add(missing, missing, 1, missing)
                        as Microsoft.Office.Interop.Excel.Worksheet;
            oSheet4.Name = "Report C - Monthly Remit Smry";

            oSheet4.Cells[1, 1] = "Report C - Monthly Remit Smry: ";
            oSheet4.Cells[1, 2] = dlabel;

            tot = 4;


            oSheet4.Cells[tot, 1] = "";
            oSheet4.Cells[tot, 2] = "Nbr 3rd Pty";
            oSheet4.Cells[tot, 3] = "Cost Per";
            oSheet4.Cells[tot, 4] = "3rd Pty Letter";
            oSheet4.Cells[tot, 5] = "Paid Bradford group exchange";
            oSheet4.Cells[tot, 6] = "Paid Agency";
            oSheet4.Cells[tot, 7] = "Paid Bradford group exchange";
            oSheet4.Cells[tot, 8] = "Paid Agency";
            oSheet4.Cells[tot, 9] = "Paid Agency";
            oSheet4.Cells[tot, 10] = "";


            tot++;

            oSheet4.Cells[tot, 1] = "Division";
            oSheet4.Cells[tot, 2] = "Letters";
            oSheet4.Cells[tot, 3] = "Letter";
            oSheet4.Cells[tot, 4] = "Cost";
            oSheet4.Cells[tot, 5] = "35% Cont.";
            oSheet4.Cells[tot, 6] = "35% Cont.";
            oSheet4.Cells[tot, 7] = "55% Cont.";
            oSheet4.Cells[tot, 8] = "55% Cont.";
            oSheet4.Cells[tot, 9] = "No Comm.";
            oSheet4.Cells[tot, 10] = "Total";


            tot++;

            decimal ncom = 0;
            decimal pb35 = 0;
            decimal pb55 = 0;
            decimal pa35 = 0;
            decimal pa55 = 0;

            decimal tncom = 0;
            decimal tpb35 = 0;
            decimal tpb55 = 0;
            decimal tpa35 = 0;
            decimal tpa55 = 0;
            decimal letterc = 0;
            int lettert = 0;

            decimal tct = 0;
            decimal ct = 0;
            decimal lcv = 0;
            int tvt = 0;

            try
            {
                foreach (string r in rooms)
                {



                    if (pbrad35.ContainsKey(r))
                        pb35 = Math.Abs(pbrad35[r]);

                    if (pbrad55.ContainsKey(r))
                        pb55 = Math.Abs(pbrad55[r]);

                    if (pagency35.ContainsKey(r))
                        pa35 = Math.Abs(pagency35[r]);

                    if (pagency55.ContainsKey(r))
                        pa55 = Math.Abs(pagency55[r]);

                    if (nocomm.ContainsKey(r))
                        ncom = Math.Abs(nocomm[r]);
                    else
                        ncom = 0m;

                    if (lettercnts_mtsry.ContainsKey(r))
                    {
                        lcv = Convert.ToDecimal(lettercnts_mtsry[r].Split('-')[1]);
                        tvt = Convert.ToInt32(lettercnts_mtsry[r].Split('-')[0]);

                    }
                    else
                    {
                        lcv = 0;
                        tvt = 0;
                    }

                    ct = ncom + pa35 * 0.65m + pa55 * 0.45m - Math.Abs(lcv) - pb35 * 0.35m - pb55 * 0.55m;

                    oSheet4.Range["A" + tot.ToString(), "J" + tot.ToString()].Value = new Object[,] { { r
                        , tvt.ToString()
                        ,0.50
                        ,Math.Abs(lcv)
                        ,pb35
                        ,pa35
                        ,pb55
                        ,pa55
                        ,ncom
                        ,ct }};


                    formatMsry(oSheet4, tot, true);
                    oSheet4.Cells.Range["J" + tot.ToString(), "J" + tot.ToString()].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellowGreen;

                    tncom += ncom;
                    tpb35 += pb35;
                    tpb55 += pb55;
                    tpa35 += pa35;
                    tpa55 += pa55;
                    letterc += lcv;
                    lettert += tvt;
                    tct += ct;

                    tot++;


                }



                oSheet4.Cells[tot, 1] = "Total All";
                oSheet4.Cells[tot, 2] = lettert;
                oSheet4.Cells[tot, 4] = letterc;
                oSheet4.Cells[tot, 5] = tpb35;
                oSheet4.Cells[tot, 6] = tpa35;
                oSheet4.Cells[tot, 7] = tpb55;
                oSheet4.Cells[tot, 8] = tpa55;
                oSheet4.Cells[tot, 9] = tncom;
                oSheet4.Cells[tot, 10] = tct;

                formatMsry(oSheet4, tot, true);

            }
            catch (Exception ex)
            {

                Console.WriteLine("@@BLOCK-8-REPORT-C@@" + ex.Message + "@@" + ex.StackTrace);
            }

            oSheet4.Cells.Range["A" + tot.ToString(), "I" + tot.ToString()].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbSkyBlue;
            oSheet4.Cells.Range["J" + tot.ToString()].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellowGreen;
            Microsoft.Office.Interop.Excel.Range cellA = oSheet4.Cells.Range["A" + tot.ToString(), "J" + tot.ToString()].Cells;
            Microsoft.Office.Interop.Excel.Borders borderA = cellA.Borders;

            borderA.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            borderA.Weight = 2d;

            tot += 3;

            oSheet4.Cells[tot, 7] = "Sub Total";
            oSheet4.Cells[tot, 8] = tct;
            oSheet4.Cells.Range["G" + tot.ToString(), "H" + tot.ToString()].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellowGreen;

            Microsoft.Office.Interop.Excel.Range cellB = oSheet4.Cells.Range["G" + tot.ToString(), "H" + tot.ToString()].Cells;
            Microsoft.Office.Interop.Excel.Borders borderB = cellB.Borders;

            borderB.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            borderB.Weight = 2d;
            formatMsry(oSheet4, tot, false);
            tot++;
            oSheet4.Cells[tot, 7] = "Overpayments Paid to Agency";
            oSheet4.Cells[tot, 8] = sumcomovg;// overpmttot;

            oSheet4.Cells.Range["G" + tot.ToString(), "H" + tot.ToString()].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellowGreen;

            Microsoft.Office.Interop.Excel.Range cellC = oSheet4.Cells.Range["G" + tot.ToString(), "H" + tot.ToString()].Cells;
            Microsoft.Office.Interop.Excel.Borders borderC = cellC.Borders;

            borderC.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            borderC.Weight = 2d;

            formatMsry(oSheet4, tot, false);

            tot++;
            oSheet4.Cells[tot, 7] = "Sub Total";
            oSheet4.Cells[tot, 8] = (Math.Abs(tct) - sumcomovg)*-1;

            oSheet4.Cells.Range["G" + tot.ToString(), "H" + tot.ToString()].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellowGreen;

            Microsoft.Office.Interop.Excel.Range cellD = oSheet4.Cells.Range["G" + tot.ToString(), "H" + tot.ToString()].Cells;
            Microsoft.Office.Interop.Excel.Borders borderD = cellD.Borders;

            borderD.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            borderD.Weight = 2d;
            formatMsry(oSheet4, tot, false);
            tot++;
            oSheet4.Cells[tot, 7] = "Bonus / Penalty";
            oSheet4.Cells[tot, 8] = "";
            oSheet4.Cells.Range["G" + tot.ToString(), "H" + tot.ToString()].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellowGreen;

            Microsoft.Office.Interop.Excel.Range cellE = oSheet4.Cells.Range["G" + tot.ToString(), "H" + tot.ToString()].Cells;
            Microsoft.Office.Interop.Excel.Borders borderE = cellE.Borders;

            borderE.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            borderE.Weight = 2d;


            tot++;
            oSheet4.Cells[tot, 7] = "TOTAL WIRE TO BGE";
            oSheet4.Cells[tot, 8] = (Math.Abs(tct) - sumcomovg) * -1;

            oSheet4.Cells.Range["G" + tot.ToString(), "H" + tot.ToString()].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellowGreen;

            Microsoft.Office.Interop.Excel.Range cellF = oSheet4.Cells.Range["G" + tot.ToString(), "H" + tot.ToString()].Cells;
            Microsoft.Office.Interop.Excel.Borders borderF = cellF.Borders;

            borderF.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            borderF.Weight = 2d;

            formatMsry(oSheet4, tot, false);

            // LTD AND MTD

            pb35 = 0;
            pb55 = 0;
            pa35 = 0;
            pa55 = 0;

            pagency35.Clear();
            pagency55.Clear();
            pbrad35.Clear();
            pbrad55.Clear();


            Dictionary<string, int> pcnt = new Dictionary<string, int>();
            Dictionary<string, decimal> balsum = new Dictionary<string, decimal>();
            Dictionary<string, decimal> amtsum = new Dictionary<string, decimal>();

            Dictionary<string, decimal> mamtsum = new Dictionary<string, decimal>();


            var amtsry = lmtd.AsEnumerable()
    .Select(x => new
    {
        room = x.Field<string>("ZZRCRM"),
        listdate = x.Field<DateTime>("ARACLSTDTE").ToString("MMM-yyyy"),
        afac = x.Field<Decimal>("AFSCAMT"),
        Commission = x.Field<Decimal>("Commission"),
        AFTRAMT = x.Field<Decimal>("AFTRAMT"),
        AFSCAMT = x.Field<Decimal>("AFSCAMT"),
        CLIENTID = x.Field<string>("ARCLID"),
        trantype = x.Field<string>("TransactionTypeCode")

    })
    .GroupBy(x => new { x.room, x.listdate })
    .Select(g => new
    {

        rm = g.Key.room,
        dt = g.Key.listdate,
        Sum = g.Sum(r => r.Commission == 0 && dps.Contains(r.trantype) ? r.afac : 0)


    });


            var placementcnt = acc.AsEnumerable()
    .Select(x => new
    {
        room = x.Field<string>("ZZRCRM"),
        listdate = x.Field<DateTime>("ARACLSTDTE").ToString("MMM-yyyy"),
        ltr = x.Field<string>("ARLHLTR"),
        pl95 = x.Field<string>("ARLHPL95"),
        dte = x.Field<DateTime?>("ARLHREQDTE")

    })
    .GroupBy(x => new { x.room, x.listdate })
    .Select(g => new
    {

        rm = g.Key.room,
        dt = g.Key.listdate,
        count = g.Count(x => x.ltr == "IDN00004" && x.pl95 != null && x.dte != null)


    });


            var mamtsry = pmt.AsEnumerable()
    .Select(x => new
    {
        room = x.Field<string>("ZZRCRM"),
        listdate = x.Field<DateTime>("ARACLSTDTE").ToString("MMM-yyyy"),
        afac = x.Field<Decimal>("AFSCAMT"),
        Commission = x.Field<Decimal>("Commission"),
        AFTRAMT = x.Field<Decimal>("AFTRAMT"),
        AFSCAMT = x.Field<Decimal>("AFSCAMT"),
        CLIENTID = x.Field<string>("ARCLID"),
        trantype = x.Field<string>("TransactionTypeCode")

    })
    .GroupBy(x => new { x.room, x.listdate })
    .Select(g => new
    {

        rm = g.Key.room,
        dt = g.Key.listdate,
        Sum = g.Sum(r => r.Commission == 0 && dps.Contains(r.trantype) ? r.afac : 0)


    });

            foreach (var i in placementcnt)
            {
                pcnt.Add(i.rm.Trim() + "." + i.dt.Remove(i.dt.Length - 4, 2), i.count);

            }
            foreach (var i in amtsry)
            {


                amtsum.Add(i.rm.Trim() + "." + i.dt.Remove(i.dt.Length - 4, 2), i.Sum);

            }

            foreach (var i in mamtsry)
            {

                mamtsum.Add(i.rm.Trim() + "." + i.dt.Remove(i.dt.Length - 4, 2), i.Sum);

            }

            DataTable allroom = new DataTable();
            DataColumn dc = new DataColumn("room", typeof(String));
            allroom.Columns.Add(dc);
            foreach (DataRow r in plc.Rows)
            {
                string mm = r["M"].ToString();
                string y = r["Y"].ToString();

                DateTime lastmt = DateTime.Today.AddMonths(-1);
                DateTime tod = DateTime.Today;

                if ((mm == lastmt.ToString("MMM") && y == lastmt.ToString("yyyy").Substring(lastmt.ToString("yyyy").Length - 2, 2)) || (mm == tod.ToString("MMM") && y == tod.ToString("yyyy").Substring(tod.ToString("yyyy").Length - 2, 2)))
                {

                    lstmthrooms.Add(r["ZZRCRM"].ToString());

                }

                DataRow[] rd = allroom.Select("room = '" + r["ZZRCRM"].ToString() + "'");

                if (!(rd != null && rd.Length > 0))
                {

                    DataRow dr = allroom.NewRow();
                    dr[0] = r["ZZRCRM"].ToString();
                    allroom.Rows.Add(dr);
                }


            }

            var mt = allroom.AsEnumerable()
    .Select(x => new
    {
        room = x.Field<string>("room")

    }).OrderBy(o => o.room).Select(g => new
    {

        room = g.room

    });

            foreach (DataRow r in lmtd.Rows)
            {

                if (Convert.ToDecimal(r["AFSCAMT"].ToString()) > 0)

                    rate = r["Comm_Rate"] == DBNull.Value || r["Comm_Rate"] == null || r["Comm_Rate"].ToString() == "" ? 0 : (decimal)r["Comm_Rate"];
                trantype = r["TransactionTypeCode"] == DBNull.Value || r["TransactionTypeCode"] == null || r["TransactionTypeCode"].ToString() == "" ? "" : r["TransactionTypeCode"].ToString();
                DateTime et = (DateTime)r["AFTRENTDTE"];
                DateTime dt = (DateTime)r["ARACLSTDTE"];

                string key = r["ZZRCRM"].ToString() + "." + string.Format("{0}-{1}", dt.ToString("MMM"), dt.ToString("yyyy").Substring(dt.ToString("yyyy").Length - 2, 2));


                lmtdData.Add(key);
                allrooms.Add(r["ZZRCRM"].ToString());


            }

            pagency35.Clear();
            pagency55.Clear();
            pbrad35.Clear();
            pbrad55.Clear();


            // for report A group by listdate and room
            pagency2 = lmtd.AsEnumerable()
    .Select(x => new
    {
        room = x.Field<string>("ZZRCRM"),
        listdate = x.Field<DateTime>("ARACLSTDTE").ToString("MMM-yyyy"),
        Commission = x.Field<Decimal>("Commission"),
        AFTRAMT = x.Field<Decimal>("AFTRAMT"),
        AFSCAMT = x.Field<Decimal>("AFSCAMT"),
        CLIENTID = x.Field<string>("ARCLID"),
        trantype = x.Field<string>("TransactionTypeCode"),
        commrate = x.Field<Decimal>("Comm_Rate")

    })
    .GroupBy(x => new { x.room, x.listdate })
    .Select(g => new
    {

        rm = g.Key.room,
        dt = g.Key.listdate,
        sum = g.Sum(x => x.commrate >= 10.0m &&
        x.commrate <= 53.0m && x.trantype != "DP" && x.trantype != "DPCOR" ? x.AFSCAMT : 0)


    });

            foreach (var i in pagency2)
            {
                pagency35.Add(i.rm.Trim() + "." + i.dt.Remove(i.dt.Length - 4, 2), i.sum);

            }

            pagency3 = lmtd.AsEnumerable()
    .Select(x => new
    {
        room = x.Field<string>("ZZRCRM"),
        listdate = x.Field<DateTime>("ARACLSTDTE").ToString("MMM-yyyy"),
        Commission = x.Field<Decimal>("Commission"),
        AFTRAMT = x.Field<Decimal>("AFTRAMT"),
        AFSCAMT = x.Field<Decimal>("AFSCAMT"),
        CLIENTID = x.Field<string>("ARCLID"),
        trantype = x.Field<string>("TransactionTypeCode"),
        commrate = x.Field<Decimal>("Comm_Rate")

    })
    .GroupBy(x => new { x.room, x.listdate })
    .Select(g => new
    {

        rm = g.Key.room,
        dt = g.Key.listdate,
        sum = g.Sum(x => x.commrate >= 54.8m &&
        x.commrate <= 55.5m && x.trantype != "DP" && x.trantype != "DPCOR" ? x.AFSCAMT : 0)


    });

            foreach (var i in pagency3)
            {
                pagency55.Add(i.rm.Trim() + "." + i.dt.Remove(i.dt.Length - 4, 2), i.sum);

            }

            //paid to bradford
            pbrad2 = lmtd.AsEnumerable()
    .Select(x => new
    {
        room = x.Field<string>("ZZRCRM"),
        listdate = x.Field<DateTime>("ARACLSTDTE").ToString("MMM-yyyy"),
        Commission = x.Field<Decimal>("Commission"),
        AFTRAMT = x.Field<Decimal>("AFTRAMT"),
        AFSCAMT = x.Field<Decimal>("AFSCAMT"),
        CLIENTID = x.Field<string>("ARCLID"),
        trantype = x.Field<string>("TransactionTypeCode"),
        commrate = x.Field<Decimal>("Comm_Rate")

    })
    .GroupBy(x => new { x.room, x.listdate })
    .Select(g => new
    {

        rm = g.Key.room,
        dt = g.Key.listdate,
        sum = g.Sum(x => x.commrate >= 10.0m &&
        x.commrate <= 53.0m && dps.Contains(x.trantype) && collectionphase.Contains(x.CLIENTID) ? x.AFSCAMT : 0)


    });


            foreach (var i in pbrad2)
            {
                pbrad35.Add(i.rm.Trim() + "." + i.dt.Remove(i.dt.Length - 4, 2), i.sum);

            }

            pbrad3 = lmtd.AsEnumerable()
    .Select(x => new
    {
        room = x.Field<string>("ZZRCRM"),
        listdate = x.Field<DateTime>("ARACLSTDTE").ToString("MMM-yyyy"),
        Commission = x.Field<Decimal>("Commission"),
        AFTRAMT = x.Field<Decimal>("AFTRAMT"),
        AFSCAMT = x.Field<Decimal>("AFSCAMT"),
        CLIENTID = x.Field<string>("ARCLID"),
        trantype = x.Field<string>("TransactionTypeCode"),
        commrate = x.Field<Decimal>("Comm_Rate")

    })
    .GroupBy(x => new { x.room, x.listdate })
    .Select(g => new
    {

        rm = g.Key.room,
        dt = g.Key.listdate,
        sum = g.Sum(x => x.commrate >= 54.8m &&
        x.commrate <= 55.5m && dps.Contains(x.trantype) && collectionphase.Contains(x.CLIENTID) ? x.AFSCAMT : 0)


    });


            foreach (var i in pbrad3)
            {
                pbrad55.Add(i.rm.Trim() + "." + i.dt.Remove(i.dt.Length - 4, 2), i.sum);

            }


            //Report - MTD

            Microsoft.Office.Interop.Excel.Worksheet oSheet5 = oWB.Sheets.Add(missing, missing, 1, missing)
            as Microsoft.Office.Interop.Excel.Worksheet;
            oSheet5.Name = "Report B - MTD";

            oSheet5.Cells[1, 1] = "Report B - Month To Date Receipts By Division: ";
            oSheet5.Cells[1, 2] = dlabel;

            rate = 0.0m;
            comovg = 0.0m;
            sumcomovg = 0.0m;
            tot = 4;
            c = "";
            j = "";

            string m = "";
            string rh = "";

            int Accts = 0, tAccts = 0;
            decimal Amt = 0, Letter = 0, pec1 = 0.00m, pec135 = 0, perc235 = 0, pec155 = 0, pec255 = 0, Total = 0, pecPaid = 0,
                trPtyLetter = 0, Per = 0, Contingency = 0, Total2 = 0, Receipts = 0, pec2 = 0.00m,
                tAmt = 0, tLetter = 0, tpec1 = 0, tpec135 = 0, tperc235 = 0, tpec155 = 0, tpec255 = 0, tTotal = 0, tpecPaid = 0,
                ttrPtyLetter = 0, tContingency = 0, tTotal2 = 0, tReceipts = 0, tpec2 = 0;


            foreach (var rr in mt)
            {
                Accts = 0; tAccts = 0;
                Amt = 0; Letter = 0; pec1 = 0; pec135 = 0; perc235 = 0; pec155 = 0; pec255 = 0; Total = 0; pecPaid = 0;
                trPtyLetter = 0; Per = 0; Contingency = 0; Total2 = 0; Receipts = 0; pec2 = 0;
                tAmt = 0; tLetter = 0; tpec1 = 0; tpec135 = 0; tperc235 = 0; tpec155 = 0; tpec255 = 0; tTotal = 0; tpecPaid = 0;
                ttrPtyLetter = 0; tContingency = 0; tTotal2 = 0; tReceipts = 0; tpec2 = 0;

                string r = rr.room;

                oSheet5.Cells[tot, 1] = "Division - " + r;
                oSheet5.Cells.Range["A" + tot.ToString(), "A" + tot.ToString()].Font.Bold = true;
                oSheet5.Cells.Range["A" + tot.ToString(), "A" + tot.ToString()].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbSkyBlue;

                Microsoft.Office.Interop.Excel.Range cell11 = oSheet5.Cells.Range["A" + tot.ToString(), "A" + tot.ToString()].Cells;

                Microsoft.Office.Interop.Excel.Borders border11 = cell11.Borders;
                border11.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                border11.Weight = 2d;

                tot++;
                tot++;

                try
                {
                    oSheet5.Range["A" + tot.ToString(), "Q" + tot.ToString()].Value = new string[,] { { ""
                        , "Placements"
                        ,""
                        ,"3rd Pty"
                        ,""
                        ,"Paid BGE"
                        ,"Paid Agency"
                        ,"Paid BGE"
                        ,"Paid Agency"
                        ,""
                        ,""
                        ,""
                        ,""
                        ,"Fees"
                        ,""
                        ,""
                        ,"Net" }};


                    oSheet5.Cells.Range["A" + tot.ToString(), "Q" + tot.ToString()].Font.Bold = true;
                    oSheet5.Cells.Range["A" + tot.ToString(), "Q" + tot.ToString()].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbSkyBlue;

                    Microsoft.Office.Interop.Excel.Range cell12 = oSheet5.Cells.Range["A" + tot.ToString(), "Q" + tot.ToString()].Cells;

                    Microsoft.Office.Interop.Excel.Borders border12 = cell12.Borders;
                    border12.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    border12.Weight = 2d;


                    tot++;

                    oSheet5.Range["A" + tot.ToString(), "Q" + tot.ToString()].Value = new string[,] { { "Month"
                        , "Accts"
                        ,"Amt"
                        ,"Letter"
                        ,"%"
                        ,"35% Cont."
                        ,"35% Cont."
                        ,"55% Cont."
                        ,"55% Cont."
                        ,"Total"
                        ,"% Paid"
                        ,"3rd Pty Letter"
                        ,"Per"
                        ,"Contingency"
                        ,"Total"
                        ,"Receipts"
                        ,"%" }};


                    oSheet5.Cells.Range["A" + tot.ToString(), "Q" + tot.ToString()].Font.Bold = true;
                    oSheet5.Cells.Range["A" + tot.ToString(), "Q" + tot.ToString()].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbSkyBlue;

                    Microsoft.Office.Interop.Excel.Range cell13 = oSheet5.Cells.Range["A" + tot.ToString(), "Q" + tot.ToString()].Cells;

                    Microsoft.Office.Interop.Excel.Borders border13 = cell13.Borders;
                    border13.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    border13.Weight = 2d;

                }
                catch (Exception ex)
                {

                    Console.WriteLine("@@BLOCK-8-REPORT-B@@" + ex.Message + "@@" + ex.StackTrace);
                }

                tot++;

                try
                {
                    foreach (DataRow rw in plc.Rows)
                    {
                        Accts = 0;
                        Amt = 0; Letter = 0; pec1 = 0; pec135 = 0; perc235 = 0; pec155 = 0; pec255 = 0; Total = 0; pecPaid = 0;
                        trPtyLetter = 0; Per = 0; Contingency = 0; Total2 = 0; Receipts = 0; pec2 = 0;

                        m = rw["M"].ToString() + "-" + rw["Y"].ToString();
                        rh = rw["ZZRCRM"].ToString();

                        string rs = rh + "." + m;


                        if (rh == r)
                        {


                            Accts = Convert.ToInt32(rw["CNT"]);

                            Amt = (decimal)rw["BAL"];

                            if (mamtsum.ContainsKey(rs))
                                Letter = Math.Abs(mamtsum[rs]);

                            if (mpbrad35.ContainsKey(rs))
                                pec135 = mpbrad35[rs];

                            if (mpagency35.ContainsKey(rs))
                                perc235 = mpagency35[rs];

                            if (mpbrad55.ContainsKey(rs))
                                pec155 = mpbrad55[rs];

                            if (mpagency55.ContainsKey(rs))
                                pec255 = mpagency55[rs];

                            Total = Letter + pec135 + pec255;


                            if (pcnt.ContainsKey(rs))
                                trPtyLetter = pcnt[rs] * 0.50m;

                            Per = 0.50m;
                            Total2 = trPtyLetter;

                            Receipts = Total - Total2;

                            if (Amt > 0)
                            {
                                pec1 = Math.Round(Letter / Amt, 4);
                                pec2 = Math.Round(Receipts / Amt, 4);
                                pecPaid = Math.Round(Total / Amt, 4);
                            }
                            else
                            {
                                pec2 = 0;
                                pecPaid = 0;
                                pec1 = 0;
                            }

                            oSheet5.Range["A" + tot.ToString(), "Q" + tot.ToString()].Value = new Object[,] { { "'"+m
                        , Accts
                        ,Amt
                        ,Math.Abs(Letter)
                        ,pec1
                        ,pec135
                        ,perc235
                        ,pec155
                        ,pec255
                        ,Total
                        ,pecPaid
                        ,trPtyLetter
                        ,Per
                        ,""
                        ,Total2
                        ,Receipts
                        ,pec2 }};

                            formatLTMTD(oSheet5, tot);


                            tAmt += Amt;
                            tLetter += Letter;
                            tpec1 += pec1;
                            tpec135 += pec135;
                            tperc235 += perc235;
                            tpec155 += pec155;
                            tpec255 += pec255;
                            tTotal += Total;
                            tpecPaid += pecPaid;
                            ttrPtyLetter += trPtyLetter;
                            tContingency += Contingency;
                            tTotal2 += Total2;
                            tReceipts += Receipts;
                            tpec2 += pec2;
                            tAccts += Accts;

                            tot++;

                        }



                    }

                    if (tAmt > 0)
                    {
                        tpec1 = Math.Round(Math.Abs(tLetter) / tAmt, 4);
                        tpec2 = Math.Round(tReceipts / tAmt, 4);
                        tpecPaid = Math.Round(Math.Abs(tTotal) / tAmt, 4);
                    }
                    else
                    {
                        tpec1 = 0;
                        tpec2 = 0;
                        tpecPaid = 0;
                    }

                    oSheet5.Range["A" + tot.ToString(), "Q" + tot.ToString()].Value = new Object[,] { { "Totals"
                        , tAccts
                        ,tAmt
                        ,Math.Abs(tLetter)
                        ,tpec1
                        ,tpec135
                        ,tperc235
                        ,tpec155
                        ,tpec255
                        ,tTotal
                        ,tpecPaid
                        ,ttrPtyLetter
                        ,""
                        ,""
                        ,tTotal2
                        ,tReceipts
                        ,tpec2 }};


                    oSheet5.Cells.Range["A" + tot.ToString(), "Q" + tot.ToString()].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbLime;
                    Microsoft.Office.Interop.Excel.Range cell4 = oSheet5.Cells.Range["A" + tot.ToString(), "Q" + tot.ToString()].Cells;
                    Microsoft.Office.Interop.Excel.Borders border4 = cell4.Borders;

                    border4.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    border4.Weight = 2d;

                    formatLTMTD(oSheet5, tot);

                }
                catch (Exception ex)
                {

                    Console.WriteLine("@@BLOCK-9-REPORT-B@@" + ex.Message + "@@" + ex.StackTrace);
                }

                tot += 3;


            }



            //Report - LTD

            Microsoft.Office.Interop.Excel.Worksheet oSheet6 = oWB.Sheets.Add(missing, missing, 1, missing)
            as Microsoft.Office.Interop.Excel.Worksheet;
            oSheet6.Name = "Report A - LTD";

            oSheet6.Cells[1, 1] = "Report A - Life To Date Receipts By Division: ";
            oSheet6.Cells[1, 2] = dlabel;

            rate = 0.0m;
            comovg = 0.0m;
            sumcomovg = 0.0m;
            tot = 4;
            c = "";
            j = "";

            m = "";
            rh = "";



            try
            {
                foreach (var rr in mt)
                {
                    Accts = 0; tAccts = 0;
                    Amt = 0; Letter = 0; pec1 = 0; pec135 = 0; perc235 = 0; pec155 = 0; pec255 = 0; Total = 0; pecPaid = 0;
                    trPtyLetter = 0; Per = 0; Contingency = 0; Total2 = 0; Receipts = 0; pec2 = 0;
                    tAmt = 0; tLetter = 0; tpec1 = 0; tpec135 = 0; tperc235 = 0; tpec155 = 0; tpec255 = 0; tTotal = 0; tpecPaid = 0;
                    ttrPtyLetter = 0; tContingency = 0; tTotal2 = 0; tReceipts = 0; tpec2 = 0;

                    string r = rr.room;

                    oSheet6.Cells[tot, 1] = "Division - " + r;
                    oSheet6.Cells.Range["A" + tot.ToString(), "A" + tot.ToString()].Font.Bold = true;
                    oSheet6.Cells.Range["A" + tot.ToString(), "A" + tot.ToString()].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbSkyBlue;

                    Microsoft.Office.Interop.Excel.Range cell11 = oSheet6.Cells.Range["A" + tot.ToString(), "A" + tot.ToString()].Cells;

                    Microsoft.Office.Interop.Excel.Borders border11 = cell11.Borders;
                    border11.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    border11.Weight = 2d;

                    tot++;
                    tot++;

                    oSheet6.Range["A" + tot.ToString(), "Q" + tot.ToString()].Value = new string[,] { { ""
                        , "Placements"
                        ,""
                        ,"3rd Pty"
                        ,""
                        ,"Paid BGE"
                        ,"Paid Agency"
                        ,"Paid BGE"
                        ,"Paid Agency"
                        ,""
                        ,""
                        ,""
                        ,""
                        ,"Fees"
                        ,""
                        ,""
                        ,"Net" }};


                    oSheet6.Cells.Range["A" + tot.ToString(), "Q" + tot.ToString()].Font.Bold = true;
                    oSheet6.Cells.Range["A" + tot.ToString(), "Q" + tot.ToString()].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbSkyBlue;

                    Microsoft.Office.Interop.Excel.Range cell12 = oSheet6.Cells.Range["A" + tot.ToString(), "Q" + tot.ToString()].Cells;

                    Microsoft.Office.Interop.Excel.Borders border12 = cell12.Borders;
                    border12.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    border12.Weight = 2d;


                    tot++;

                    oSheet6.Range["A" + tot.ToString(), "Q" + tot.ToString()].Value = new string[,] { { "Month"
                        , "Accts"
                        ,"Amt"
                        ,"Letter"
                        ,"%"
                        ,"35% Cont."
                        ,"35% Cont."
                        ,"55% Cont."
                        ,"55% Cont."
                        ,"Total"
                        ,"% Paid"
                        ,"3rd Pty Letter"
                        ,"Per"
                        ,"Contingency"
                        ,"Total"
                        ,"Receipts"
                        ,"%" }};


                    oSheet6.Cells.Range["A" + tot.ToString(), "Q" + tot.ToString()].Font.Bold = true;
                    oSheet6.Cells.Range["A" + tot.ToString(), "Q" + tot.ToString()].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbSkyBlue;

                    Microsoft.Office.Interop.Excel.Range cell13 = oSheet6.Cells.Range["A" + tot.ToString(), "Q" + tot.ToString()].Cells;

                    Microsoft.Office.Interop.Excel.Borders border13 = cell13.Borders;
                    border13.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    border13.Weight = 2d;

                    tot++;

                    foreach (DataRow rw in plc.Rows)
                    {
                        Accts = 0;
                        Amt = 0; Letter = 0; pec1 = 0; pec135 = 0; perc235 = 0; pec155 = 0; pec255 = 0; Total = 0; pecPaid = 0;
                        trPtyLetter = 0; Per = 0; Contingency = 0; Total2 = 0; Receipts = 0; pec2 = 0;


                        m = rw["M"].ToString() + "-" + rw["Y"].ToString();
                        rh = rw["ZZRCRM"].ToString();

                        string rs = rh + "." + m;

                        if (rh == r)
                        {

                            DataRow[] am = ABlcnt.Select("ZZRCRM = '" + rh + "' AND M = '" + m.Split('-')[0] + "' AND Y=" + m.Split('-')[1]);

                            Accts = Convert.ToInt32(rw["CNT"]);

                            Amt = (decimal)rw["BAL"];

                            Letter = amtsum.ContainsKey(rs) ? Math.Abs(amtsum[rs]) : 0;

                            if (pbrad35.ContainsKey(rs))
                                pec135 = pbrad35[rs];
                            else
                                pec135 = pbrad35.ContainsKey(rs) ? pbrad35[rs] : 0;

                            if (pagency35.ContainsKey(rs))
                                perc235 = pagency35[rs];
                            else
                                perc235 = pagency35.ContainsKey(rs) ? pagency35[rs] : 0;

                            if (pbrad55.ContainsKey(rs))
                                pec155 = pbrad55[rs];
                            else
                                pec155 = pbrad55.ContainsKey(rs) ? pbrad55[rs] : 0;

                            if (pagency55.ContainsKey(rs))
                                pec255 = pagency55[rs];
                            else
                                pec255 = pagency55.ContainsKey(rs) ? pagency55[rs] : 0;

                            Total = Letter + pec135 + pec255;

                            if (am != null && am.Length > 0)
                                trPtyLetter = Convert.ToInt32(am[0]["CNT"]) * 0.50m;

                            Per = 0.50m;
                            Total2 = trPtyLetter;

                            Receipts = Total - Total2;

                            if (Amt > 0)
                            {
                                pec1 = Math.Round(Letter / Amt, 4);
                                pec2 = Math.Round(Receipts / Amt, 4);
                                pecPaid = Math.Round(Total / Amt, 4);
                            }
                            else
                            {
                                pec2 = 0;
                                pecPaid = 0;
                                pec1 = 0;
                            }

                            oSheet6.Range["A" + tot.ToString(), "Q" + tot.ToString()].Value = new Object[,] { { "'"+m
                        , Accts
                        ,Amt
                        ,Letter
                        ,pec1
                        ,pec135
                        ,perc235
                        ,pec155
                        ,pec255
                        ,Total
                        ,pecPaid
                        ,trPtyLetter
                        ,Per
                        ,""
                        ,Total2
                        ,Receipts
                        ,pec2 }};

                            formatLTMTD(oSheet6, tot);

                            tAmt += Amt;
                            tLetter += Letter;
                            tpec1 += pec1;
                            tpec135 += pec135;
                            tperc235 += perc235;
                            tpec155 += pec155;
                            tpec255 += pec255;
                            tTotal += Total;
                            tpecPaid += pecPaid;
                            ttrPtyLetter += trPtyLetter;
                            tContingency += Contingency;
                            tTotal2 += Total2;
                            tReceipts += Receipts;
                            tpec2 += pec2;
                            tAccts += Accts;

                            tot++;

                        }



                    }

                    if (tAmt > 0)
                    {
                        tpec1 = Math.Round(Math.Abs(tLetter) / tAmt, 4);
                        tpec2 = Math.Round(tReceipts / tAmt, 4);
                        tpecPaid = Math.Round(Math.Abs(tTotal) / tAmt, 4);
                    }
                    else
                    {
                        tpec1 = 0;
                        tpec2 = 0;
                        tpecPaid = 0;
                    }


                    oSheet6.Range["A" + tot.ToString(), "Q" + tot.ToString()].Value = new Object[,] { { "Totals"
                        , tAccts
                        ,tAmt
                        ,Math.Abs(tLetter)
                        ,tpec1
                        ,tpec135
                        ,tperc235
                        ,tpec155
                        ,tpec255
                        ,tTotal
                        ,tpecPaid
                        ,ttrPtyLetter
                        ,""
                        ,""
                        ,tTotal2
                        ,tReceipts
                        ,tpec2 }};



                    oSheet6.Cells.Range["A" + tot.ToString(), "Q" + tot.ToString()].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbLime;
                    Microsoft.Office.Interop.Excel.Range cell4 = oSheet6.Cells.Range["A" + tot.ToString(), "Q" + tot.ToString()].Cells;
                    Microsoft.Office.Interop.Excel.Borders border4 = cell4.Borders;

                    border4.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    border4.Weight = 2d;

                    formatLTMTD(oSheet6, tot);

                    tot += 3;


                }
            }
            catch (Exception ex)
            {

                Console.WriteLine("@@BLOCK-10-REPORT-A@@" + ex.Message + "@@" + ex.StackTrace);
            }


            //styling

            oSheet.Columns.AutoFit();
            oSheet.Cells.Range["A1", "B1"].Font.Bold = true;
            oSheet.Cells.Range["A1", "B1"].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbSkyBlue;
            oSheet.Cells.Range["A2"].Font.Bold = true;

            oSheet.Cells.Range["A4", "J4"].Font.Bold = true;
            oSheet.Cells.Range["A4", "J4"].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbSkyBlue;

            Microsoft.Office.Interop.Excel.Range cell = oSheet.Cells.Range["A1", "B1"].Cells;
            Microsoft.Office.Interop.Excel.Range cell1 = oSheet.Cells.Range["A4", "J4"].Cells;
            Microsoft.Office.Interop.Excel.Borders border = cell.Borders;
            Microsoft.Office.Interop.Excel.Borders border1 = cell1.Borders;

            border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            border.Weight = 2d;

            border1.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            border1.Weight = 2d;

            oSheet.Cells.Font.Size = 12;
            oSheet.Cells.RowHeight = 17;


            try
            {
                oSheet2.Columns.AutoFit();
                oSheet2.Cells.Range["A1", "B1"].Font.Bold = true;
                oSheet2.Cells.Range["A1", "B1"].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbSkyBlue;

                oSheet2.Cells.Range["A4", "N4"].Font.Bold = true;
                oSheet2.Cells.Range["A4", "N4"].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbSkyBlue;

                Microsoft.Office.Interop.Excel.Range cell2 = oSheet2.Cells.Range["A1", "B1"].Cells;
                Microsoft.Office.Interop.Excel.Range cell3 = oSheet2.Cells.Range["A4", "N4"].Cells;
                Microsoft.Office.Interop.Excel.Borders border2 = cell2.Borders;
                Microsoft.Office.Interop.Excel.Borders border3 = cell3.Borders;

                border2.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                border2.Weight = 2d;

                border3.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                border3.Weight = 2d;

                oSheet2.Cells.Font.Size = 12;
                oSheet2.Cells.RowHeight = 17;

                oSheet3.Columns.AutoFit();
                oSheet3.Cells.Range["A1", "B1"].Font.Bold = true;
                oSheet3.Cells.Range["A1", "B1"].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbSkyBlue;

                oSheet3.Cells.Range["A4", "K4"].Font.Bold = true;
                oSheet3.Cells.Range["A4", "K4"].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbSkyBlue;

                Microsoft.Office.Interop.Excel.Range cell5 = oSheet3.Cells.Range["A1", "B1"].Cells;
                Microsoft.Office.Interop.Excel.Range cell6 = oSheet3.Cells.Range["A4", "K4"].Cells;
                Microsoft.Office.Interop.Excel.Borders border5 = cell5.Borders;
                Microsoft.Office.Interop.Excel.Borders border6 = cell6.Borders;

                border5.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                border5.Weight = 2d;

                border6.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                border6.Weight = 2d;

                oSheet3.Cells.Font.Size = 12;
                oSheet3.Cells.RowHeight = 17;

                oSheet4.Columns.AutoFit();
                oSheet4.Cells.Range["A1", "B1"].Font.Bold = true;
                oSheet4.Cells.Range["A1", "B1"].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbSkyBlue;

                oSheet4.Cells.Range["A4", "J5"].Font.Bold = true;
                oSheet4.Cells.Range["A4", "I5"].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbSkyBlue;
                oSheet4.Cells.Range["J5"].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellowGreen;
                oSheet4.Cells.Range["J4"].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellowGreen;

                Microsoft.Office.Interop.Excel.Range cell7 = oSheet4.Cells.Range["A1", "B1"].Cells;
                Microsoft.Office.Interop.Excel.Range cell8 = oSheet4.Cells.Range["A4", "J5"].Cells;
                Microsoft.Office.Interop.Excel.Borders border7 = cell7.Borders;
                Microsoft.Office.Interop.Excel.Borders border8 = cell8.Borders;

                border7.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                border7.Weight = 2d;

                border8.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                border8.Weight = 2d;

                oSheet4.Cells.Font.Size = 12;
                oSheet4.Cells.RowHeight = 17;


                oSheet5.Columns.AutoFit();
                oSheet5.Cells.Range["A1", "B1"].Font.Bold = true;
                oSheet5.Cells.Range["A1", "B1"].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbSkyBlue;
                Microsoft.Office.Interop.Excel.Range cell9 = oSheet5.Cells.Range["A1", "B1"].Cells;

                Microsoft.Office.Interop.Excel.Borders border9 = cell9.Borders;


                border9.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                border9.Weight = 2d;
                oSheet5.Cells.Font.Size = 12;
                oSheet5.Cells.RowHeight = 17;


                oSheet6.Columns.AutoFit();
                oSheet6.Cells.Range["A1", "B1"].Font.Bold = true;
                oSheet6.Cells.Range["A1", "B1"].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbSkyBlue;
                Microsoft.Office.Interop.Excel.Range cell10 = oSheet6.Cells.Range["A1", "B1"].Cells;

                Microsoft.Office.Interop.Excel.Borders border10 = cell10.Borders;
                border10.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                border10.Weight = 2d;
                oSheet6.Cells.Font.Size = 12;
                oSheet6.Cells.RowHeight = 17;

            }
            catch (Exception e)
            {
                
                Console.WriteLine("@@BLOCK-11-REPORT-B@@" + e.Message + "@@" + e.StackTrace);
            }

            oXL.Columns.AutoFit();

            string fp = "BradfordPortfolioObReport_" + DateTime.Today.ToString("MMddyyyy");

            string fileName = _fileLocation + fp + ".xlsx";
            oWB.SaveAs(fileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook,
                missing, missing, missing, missing,
                Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                missing, missing, missing, missing, missing);
            oWB.Close(missing, missing, missing);
            oXL.UserControl = true;
            oXL.Quit();
            
            //Compress(fileName, _fileLocation);





        }


        public static void Compress(string f, string directoryPath)
        {
            FileInfo fileToCompress = new FileInfo(f);


            using (FileStream originalFileStream = fileToCompress.OpenRead())
            {
                if ((File.GetAttributes(fileToCompress.FullName) &
                   FileAttributes.Hidden) != FileAttributes.Hidden & fileToCompress.Extension != ".gz")
                {
                    using (FileStream compressedFileStream = File.Create(fileToCompress.FullName + ".gz"))
                    {
                        using (GZipStream compressionStream = new GZipStream(compressedFileStream,
                           CompressionMode.Compress))
                        {
                            originalFileStream.CopyTo(compressionStream);
                        }
                    }
                    FileInfo info = new FileInfo(directoryPath + Path.DirectorySeparatorChar + fileToCompress.Name + ".gz");
                    Console.WriteLine($"Compressed {fileToCompress.Name} from {fileToCompress.Length.ToString()} to {info.Length.ToString()} bytes.");
                }
            }

        }
        void formatLTMTD(Microsoft.Office.Interop.Excel.Worksheet osheet, int row)
        {

            Microsoft.Office.Interop.Excel.Range vF8 = osheet.Range["C" + row.ToString(), "D" + row.ToString()];


            vF8.NumberFormat = "$#,##0.00";

            Microsoft.Office.Interop.Excel.Range vF9 = osheet.Range["F" + row.ToString(), "J" + row.ToString()];


            vF9.NumberFormat = "$#,##0.00";

            Microsoft.Office.Interop.Excel.Range vF10 = osheet.Range["L" + row.ToString(), "O" + row.ToString()];


            vF10.NumberFormat = "$#,##0.00";

            osheet.Range["P" + row.ToString(), "P" + row.ToString()].NumberFormat = "$#,##0.00;$(#,##0.00)";

            Microsoft.Office.Interop.Excel.Range vF11 = osheet.Range["E" + row.ToString(), "E" + row.ToString()];


            vF11.NumberFormat = "0.00%";

            Microsoft.Office.Interop.Excel.Range vF12 = osheet.Range["K" + row.ToString(), "K" + row.ToString()];


            vF12.NumberFormat = "0.00%";

            Microsoft.Office.Interop.Excel.Range vF13 = osheet.Range["Q" + row.ToString(), "Q" + row.ToString()];


            vF13.NumberFormat = "0.00%";
        }

        void formatMsry(Microsoft.Office.Interop.Excel.Worksheet osheet, int row, bool isrow)
        {
            if (isrow)
            {
                Microsoft.Office.Interop.Excel.Range vF7 = osheet.Range["C" + row.ToString(), "J" + row.ToString()];

                vF7.NumberFormat = "$#,##0.00";

            }
            else
            {
                Microsoft.Office.Interop.Excel.Range vF7 = osheet.Range["H" + row.ToString(), "H" + row.ToString()];

                vF7.NumberFormat = "$#,##0.00";
            }

        }

        void formatsheet23(Microsoft.Office.Interop.Excel.Worksheet osheet, int row, bool is2)
        {

            if (is2)
            {

                osheet.Range["G" + row.ToString(), "J" + row.ToString()].NumberFormat = "$#,##0.00";

                osheet.Range["L" + row.ToString(), "L" + row.ToString()].NumberFormat = "$#,##0.00";


            }
            else
            {
                osheet.Range["G" + row.ToString(), "H" + row.ToString()].NumberFormat = "$#,##0.00";

            }

        }



        void formatrDF(Microsoft.Office.Interop.Excel.Worksheet osheet, int row, string s, string e)
        {

            Microsoft.Office.Interop.Excel.Range vF7 = osheet.Range[s + row.ToString(), e + row.ToString()];

            vF7.NumberFormat = "$#,##0.00";



        }



    }


}