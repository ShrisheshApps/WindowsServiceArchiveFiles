using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.ServiceProcess;
using System.Timers;

namespace WindowsServiceArchiveFiles
{
    public partial class ArchiveService : ServiceBase
    {
        string appDomain = AppDomain.CurrentDomain.BaseDirectory + "Logs";
        System.Timers.Timer timer = new System.Timers.Timer();
        private readonly string constr = ConfigurationManager.ConnectionStrings["Conn"].ConnectionString.ToString();

        private void CreateFolder(string folderPath)
        {
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
        }
        public void MoveFiles()
        {
            DataSet ds = new DataSet();
            DataTable tbl;
            string srcFolder = @"C:\Users\dell\Documents\Archivetest\OldFiles\";
            string destFolder = @"C:\Users\dell\Documents\Archivetest\NewFolder\";
            DateTime startDate = GetMinDateModified();
            DateTime endDate = startDate.AddMonths(9);
            // list files having AuditDate between startDate and endDate
            string cmdtxt = "Select Distinct Filename from tblFiles Where AuditDate >= '" + startDate + "' And AuditDate <= '" + endDate + "'";
            using (SqlConnection con = new SqlConnection(constr))
            {
                using (SqlDataAdapter sda = new SqlDataAdapter(cmdtxt, con))
                {
                    sda.Fill(ds);
                    tbl = ds.Tables[0];
                }
            }
            //lookup files in C:\Users\dell\Documents\Archivetest\OldFiles
            foreach (DataRow rw in tbl.Rows)
            {
                var f = rw[0].ToString();
                if (File.Exists(srcFolder + f))
                {
                    File.AppendAllText(appDomain + "\\log.txt", DateTime.Now + "\t" + f + System.Environment.NewLine);
                    // create yearFolder based on file modified date
                    FileInfo fileInfo = new FileInfo(srcFolder + f);
                    string targetFolder = destFolder + fileInfo.LastWriteTime.Year.ToString() ;
                    CreateFolder(targetFolder);  
                    File.Move(srcFolder + f, targetFolder + "\\" + f);
                    try
                    {
                        InsertRecordForMovedFile(srcFolder + f, targetFolder + "\\" + f, f, "");
                    }
                    catch (Exception ex)
                    {
                        InsertRecordForMovedFile(srcFolder + f, targetFolder + "\\" + f, f, ex.Message);
                    }
                }
            }
        }

        private DateTime GetMinDateModified()
        {
            try
            {
                using (SqlConnection con = new SqlConnection(constr))
                {
                    using (SqlCommand cmd = new SqlCommand("Select Min(AuditDate) from tblFiles", con))
                    {
                        cmd.CommandType = CommandType.Text;
                        con.Open();
                        return Convert.ToDateTime(cmd.ExecuteScalar());
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
        public void InsertRecords()
        {
            string folderpath = @"C:\Users\dell\Documents\SomeFiles\";
            using (SqlConnection con = new SqlConnection(constr))
            {
                foreach (var file in Directory.GetFiles(folderpath))
                {
                    FileInfo fileInfo = new FileInfo(file);

                    string cmdTxt = "INSERT INTO tblFiles Values('" + fileInfo.Name + "', '" + fileInfo.LastWriteTime.Date + "')";
                    using (SqlCommand cmd = new SqlCommand(cmdTxt, con))
                    {
                        if (System.Data.ConnectionState.Open == con.State)
                        {
                            //
                        }
                        else
                        {
                            con.Open();
                        }
                        cmd.ExecuteNonQuery();
                    }
                }
            }
        }
        public ArchiveService()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            timer.Interval = 1000; // 120 seconds
            timer.AutoReset = true;
            timer.Elapsed += new System.Timers.ElapsedEventHandler(TimeArchive);
            timer.Enabled = true;
            EventLog.WriteEntry("Archive service started...");
        }
        private void RunArchive()
        {
            CreateFolder(appDomain);
            File.AppendAllText(appDomain + "\\log2.txt", DateTime.Now + System.Environment.NewLine);
            MoveFiles();
        }
        private void TimeArchive(object sender, ElapsedEventArgs e)
        {
            CreateFolder(appDomain);
            File.AppendAllText(appDomain + "\\log2.txt", DateTime.Now + System.Environment.NewLine);
            MoveFiles();
        }

        protected override void OnStop()
        {
            EventLog.WriteEntry("Archive service ended...");
            timer.Enabled = false;
        }

        private void InsertRecordForMovedFile(string source, string destination, string filename, string error)
        {
            try
            {
                using (var con = new SqlConnection(constr))
                {
                    using (SqlCommand cmd = new SqlCommand("usp_MoveFile", con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@SourcePath", SqlDbType.NVarChar).Value = source;
                        cmd.Parameters.Add("@DestPath", SqlDbType.NVarChar).Value = destination;
                        cmd.Parameters.Add("@FileName", SqlDbType.NVarChar).Value = filename;
                        cmd.Parameters.Add("@MovedDate", SqlDbType.Date).Value = DateTime.Today.ToString();
                        cmd.Parameters.Add("@ErrorMessage", SqlDbType.NVarChar).Value = error;
                        con.Open();
                        cmd.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
    }
}
