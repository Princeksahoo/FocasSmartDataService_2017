using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Data.Common;
using System.Net;
using System.Net.Mail;
using System.Configuration;
using System.IO;
using System.Reflection;

namespace FocasSmartDataCollection
{
    class BackupDB
    {
        public static string DBName = string.Empty;
        public static string BackupPath = string.Empty;

        public BackupDB()
        {

        }

        public void fun_caller()
        {
            DBBackUpParameter();
            if (!string.IsNullOrEmpty(DBName) && !string.IsNullOrEmpty(BackupPath))
            {
                bool isfirstTimeBackup = false;
                DateTime dt = DBBackUpDate(out isfirstTimeBackup);
                int mindaysbackup = MinDaysForDBbackup();

                if ((DateTime.Now > dt.AddDays(mindaysbackup)) || isfirstTimeBackup)              
                {
                    CheckDBsize(DBName);
                    BackupDatabase(DBName, BackupPath);
                    Logger.WriteDebugLog(string.Format("Next database backup will be done after {0} date...", DateTime.Now.AddDays(mindaysbackup).ToString("dd-MMM-yyyy hh:mm:ss tt")));
                    TruncateDatabaseLog();
                    CheckDBsize(DBName);
                }
            }
        }

        public static void TruncateDatabaseLog()
        {
            string sqlQuery = string.Empty;
            object logFileNameOfDB = null;
            SqlCommand cmd = null;
            SqlConnection conn = ConnectionManager.GetConnection();

            try
            {               
                sqlQuery = "SELECT name FROM dbo.sysfiles WHERE fileid = 2";
                cmd = new SqlCommand(sqlQuery, conn);
                logFileNameOfDB = cmd.ExecuteScalar();               

                if (logFileNameOfDB == null || Convert.IsDBNull(logFileNameOfDB))
                {
                    Logger.WriteDebugLog(string.Format("No log file name return by Sql Query. Query = {0}", sqlQuery));
                    return;
                }
             
                sqlQuery = "DBCC SHRINKFILE(" + logFileNameOfDB.ToString().Trim() + ", 1)";
               
                cmd = new SqlCommand(sqlQuery, conn);
                cmd.CommandTimeout = 60 * 4;
                Logger.WriteDebugLog(string.Format("TRUNCATE of database Log {0} started.", DBName));
                cmd.ExecuteNonQuery();
                Logger.WriteDebugLog(string.Format("Log file {0} has been truncated for the databse {1} ", logFileNameOfDB.ToString().Trim(), DBName));
            }
            catch (Exception ex)
            {
                Logger.WriteDebugLog("Exception : TruncateDatabaseLog() Message = " + ex);
            }
            finally
            {
                if (cmd != null) cmd.Dispose();
                if (conn != null) conn.Close();
            }
        }

        public static bool BackupDatabase(string db, string back_path)
        {
            bool isBackupSuccessfull = true;
            SqlConnection conn = ConnectionManager.GetConnection();
            SqlCommand cmd = null;
            try
            {
                if (!Directory.Exists(back_path))
                {
                    Directory.CreateDirectory(back_path);
                }

                Logger.WriteDebugLog(string.Format("backup of database {0} is started.", db));
                string appPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

                string backupName = string.Format("{0}_{1}.bak", db, DateTime.Now.ToString("yyyyMMdd_HHmmss"));
                string backupFilePath = Path.Combine(back_path, backupName);
                if (File.Exists(backupFilePath))
                {
                    File.Delete(backupFilePath);
                }
               
                string sqlqry = "backup database [" + db + "] to disk ='" + backupFilePath + "'";
                cmd = new SqlCommand(sqlqry, conn);
                cmd.CommandTimeout = 60 * 60; // 10 mins
                cmd.ExecuteNonQuery();
                Logger.WriteDebugLog(string.Format("backup of database {0} is done successfully at {1}. Query = {2}.", db, backupFilePath, sqlqry));
                cmd.Dispose();
              
                /*
                Logger.WriteDebugLog("Zipping the backup of database.");
                List<string> fileToZip = new List<string>();
                fileToZip.Add(backupFilePath);
                string zippedFileName = MoveFilesCls.zipFilesWithFolders(db + "-Database", fileToZip);
                Logger.WriteDebugLog("Zipping the backup of database completed.");
                File.Move(zippedFileName, Path.Combine(back_path, Path.GetFileName(zippedFileName)));              
                Logger.WriteDebugLog(string.Format("Database backup file has been Zipped and moved to {0}", back_path));
                */
            }
            catch (Exception ex)
            {
                isBackupSuccessfull = false;
                Logger.WriteErrorLog("Exception in BackupDatabase() : " + ex.Message);
            }
            finally
            {
                if (cmd != null) cmd.Dispose();
                if (conn != null) conn.Close();
            }
            return isBackupSuccessfull;
        }

        public static void CheckDBsize(string db)
        {
            SqlConnection conn = ConnectionManager.GetConnection();
            string sqlQuery = "SELECT sum((size*8)/1024 ) SizeMB FROM dbo.sysfiles"; // WHERE DB_NAME(database_id) ='" + db + "' ";
            SqlCommand cmd2 = new SqlCommand(sqlQuery, conn);
            object DBsize1 = null;
            try
            {
                DBsize1 = cmd2.ExecuteScalar();
                if (DBsize1 == null || Convert.IsDBNull(DBsize1))
                {
                    Logger.WriteErrorLog(string.Format("Size of the db {0} is null. Check the query {1}", db, sqlQuery));
                }
                else
                {
                    Logger.WriteDebugLog(string.Format("Size of the db {0} is {1} MB.", db, int.Parse(DBsize1.ToString())));
                }

                //if (int.Parse(DBsize1.ToString()) >= 1536)//1536 MB = 1.5 gb
                //{
                    //string Email_flag = ConfigurationManager.AppSettings["Email"].ToString();
                    //if (Email_flag.ToLower() == "true")
                    //{
                    //    SendEmail(db, int.Parse(DBsize1.ToString()));
                    //}
                //}
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog("Exception in CheckDBsize() : Message = " + ex.Message);
            }
            finally
            {
                if (cmd2 != null) cmd2.Dispose();
                if (conn != null) conn.Close();
            }
        }

        public static void DBBackUpParameter()
        {
            SqlConnection conn = ConnectionManager.GetConnection();
            string sqlQuery = "Select * from Focas_Defaults where Parameter='dataBaseBackUpValues'";
            SqlCommand cmd = new SqlCommand(sqlQuery, conn);
            SqlDataReader sqlrdr = null;

            try
            {
                sqlrdr = cmd.ExecuteReader();
                if (sqlrdr.HasRows)
                {
                    sqlrdr.Read();
                    BackupPath = sqlrdr.IsDBNull(1) ? Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) : sqlrdr[1].ToString();
                    DBName = conn.Database;
                }
                else
                {
                    Logger.WriteErrorLog("Parameter 'dataBaseBackUpValues' is not defined in Focas_Defaults. Unable to take Database backup.");
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog("Exception in DBBackUpParameter(); Message = " + ex.Message);
            }
            finally
            {
                if (sqlrdr != null) sqlrdr.Close();
                if (cmd != null) cmd.Dispose();
                if (conn != null) conn.Close();
            }
        }
           

        public static DateTime DBBackUpDate(out bool isFirstTimeBackup , bool writeMessage = false)
        {
            isFirstTimeBackup = false;
            SqlConnection conn = ConnectionManager.GetConnection();
            //string sqlQuery = "SELECT Replace(CONVERT(VARCHAR(10), MAX(bus.backup_finish_date), 120),'/','-') AS LastBackUpTime FROM msdb.dbo.backupset bus  WHERE bus.database_name ='" + conn.Database + "'";
            string sqlQuery = "SELECT MAX(bus.backup_finish_date) AS LastBackUpTime FROM msdb.dbo.backupset bus  WHERE bus.database_name ='" + conn.Database + "'";
            SqlCommand cmd2 = new SqlCommand(sqlQuery, conn);
            object DBBackUP_date = null;
            DateTime dt = DateTime.Now;
            try
            {
                DBBackUP_date = cmd2.ExecuteScalar();
                if (DBBackUP_date == null || Convert.IsDBNull(DBBackUP_date))
                {
                    isFirstTimeBackup = true;
                    Logger.WriteDebugLog(string.Format("First time database backup is taking place on {0}", DateTime.Now.ToString()));
                    return dt;
                }
                dt = (DateTime)DBBackUP_date;
                if (writeMessage == true)
                {
                    Logger.WriteDebugLog(string.Format("Last database backup date is {0}", dt.ToString("dd-MMM-yyyy hh:mm:ss tt")));
                    int mindaysbackup = MinDaysForDBbackup();                   
                    Logger.WriteDebugLog(string.Format("Next database backup will be done after {0} date...", dt.AddDays(mindaysbackup).ToString("dd-MMM-yyyy hh:mm:ss tt")));                   
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog("Exception in DBBackUpDate() : Message = " + ex.Message);
            }
            finally
            {
                if (cmd2 != null) cmd2.Dispose();
                if (conn != null) conn.Close();
            }           

            return dt;

        }
       
        public static int MinDaysForDBbackup()
        {
            SqlConnection conn = ConnectionManager.GetConnection();
            string sqlQuery = "Select ValueinText2 from Focas_Defaults where Parameter='dataBaseBackUpValues'";
            SqlCommand cmd = new SqlCommand(sqlQuery, conn);
            object mindays_DB = 7;
            int DatabaseBackupFrequency = 7;
            try
            {
                mindays_DB = cmd.ExecuteScalar();
                DatabaseBackupFrequency = (mindays_DB == null || Convert.IsDBNull(mindays_DB)) ? 7 : int.Parse(mindays_DB.ToString());
                if (DatabaseBackupFrequency <= 7) DatabaseBackupFrequency = 7;
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog("Exception : " + ex.Message);
            }
            finally
            {
                if (cmd != null) cmd.Dispose();
                if (conn != null) conn.Close();
            }
            return DatabaseBackupFrequency;
        }
    }
}



















































        