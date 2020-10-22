using System;
using System.Collections.Generic;
using System.Text;
using System.Data.SqlClient;

namespace FocasSmartDataCollection
{
    public static class ServiceStop
    {
        public static volatile int stop_service = 0;
        public static void DeleteOnlinemachineList(string machineID)
        {
            SqlConnection conn = ConnectionManager.GetConnection();
            try
            {
                string qry = "delete from onlinemachinelist where machineid='" + machineID + "'";
                SqlCommand cmd = new SqlCommand(qry, conn);
                cmd.ExecuteNonQuery();
                Logger.WriteDebugLog(string.Format("{0} machine has been deleted from onlinemachinelist", machineID));
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.ToString());
            }
            finally
            {
                if (conn != null)
                {
                    conn.Close();                   
                }
            }
        }        
    }
}
