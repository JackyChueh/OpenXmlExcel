using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;
using System.Data.SqlClient;

class DataAccessLayer
{

    #region DataTable
    /// <summary>
    /// 查詢資料表
    /// </summary>
    /// <param name="SqlStr">SQL語法</param>
    /// <param name="ConnStr">資料庫連線字串</param>
    /// <returns>DataTable</returns>
    public static DataTable SelectDataTable(string SqlStr, string ConnStr)
    {
        try
        {
            SqlConnection SqlConn = new SqlConnection(ConnStr);

            SqlCommand SqlCmd = new SqlCommand(SqlStr, SqlConn);

            return SelectDataTable(SqlCmd, SqlConn);
        }
        catch (Exception ex)
        {
            throw ex;
        }
    }

    /// <summary>
    /// 查詢資料表
    /// </summary>
    /// <param name="SqlStr">SQL語法</param>
    /// <param name="SqlPara">語法參數</param>
    /// <param name="ConnStr">資料庫連線字串</param>
    /// <returns>DataTable</returns>
    public static DataTable SelectDataTable(string SqlStr, Hashtable SqlPara, string ConnStr)
    {
        try
        {
            SqlConnection SqlConn = new SqlConnection(ConnStr);

            SqlCommand SqlCmd = new SqlCommand(SqlStr, SqlConn);

            if (SqlPara != null)
            {
                foreach (string K in SqlPara.Keys)
                    SqlCmd.Parameters.AddWithValue(K, SqlPara[K]);
            }

            return SelectDataTable(SqlCmd, SqlConn);
        }
        catch (Exception ex)
        {
            throw ex;
        }
    }

    /// <summary>
    /// 查詢資料表
    /// </summary>
    /// <param name="SqlStr">SQL語法</param>
    /// <param name="SqlPara">語法參數</param>
    /// <param name="ConnStr">資料庫連線字串</param>
    /// <returns>DataTable</returns>
    public static DataTable SelectDataTable(string SqlStr, List<SqlParameter> SqlPara, string ConnStr)
    {
        try
        {
            SqlConnection SqlConn = new SqlConnection(ConnStr);

            SqlCommand SqlCmd = new SqlCommand(SqlStr, SqlConn);

            if (SqlPara != null && SqlPara.Count > 0)
            {
                SqlCmd.Parameters.AddRange(SqlPara.ToArray<SqlParameter>());
            }

            return SelectDataTable(SqlCmd, SqlConn);
        }
        catch (Exception ex)
        {
            throw ex;
        }
    }

    /// <summary>
    /// 查詢資料表
    /// </summary>
    /// <param name="SqlCmd">SQL命令物件</param>
    /// <param name="SqlConn">資料庫連線物件</param>
    /// <returns>DataTable</returns>
    public static DataTable SelectDataTable(SqlCommand SqlCmd, SqlConnection SqlConn)
    {
        SqlDataAdapter SqlAda = null;
        try
        {
            DataTable dt = new DataTable();

            SqlCmd.Connection = SqlConn;
            SqlAda = new SqlDataAdapter(SqlCmd);
            SqlConn.Open();
            SqlAda.Fill(dt);

            return dt;
        }
        catch (Exception ex)
        {
            throw ex;
        }
        finally
        {
            SqlAda.Dispose();
            SqlCmd.Dispose();
            SqlConn.Close();
            SqlConn.Dispose();
        }
    }


    #endregion

    #region Execute

    public static int ExecuteSql(string SqlStr, string ConnStr, string DBName)
    {
        SqlConnection SqlConn = null;
        SqlCommand SqlCmd = null;
        try
        {
            SqlConn = new SqlConnection(ConnStr);
            SqlCmd = new SqlCommand();
            SqlConn.Open();
            if (!string.IsNullOrEmpty(DBName))
            {
                SqlConn.ChangeDatabase(DBName);
            }
            SqlCmd.Connection = SqlConn;
            SqlCmd.Transaction = SqlConn.BeginTransaction();
            SqlCmd.CommandText = SqlStr;

            int i = SqlCmd.ExecuteNonQuery();
            SqlCmd.Transaction.Commit();
            return i;
        }
        catch (Exception ex)
        {
            SqlCmd.Transaction.Rollback();
            throw new Exception(ex.Message);
        }
        finally
        {
            SqlCmd.Dispose();
            SqlConn.Close();
        }
    }

    public static int ExecuteSql(string SqlStr, List<SqlParameter> SqlPara, string ConnStr)
    {
        SqlConnection SqlConn = null;
        SqlCommand SqlCmd = null;
        try
        {
            SqlConn = new SqlConnection(ConnStr);
            SqlCmd = new SqlCommand();
            SqlConn.Open();

            SqlCmd.Connection = SqlConn;
            SqlCmd.Transaction = SqlConn.BeginTransaction();
            SqlCmd.CommandText = SqlStr;

            if (SqlPara != null)
            {
                SqlCmd.Parameters.AddRange(SqlPara.ToArray<SqlParameter>());
            }

            int i = SqlCmd.ExecuteNonQuery();
            SqlCmd.Transaction.Commit();
            return i;

        }
        catch (Exception ex)
        {
            SqlCmd.Transaction.Rollback();
            throw new Exception(ex.Message);
        }
        finally
        {
            SqlPara = null;
            SqlCmd.Dispose();
            SqlConn.Close();
        }
    }
    public static int ExecuteSql(string SqlStr, List<SqlParameter> SqlPara, string ConnStr, int Timeout)
    {
        SqlConnection SqlConn = null;
        SqlCommand SqlCmd = null;
        try
        {
            SqlConn = new SqlConnection(ConnStr);
            SqlCmd = new SqlCommand();
            SqlConn.Open();

            SqlCmd.Connection = SqlConn;
            SqlCmd.Transaction = SqlConn.BeginTransaction();
            SqlCmd.CommandText = SqlStr;
            SqlCmd.CommandTimeout = Timeout;

            if (SqlPara != null)
            {
                SqlCmd.Parameters.AddRange(SqlPara.ToArray<SqlParameter>());
            }

            int i = SqlCmd.ExecuteNonQuery();
            SqlCmd.Transaction.Commit();
            return i;

        }
        catch (Exception ex)
        {
            SqlCmd.Transaction.Rollback();
            throw new Exception(ex.Message);
        }
        finally
        {
            SqlPara = null;
            SqlCmd.Dispose();
            SqlConn.Close();
        }
    }

    public static int ExecuteSqlCmds(string SqlStr, List<SqlCommand> SqlCmds, string ConnStr)
    {
        int EffectRows = 0;
        SqlConnection SqlConn = null;
        SqlTransaction SqlTran = null;
        try
        {
            SqlConn = new SqlConnection(ConnStr);
            SqlConn.Open();
            SqlTran = SqlConn.BeginTransaction();

            foreach (SqlCommand SqlCmd in SqlCmds)
            {
                SqlCmd.Connection = SqlConn;
                SqlCmd.Transaction = SqlTran;
                int i = SqlCmd.ExecuteNonQuery();
                EffectRows += i;
            }

            SqlTran.Commit();
            return EffectRows;

        }
        catch (Exception ex)
        {
            SqlTran.Rollback();
            throw ex;
        }
        finally
        {
            //SqlPara = null;
            //SqlCmd.Dispose();
        }
    }

    public static DataTable ExecuteSp(string SqlStr, List<SqlParameter> SqlPara, string ConnStr)
    {
        SqlConnection SqlConn = null;
        SqlCommand SqlCmd = null;
        SqlDataReader SqlReader = null;
        try
        {
            SqlConn = new SqlConnection(ConnStr);
            SqlCmd = new SqlCommand();
            SqlConn.Open();

            SqlCmd.Connection = SqlConn;
            SqlCmd.CommandText = SqlStr;
            SqlCmd.CommandType = CommandType.StoredProcedure;

            if (SqlPara != null)
            {
                SqlCmd.Parameters.AddRange(SqlPara.ToArray<SqlParameter>());
            }
            SqlReader = SqlCmd.ExecuteReader();
            //SqlReader.Read();

            DataTable dt = new DataTable();
            dt.Load(SqlReader);

            return dt;

        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
        finally
        {
            SqlReader.Close();
            SqlPara = null;
            SqlCmd.Dispose();
            SqlConn.Close();
        }
    }
    #endregion
}
