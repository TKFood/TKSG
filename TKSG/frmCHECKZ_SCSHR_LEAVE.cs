﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
using NPOI.SS.UserModel;
using System.Configuration;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util;
using System.Reflection;
using System.Threading;
using FastReport;
using FastReport.Data;
using System.Xml;

namespace TKSG
{
    public partial class frmCHECKZ_SCSHR_LEAVE : Form
    {

        String connectionStringTKSG = "server=192.168.1.105;database=TKSG;uid=sa;pwd=dsc";
        String connectionStringTKGAFFAIRS = "server=192.168.1.105;database=TKGAFFAIRS;uid=sa;pwd=dsc";
        String connectionStringUOF = "server=192.168.1.223;database=UOF;uid=TKUOF;pwd=TKUOF123456";

        string DB = "UOFTEST";

        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
        SqlDataAdapter adapter2 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder2 = new SqlCommandBuilder();

        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds1 = new DataSet();
        DataSet ds2 = new DataSet();

        string tablename = null;
        string EDITID;
        int result;
        Thread TD;

        string TaskId;
        string CARDNO;


        //用STATUS來控制在1分鐘內不得連續刷卡
        string STATUS = "Y";

        public frmCHECKZ_SCSHR_LEAVE()
        {
            InitializeComponent();

            label6.Text = DateTime.Now.ToString("yyyy/MM/dd HH:mm");

            timer1.Enabled = true;
            timer1.Interval = 1000 * 30;
            timer1.Start();
        }


        #region FUNCTION
        private void timer1_Tick(object sender, EventArgs e)
        {
            label6.Text = DateTime.Now.ToString("yyyy/MM/dd HH:mm");
            STATUS = "Y";

            if (!string.IsNullOrEmpty(textBox1.Text))
            {
                CARDNO = textBox1.Text;
                textBox1.Text = null;

                SEARCHHREngFrm001B(CARDNO);
                CARDNO = null;
            }


            if (!string.IsNullOrEmpty(textBox2.Text))
            {
                CARDNO = textBox2.Text;
                textBox2.Text = null;

                SEARCHHREngFrm001B(CARDNO);
                CARDNO = null;
            }
        }
        public void ADDZ_SCSHR_LEAVE()
        {
            DataSet Z_SCSHR_LEAV = SEARCHHZ_SCSHR_LEAV();
           
            if(Z_SCSHR_LEAV.Tables[0].Rows.Count>0)
            {
                try
                {
                    connectionString = connectionStringTKGAFFAIRS;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();
                 
                    foreach(DataRow row in Z_SCSHR_LEAV.Tables[0].Rows)
                    {
                        sbSql.AppendFormat(@" 
                                            INSERT INTO [TKGAFFAIRS].[dbo].[Z_SCSHR_LEAVE]
                                            ([DOC_NBR],[TASK_STATUS],[TASK_RESULT],[GROUP_CODE],[APPLICANT],[APPLICANTGUID],[APPLICANTCOMP],[APPLICANTDEPT],[APPLICANTDATE],[LEAEMP],[LEAAGENT],[LEACODE],[LEACODENAME],[SP_DATE],[SP_NAME],[STARTTIME],[ENDTIME],[LEAHOURS],[LEADAYS],[REMARK],[CANCEL_DOC_NBR],[CANCEL_STATUS],[SCSHR],[SCSHRMSG],[CRADNO])
                                            VALUES
                                            ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}','{17}','{18}','{19}','{20}','{21}','{22}','{23}','{24}')

                                            ", row["DOC_NBR"].ToString(), row["TASK_STATUS"].ToString(), row["TASK_RESULT"].ToString(), row["GROUP_CODE"].ToString(), row["APPLICANT"].ToString(), row["APPLICANTGUID"].ToString(), row["APPLICANTCOMP"].ToString(), row["APPLICANTDEPT"].ToString(), row["APPLICANTDATE"].ToString(), row["LEAEMP"].ToString(), row["LEAAGENT"].ToString(), row["LEACODE"].ToString(), row["LEACODENAME"].ToString(), row["SP_DATE"].ToString(), row["SP_NAME"].ToString(), row["STARTTIME"].ToString(), row["ENDTIME"].ToString(), row["LEAHOURS"].ToString(), row["LEADAYS"].ToString(), row["REMARK"].ToString(), row["CANCEL_DOC_NBR"].ToString(), row["CANCEL_STATUS"].ToString(), row["SCSHR"].ToString(), row["SCSHRMSG"].ToString(), row["CardNo"].ToString());
                    }

                    sbSql.AppendFormat(@"   ");

                    cmd.Connection = sqlConn;
                    cmd.CommandTimeout = 60;
                    cmd.CommandText = sbSql.ToString();
                    cmd.Transaction = tran;
                    result = cmd.ExecuteNonQuery();

                    if (result == 0)
                    {
                        tran.Rollback();    //交易取消
                    }
                    else
                    {
                        tran.Commit();      //執行交易  

                        MessageBox.Show("完成");
                    }
                }
                catch
                {

                }

                finally
                {
                    sqlConn.Close();
                }
            }

        }


        public DataSet SEARCHHZ_SCSHR_LEAV()
        {
            DataSet ds = new DataSet();

            try
            {
                connectionString = connectionStringTKSG;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                StringBuilder query = new StringBuilder();

                sbSql.AppendFormat(@" 
                                    SELECT  
                                    [DOC_NBR]
                                    ,[TASK_STATUS]
                                    ,[TASK_RESULT]
                                    ,[GROUP_CODE]
                                    ,[APPLICANT]
                                    ,[APPLICANTGUID]
                                    ,[APPLICANTCOMP]
                                    ,[APPLICANTDEPT]
                                    ,CONVERT(NVARCHAR,[APPLICANTDATE],111) APPLICANTDATE
                                    ,[LEAEMP]
                                    ,[LEAAGENT]
                                    ,[LEACODE]
                                    ,[LEACODENAME]
                                    ,CONVERT(NVARCHAR,[SP_DATE],111) [SP_DATE] 
                                    ,[SP_NAME]
                                    ,CONVERT(NVARCHAR,[STARTTIME],120) [STARTTIME] 
                                    ,CONVERT(NVARCHAR,[ENDTIME],120) [ENDTIME] 
                                    ,[LEAHOURS]
                                    ,[LEADAYS]
                                    ,[REMARK]
                                    ,[CANCEL_DOC_NBR]
                                    ,[CANCEL_STATUS]
                                    ,[SCSHR]
                                    ,[SCSHRMSG]
                                    ,[CardNo]
                                    FROM [192.168.1.223].[{0}].[dbo].[Z_SCSHR_LEAVE]
                                    LEFT JOIN [192.168.1.225].[CHIYU].[dbo].[Person] ON [APPLICANT]=[EmployeeID] COLLATE Chinese_PRC_CI_AS
                                    WHERE [DOC_NBR] COLLATE Chinese_Taiwan_Stroke_BIN NOT IN (SELECT [DOC_NBR] FROM [TKGAFFAIRS].[dbo].[Z_SCSHR_LEAVE]) 
                                    

                                    ", DB);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds.Clear();
                adapter1.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count >= 0)
                {
                    return ds;
                }
                else
                {
                    return null;
                }

               

            }
            catch
            {
                return null;
            }
            finally
            {
                sqlConn.Close();
            }
           
        }

        public void SEARCHHREngFrm001B(string CARDNO)
        {

            try
            {
                connectionString = connectionStringTKGAFFAIRS;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                StringBuilder query = new StringBuilder();

                if (string.IsNullOrEmpty(CARDNO))
                {
                    sbSql.AppendFormat(@"  SELECT ");
                    sbSql.AppendFormat(@"  [HREngFrm001User] AS '申請人',[HREngFrm001Rank] AS '職級',[HREngFrm001OutDate] AS '外出日期',[HREngFrm001Transp] AS '交通工具',[HREngFrm001LicPlate] AS '車牌',[HREngFrm001DefOutTime] AS '預計外出時間',[HREngFrm001OutTime] AS '實際外出時間',[HREngFrm001DefBakTime] AS '預計返廠時間',[HREngFrm001BakTime] AS '實際返廠時間'");
                    sbSql.AppendFormat(@"  ,[TaskId] AS 'TaskId',[HREngFrm001SN] AS '表單編號',[HREngFrm001Date] AS '申請日期',[HREngFrm001UsrDpt] AS '部門',[HREngFrm001Location] AS '外出地點',[HREngFrm001Agent] AS '代理人',[HREngFrm001Cause] AS '外出原因',[HREngFrm001FF] AS '是否由公司出發',[HREngFrm001CH] AS '是否回廠',[CRADNO] AS '卡號'");
                    sbSql.AppendFormat(@"  FROM [TKGAFFAIRS].[dbo].[HREngFrm001]");
                    sbSql.AppendFormat(@"  WHERE ISNULL([HREngFrm001SN],'')<>''");
                    sbSql.AppendFormat(@"  AND [HREngFrm001OutDate]='{0}' ", DateTime.Now.ToString("yyyy/MM/dd"));
                    sbSql.AppendFormat(@"  ORDER BY [HREngFrm001User],[HREngFrm001DefOutTime]");

                    sbSql.AppendFormat(@"  

                                        ");
                }
                else
                {
                    sbSql.AppendFormat(@"  SELECT ");
                    sbSql.AppendFormat(@"  [HREngFrm001User] AS '申請人',[HREngFrm001Rank] AS '職級',[HREngFrm001OutDate] AS '外出日期',[HREngFrm001Transp] AS '交通工具',[HREngFrm001LicPlate] AS '車牌',[HREngFrm001DefOutTime] AS '預計外出時間',[HREngFrm001OutTime] AS '實際外出時間',[HREngFrm001DefBakTime] AS '預計返廠時間',[HREngFrm001BakTime] AS '實際返廠時間'");
                    sbSql.AppendFormat(@"  ,[TaskId] AS 'TaskId',[HREngFrm001SN] AS '表單編號',[HREngFrm001Date] AS '申請日期',[HREngFrm001UsrDpt] AS '部門',[HREngFrm001Location] AS '外出地點',[HREngFrm001Agent] AS '代理人',[HREngFrm001Cause] AS '外出原因',[HREngFrm001FF] AS '是否由公司出發',[HREngFrm001CH] AS '是否回廠',[CRADNO] AS '卡號'");
                    sbSql.AppendFormat(@"  FROM [TKGAFFAIRS].[dbo].[HREngFrm001]");
                    sbSql.AppendFormat(@"  WHERE ISNULL([HREngFrm001SN],'')<>''");
                    sbSql.AppendFormat(@"  AND  [HREngFrm001OutDate]='{0}' AND [CRADNO]='{1}' ", DateTime.Now.ToString("yyyy/MM/dd"), CARDNO);
                    sbSql.AppendFormat(@"  ORDER BY [HREngFrm001DefOutTime]");

                    sbSql.AppendFormat(@"  

                                        ");
                }

               
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
                sqlConn.Open();
                ds2.Clear();
                adapter2.Fill(ds2, "ds2");
                sqlConn.Close();


                if (ds2.Tables["ds2"].Rows.Count == 0)
                {
                    dataGridView1.DataSource = null;

                }
                else
                {
                    if (ds2.Tables["ds2"].Rows.Count >= 1)
                    {
                        dataGridView1.DataSource = ds2.Tables["ds2"];
                        dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Arial", 10);

                        dataGridView1.AutoResizeColumns();

                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            DataGridViewRow row = dataGridView1.Rows[i];
                            row.Height = 60;
                        }



                    }

                }

            }
            catch
            {

            }
            finally
            {

            }
        }


        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            ADDZ_SCSHR_LEAVE();
        }

        #endregion

      
    }
}