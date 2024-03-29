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

        //String connectionStringTKSG = "server=192.168.1.105;database=TKSG;uid=sa;pwd=dsc";
        //String connectionStringTKGAFFAIRS = "server=192.168.1.105;database=TKGAFFAIRS;uid=sa;pwd=dsc";

        String connectionStringTKSG = "server=192.168.1.105;database=TKSG;uid=tkdb;pwd=tkfood";
        String connectionStringTKGAFFAIRS = "server=192.168.1.105;database=TKGAFFAIRS;uid=tkdb;pwd=tkfood";
        String connectionStringUOF = "server=192.168.1.223;database=UOF;uid=TKUOF;pwd=TKUOF123456";

        string DB = "UOF";
        string IP = "192.168.1.239";

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
            timer1.Interval = 1000 * 60;
            timer1.Start();
        }


        #region FUNCTION
        private void timer1_Tick(object sender, EventArgs e)
        {
            //轉入刷卡資料+卡號
            ADDZ_SCSHR_LEAVE();

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
                    using (SqlConnection sqlConn = new SqlConnection(connectionStringTKGAFFAIRS))
                    {
                        sqlConn.Open();
                        using (SqlTransaction tran = sqlConn.BeginTransaction())
                        {
                            try
                            {
                                foreach (DataRow row in Z_SCSHR_LEAV.Tables[0].Rows)
                                {
                                    using (SqlCommand cmd = new SqlCommand())
                                    {
                                        cmd.Connection = sqlConn;
                                        cmd.Transaction = tran;
                                        cmd.CommandTimeout = 60;

                                        cmd.CommandText = @"
                                                            INSERT INTO [TKGAFFAIRS].[dbo].[Z_SCSHR_LEAVE]
                                                            ([DOC_NBR],[TASK_STATUS],[TASK_RESULT],[GROUP_CODE],[APPLICANT],[APPLICANTGUID],[APPLICANTCOMP],[APPLICANTDEPT],[APPLICANTDATE],[LEAEMP],[LEAAGENT],[LEACODE],[LEACODENAME],[SP_DATE],[SP_NAME],[STARTTIME],[ENDTIME],[LEAHOURS],[LEADAYS],[REMARK],[CANCEL_DOC_NBR],[CANCEL_STATUS],[SCSHR],[SCSHRMSG],[CRADNO],[NAME])
                                                            VALUES
                                                            (@DOC_NBR, @TASK_STATUS, @TASK_RESULT, @GROUP_CODE, @APPLICANT, @APPLICANTGUID, @APPLICANTCOMP, @APPLICANTDEPT, @APPLICANTDATE, @LEAEMP, @LEAAGENT, @LEACODE, @LEACODENAME, @SP_DATE, @SP_NAME, @STARTTIME, @ENDTIME, @LEAHOURS, @LEADAYS, @REMARK, @CANCEL_DOC_NBR, @CANCEL_STATUS, @SCSHR, @SCSHRMSG, @CardNo, @Name)
                                                        ";

                                        // 添加參數並設定值
                                        cmd.Parameters.Add(new SqlParameter("@DOC_NBR", SqlDbType.VarChar) { Value = row["DOC_NBR"].ToString() });
                                        cmd.Parameters.Add(new SqlParameter("@TASK_STATUS", SqlDbType.VarChar) { Value = row["TASK_STATUS"].ToString() });
                                        cmd.Parameters.Add(new SqlParameter("@TASK_RESULT", SqlDbType.VarChar) { Value = row["TASK_RESULT"].ToString() });
                                        cmd.Parameters.Add(new SqlParameter("@GROUP_CODE", SqlDbType.VarChar) { Value = row["GROUP_CODE"].ToString() });
                                        cmd.Parameters.Add(new SqlParameter("@APPLICANT", SqlDbType.VarChar) { Value = row["APPLICANT"].ToString() });
                                        cmd.Parameters.Add(new SqlParameter("@APPLICANTGUID", SqlDbType.VarChar) { Value = row["APPLICANTGUID"].ToString() });
                                        cmd.Parameters.Add(new SqlParameter("@APPLICANTCOMP", SqlDbType.VarChar) { Value = row["APPLICANTCOMP"].ToString() });
                                        cmd.Parameters.Add(new SqlParameter("@APPLICANTDEPT", SqlDbType.VarChar) { Value = row["APPLICANTDEPT"].ToString() });
                                        cmd.Parameters.Add(new SqlParameter("@APPLICANTDATE", SqlDbType.VarChar) { Value = row["APPLICANTDATE"].ToString() });
                                        cmd.Parameters.Add(new SqlParameter("@LEAEMP", SqlDbType.VarChar) { Value = row["LEAEMP"].ToString() });
                                        cmd.Parameters.Add(new SqlParameter("@LEAAGENT", SqlDbType.VarChar) { Value = row["LEAAGENT"].ToString() });
                                        cmd.Parameters.Add(new SqlParameter("@LEACODE", SqlDbType.VarChar) { Value = row["LEACODE"].ToString() });
                                        cmd.Parameters.Add(new SqlParameter("@LEACODENAME", SqlDbType.VarChar) { Value = row["LEACODENAME"].ToString() });
                                        cmd.Parameters.Add(new SqlParameter("@SP_DATE", SqlDbType.VarChar) { Value = row["SP_DATE"].ToString() });
                                        cmd.Parameters.Add(new SqlParameter("@SP_NAME", SqlDbType.VarChar) { Value = row["SP_NAME"].ToString() });
                                        cmd.Parameters.Add(new SqlParameter("@STARTTIME", SqlDbType.VarChar) { Value = row["STARTTIME"].ToString() });
                                        cmd.Parameters.Add(new SqlParameter("@ENDTIME", SqlDbType.VarChar) { Value = row["ENDTIME"].ToString() });
                                        cmd.Parameters.Add(new SqlParameter("@LEAHOURS", SqlDbType.VarChar) { Value = row["LEAHOURS"].ToString() });
                                        cmd.Parameters.Add(new SqlParameter("@LEADAYS", SqlDbType.VarChar) { Value = row["LEADAYS"].ToString() });
                                        cmd.Parameters.Add(new SqlParameter("@REMARK", SqlDbType.VarChar) { Value = row["REMARK"].ToString() });
                                        cmd.Parameters.Add(new SqlParameter("@CANCEL_DOC_NBR", SqlDbType.VarChar) { Value = row["CANCEL_DOC_NBR"].ToString() });
                                        cmd.Parameters.Add(new SqlParameter("@CANCEL_STATUS", SqlDbType.VarChar) { Value = row["CANCEL_STATUS"].ToString() });
                                        cmd.Parameters.Add(new SqlParameter("@SCSHR", SqlDbType.VarChar) { Value = row["SCSHR"].ToString() });
                                        cmd.Parameters.Add(new SqlParameter("@SCSHRMSG", SqlDbType.VarChar) { Value = row["SCSHRMSG"].ToString() });
                                        cmd.Parameters.Add(new SqlParameter("@CardNo", SqlDbType.VarChar) { Value = row["CardNo"].ToString() });
                                        cmd.Parameters.Add(new SqlParameter("@Name", SqlDbType.VarChar) { Value = row["Name"].ToString() });

                                       
                                        // 繼續添加其他參數...

                                        // 繼續添加其他參數...

                                        // 執行SQL命令
                                        int result = cmd.ExecuteNonQuery();

                                        if (result == 0)
                                        {
                                            tran.Rollback(); // 交易取消
                                        }
                                        else
                                        {
                                            // 如果有需要，可以在這裡添加其他處理邏輯

                                            tran.Commit(); // 執行交易
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                // 處理例外狀況
                                tran.Rollback();
                                MessageBox.Show($"發生錯誤: {ex.Message}");
                                //Console.WriteLine($"發生錯誤: {ex.Message}");
                                // 可以添加日誌或其他錯誤處理邏輯
                            }
                        }
                    }
                }
                catch
                {
                    PREPARE_MAIL();
                }
                 
                finally
                {
                    sqlConn.Close(); 
                }
            }


            //更新簽核的狀態 TASK_RESULT=0，結案
            UPDATE_Z_SCSHR_LEAVE();
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

                //TASK_RESULT IN ('0') 核準
                //ISNULL(TASK_RESULT,'')='' 啟單
                // AND (TASK_RESULT IN ('0') OR ISNULL(TASK_RESULT,'')='') 

                sbSql.AppendFormat(@" 
                                    SELECT  
                                    [DOC_NBR]
                                    ,[TASK_STATUS]
                                    ,[TASK_RESULT]
                                    ,[GROUP_CODE]
                                    ,[TB_EB_USER].ACCOUNT AS [APPLICANT]
                                    ,[LEAEMP] AS [APPLICANTGUID]
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
                                    ,REPLACE(REPLACE(REPLACE([REMARK],'`','') ,'‵',''),'\', '') AS REMARK
                                    ,[CANCEL_DOC_NBR]
                                    ,[CANCEL_STATUS]
                                    ,[SCSHR]
                                    ,[SCSHRMSG]
                                    ,[CardNo]
                                    ,[Name]
                                    FROM [192.168.1.223].[{0}].[dbo].[Z_SCSHR_LEAVE]
                                    LEFT JOIN [192.168.1.223].[UOF].[dbo].[TB_EB_USER] ON [TB_EB_USER].[USER_GUID]=[Z_SCSHR_LEAVE].[LEAEMP]
                                    LEFT JOIN [192.168.1.225].[CHIYU].[dbo].[Person] ON [APPLICANT]=[EmployeeID] COLLATE Chinese_PRC_CI_AS
                                    WHERE  1=1
                                    AND (TASK_RESULT IN ('0') OR ISNULL(TASK_RESULT,'')='') 
                                    AND [DOC_NBR] COLLATE Chinese_Taiwan_Stroke_BIN NOT IN (SELECT [DOC_NBR] FROM [TKGAFFAIRS].[dbo].[Z_SCSHR_LEAVE]) 
                           

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

        public void UPDATE_Z_SCSHR_LEAVE()
        {
            try
            {
                connectionString = connectionStringTKGAFFAIRS;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@" 
                                            UPDATE [TKGAFFAIRS].[dbo].[Z_SCSHR_LEAVE]
                                            SET [Z_SCSHR_LEAVE].[TASK_RESULT]=TEMP.TASK_RESULT
                                            FROM 
                                            (
                                            SELECT 
                                            ORI_CSHR_LEAVE.[DOC_NBR] AS DOC_NBR
                                            ,ORI_CSHR_LEAVE.[TASK_STATUS] AS TASK_STATUS
                                            ,ORI_CSHR_LEAVE.[TASK_RESULT] AS TASK_RESULT
                                            ,TO_Z_SCSHR_LEAVE.[DOC_NBR] NEWDOC_NBR
                                            ,TO_Z_SCSHR_LEAVE.[TASK_STATUS]  NEWTASK_STATUS
                                            ,TO_Z_SCSHR_LEAVE.[TASK_RESULT]  NEWTASK_RESULT
                                            FROM [192.168.1.223].[UOF].[dbo].[Z_SCSHR_LEAVE] AS ORI_CSHR_LEAVE ,[TKGAFFAIRS].[dbo].[Z_SCSHR_LEAVE] AS  TO_Z_SCSHR_LEAVE
                                            WHERE ORI_CSHR_LEAVE.[DOC_NBR]=TO_Z_SCSHR_LEAVE.[DOC_NBR] COLLATE Chinese_Taiwan_Stroke_BIN
                                            AND ISNULL(TO_Z_SCSHR_LEAVE.[TASK_RESULT],'')=''
                                            AND ORI_CSHR_LEAVE.[TASK_RESULT]='0'
                                            ) AS TEMP 
                                            WHERE TEMP.DOC_NBR=[Z_SCSHR_LEAVE].DOC_NBR COLLATE Chinese_Taiwan_Stroke_BIN
                                            ");

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

                    //MessageBox.Show("完成");
                }
            }
            catch
            {
                PREPARE_MAIL();
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
                    sbSql.AppendFormat(@"  
                                        SELECT 
                                         [LEACODENAME]
                                         ,[NAME] AS '申請人'
                                        ,[CRADNO] AS '卡號'
                                        ,[DOC_NBR]  AS '表單編號'
                                        ,[STARTTIME]  AS '預計外出時間'
                                        ,[ENDTIME] AS '預計返廠時間'
                                        ,[TASK_STATUS]
                                        ,[TASK_RESULT]
                                        ,[GROUP_CODE]
                                        ,[APPLICANT]
                                        ,[APPLICANTGUID]
                                        ,[APPLICANTCOMP]
                                        ,[APPLICANTDEPT]
                                        ,[APPLICANTDATE]
                                        ,[LEAEMP]
                                        ,[LEAAGENT]
                                        
                                        ,[LEACODENAME]
                                        ,[SP_DATE]
                                        ,[SP_NAME]
                                        ,[LEAHOURS]
                                        ,[LEADAYS]
                                        ,[REMARK]
                                        ,[CANCEL_DOC_NBR]
                                        ,[CANCEL_STATUS]
                                        ,[SCSHR]
                                        ,[SCSHRMSG]
                                        FROM [TKGAFFAIRS].[dbo].[Z_SCSHR_LEAVE]
                                        WHERE CONVERT(NVARCHAR,[STARTTIME],112)='{0}'

                                        ", DateTime.Now.ToString("yyyyMMdd"));
                }
                else
                {
                    sbSql.AppendFormat(@"  
                                        SELECT 
                                        [LEACODENAME]
                                        ,[NAME] AS '申請人'                                      
                                        ,[CRADNO] AS '卡號'
                                        ,[DOC_NBR]  AS '表單編號'
                                        ,[STARTTIME]  AS '預計外出時間'
                                        ,[ENDTIME] AS '預計返廠時間'
                                        ,[TASK_STATUS]
                                        ,[TASK_RESULT]
                                        ,[GROUP_CODE]
                                        ,[APPLICANT]
                                        ,[APPLICANTGUID]
                                        ,[APPLICANTCOMP]
                                        ,[APPLICANTDEPT]
                                        ,[APPLICANTDATE]
                                        ,[LEAEMP]
                                        ,[LEAAGENT]
                                        ,[LEACODE]
                                       
                                        ,[SP_DATE]
                                        ,[SP_NAME]
                                        ,[LEAHOURS]
                                        ,[LEADAYS]
                                        ,[REMARK]
                                        ,[CANCEL_DOC_NBR]
                                        ,[CANCEL_STATUS]
                                        ,[SCSHR]
                                        ,[SCSHRMSG]
                                        FROM [TKGAFFAIRS].[dbo].[Z_SCSHR_LEAVE]
                                        WHERE CONVERT(NVARCHAR,[STARTTIME],112)='{0}'
                                        AND  [CRADNO]='{1}'

                                        ", DateTime.Now.ToString("yyyyMMdd"), CARDNO);

                }
               
               
            

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
        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (!string.IsNullOrEmpty(textBox1.Text.Trim()))
                {
                    //log
                    ADD_LOG_TB_EIP_DUTY_TEMP(textBox1.Text.Trim(), "離廠","");

                    SEARCHHREngFrm001textBox1(textBox1.Text.Trim());
                }
            }
        }

        private void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {

                if (!string.IsNullOrEmpty(textBox2.Text.Trim()))
                {

                    //log
                    ADD_LOG_TB_EIP_DUTY_TEMP(textBox2.Text.Trim(), "回廠","");


                    SEARCHHREngFrm001textBox2(textBox2.Text.Trim());
                }
            }
        }
        public void SEARCHHREngFrm001textBox1(string CARDNO)
        {
            DataSet ds = new DataSet();

            try
            {
                connectionString = connectionStringTKGAFFAIRS;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                StringBuilder query = new StringBuilder();
                
                sbSql.AppendFormat(@"  
                                    SELECT 
                                    [LEACODENAME]
                                    ,[NAME] AS '申請人'
                                    ,[CRADNO] AS '卡號'
                                    ,[DOC_NBR]  AS '表單編號'
                                    ,[STARTTIME]  AS '預計外出時間'
                                    ,[ENDTIME] AS '預計返廠時間'
                                    ,[TASK_STATUS]
                                    ,[TASK_RESULT]
                                    ,[GROUP_CODE]
                                    ,[APPLICANT]
                                    ,[APPLICANTGUID]
                                    ,[APPLICANTCOMP]
                                    ,[APPLICANTDEPT]
                                    ,[APPLICANTDATE]
                                    ,[LEAEMP]
                                    ,[LEAAGENT]
                                    ,[LEACODE]
                                    
                                    ,[SP_DATE]
                                    ,[SP_NAME]
                                    ,[LEAHOURS]
                                    ,[LEADAYS]
                                    ,[REMARK]
                                    ,[CANCEL_DOC_NBR]
                                    ,[CANCEL_STATUS]
                                    ,[SCSHR]
                                    ,[SCSHRMSG]
                                    FROM [TKGAFFAIRS].[dbo].[Z_SCSHR_LEAVE]
                                    WHERE [LEACODE] IN (SELECT  [LEACODE]   FROM [TKGAFFAIRS].[dbo].[Z_SCSHR_LEAVELEACODE])
                                    AND [TASK_RESULT]='0'
                                    AND CONVERT(NVARCHAR,[STARTTIME],112)='{0}'
                                    AND [CRADNO]='{1}'
                              
                                     ", DateTime.Now.ToString("yyyyMMdd"), CARDNO);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();


                if (ds.Tables["TEMPds1"].Rows.Count == 0)
                {
                    //dataGridView1.DataSource = null;

                    string CRADNO = textBox1.Text.Trim();

                    CHECKWHITELIST("離開公司");

                    SEARCHHREngFrm001C(CRADNO);
                }
                else
                {
                    if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        //檢查是否已核單，核過才能出門
                        //TASK_RESULT=0
                        string TASK_RESULT = ds.Tables["TEMPds1"].Rows[0]["TASK_RESULT"].ToString();
                        
                        if(TASK_RESULT.Equals("0"))
                        {
                            string NAME = ds.Tables["TEMPds1"].Rows[0]["申請人"].ToString();
                            string CRADNO = ds.Tables["TEMPds1"].Rows[0]["卡號"].ToString();
                            string APPLICANT = ds.Tables["TEMPds1"].Rows[0]["APPLICANT"].ToString();

                            ADDTB_EIP_DUTY_TEMP(APPLICANT, "Off", IP);

                            SEARCHHREngFrm001B(CARDNO);

                            MessageBox.Show("實際外出時間: " + DateTime.Now.ToString("HH:mm") + " " + NAME);

                            textBox1.Text = null;
                        }
                        else
                        {
                            MessageBox.Show("主管未核準申請單，不允許離廠");
                            CHECKWHITELIST("離開公司");
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

        public void SEARCHHREngFrm001textBox2(string CARDNO)
        {
            DataSet ds = new DataSet();
            try
            {
                connectionString = connectionStringTKGAFFAIRS;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                StringBuilder query = new StringBuilder();
                sbSql.AppendFormat(@"  
                                    SELECT 
                                    [LEACODENAME]
                                    ,[NAME] AS '申請人'
                                    ,[CRADNO] AS '卡號'
                                    ,[DOC_NBR]  AS '表單編號'
                                    ,[STARTTIME]  AS '預計外出時間'
                                    ,[ENDTIME] AS '預計返廠時間'
                                    ,[TASK_STATUS]
                                    ,[TASK_RESULT]
                                    ,[GROUP_CODE]
                                    ,[APPLICANT]
                                    ,[APPLICANTGUID]
                                    ,[APPLICANTCOMP]
                                    ,[APPLICANTDEPT]
                                    ,[APPLICANTDATE]
                                    ,[LEAEMP]
                                    ,[LEAAGENT]
                                    ,[LEACODE]
                                    
                                    ,[SP_DATE]
                                    ,[SP_NAME]
                                    ,[LEAHOURS]
                                    ,[LEADAYS]
                                    ,[REMARK]
                                    ,[CANCEL_DOC_NBR]
                                    ,[CANCEL_STATUS]
                                    ,[SCSHR]
                                    ,[SCSHRMSG]
                                    FROM [TKGAFFAIRS].[dbo].[Z_SCSHR_LEAVE]
                                    WHERE [LEACODE] IN (SELECT  [LEACODE]   FROM [TKGAFFAIRS].[dbo].[Z_SCSHR_LEAVELEACODE])
                                    AND [TASK_RESULT]='0'
                                    AND CONVERT(NVARCHAR,[STARTTIME],112)='{0}'
                                    AND [CRADNO]='{1}'
                                     ", DateTime.Now.ToString("yyyyMMdd"), CARDNO);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();


                if (ds.Tables["TEMPds1"].Rows.Count == 0)
                {
                    //dataGridView1.DataSource = null;
                    CARDNO = textBox2.Text.Trim();

                    CHECKWHITELIST("返回公司");

                    SEARCHHREngFrm001C(CARDNO);
                }
                else
                {
                    if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        string NAME = ds.Tables["TEMPds1"].Rows[0]["申請人"].ToString();
                        string CRADNO = ds.Tables["TEMPds1"].Rows[0]["卡號"].ToString();
                        string APPLICANT = ds.Tables["TEMPds1"].Rows[0]["APPLICANT"].ToString();

                        ADDTB_EIP_DUTY_TEMP(APPLICANT, "Work", IP);

                        SEARCHHREngFrm001B(CARDNO);

                        MessageBox.Show("實際回廠時間: " + DateTime.Now.ToString("HH:mm") + " " + NAME);

                        textBox2.Text = null;

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

        public void CHECKWHITELIST(string MODIFYCASUE)
        {
            string STATUS1 = "N";
            string STATUS2 = "N";

            if (!string.IsNullOrEmpty(textBox1.Text))
            {
                DataSet ds = new DataSet();
                SqlDataAdapter adapter = new SqlDataAdapter();
                SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();

                try
                {
                    connectionString = connectionStringTKGAFFAIRS;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();


                    sbSql.AppendFormat(@"  
                                     SELECT [WHITELIST].[ID],[Person].[CardNo],[WHITELIST].[NAME] 
                                     FROM [TKGAFFAIRS].[dbo].[WHITELIST]
                                     LEFT JOIN [192.168.1.225].[CHIYU].[dbo].[Person] ON [WHITELIST].ID=[Person].[UserID]
                                     WHERE ISNULL([Person].[CardNo],'')<>''
                                    ");


                    adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder = new SqlCommandBuilder(adapter);
                    sqlConn.Open();
                    ds.Clear();
                    adapter.Fill(ds, "ds");
                    sqlConn.Close();


                    if (ds.Tables["ds"].Rows.Count == 0)
                    {

                    }
                    else
                    {
                        if (ds.Tables["ds"].Rows.Count >= 1)
                        {
                            foreach (DataRow dr in ds.Tables["ds"].Rows)
                            {
                                if (dr["CardNo"].ToString().Trim().Equals(textBox1.Text.Trim()))
                                {                                    
                                    string NAME = dr["NAME"].ToString().Trim();
                                    string CRADNO = dr["CardNo"].ToString().Trim();
                                    string ID = dr["ID"].ToString().Trim();

                                    ADDTB_EIP_DUTY_TEMP(ID, "Off", IP);

                                    STATUS1 = "Y";
                                    MessageBox.Show("確認 白名單 人員，可以進出 :" + textBox1.Text.Trim()+" "+ NAME);

                                    textBox1.Text = null;
                                    //MessageBox.Show(textBox1.Text);
                                }
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

                if (STATUS1.Equals("N"))
                {
                    MessageBox.Show("查無申請單，不可進出");
                    textBox1.Text = null;
                }
            }
            else if (!string.IsNullOrEmpty(textBox2.Text))
            {
                DataSet ds = new DataSet();
                SqlDataAdapter adapter = new SqlDataAdapter();
                SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();

                try
                {
                    connectionString = connectionStringTKGAFFAIRS;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@"  
                                         SELECT [WHITELIST].[ID],[Person].[CardNo],[WHITELIST].[NAME] 
                                         FROM [TKGAFFAIRS].[dbo].[WHITELIST]
                                         LEFT JOIN [192.168.1.225].[CHIYU].[dbo].[Person] ON [WHITELIST].ID=[Person].[UserID]
                                         WHERE ISNULL([Person].[CardNo],'')<>''
                                        ");


                    adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder = new SqlCommandBuilder(adapter);
                    sqlConn.Open();
                    ds.Clear();
                    adapter.Fill(ds, "ds");
                    sqlConn.Close();


                    if (ds.Tables["ds"].Rows.Count == 0)
                    {

                    }
                    else
                    {
                        if (ds.Tables["ds"].Rows.Count >= 1)
                        {
                            foreach (DataRow dr in ds.Tables["ds"].Rows)
                            {
                                if (dr["CardNo"].ToString().Trim().Equals(textBox2.Text.Trim()))
                                {
                                    string NAME = dr["NAME"].ToString().Trim();
                                    string CRADNO = dr["CardNo"].ToString().Trim();
                                    string ID = dr["ID"].ToString().Trim();

                                    ADDTB_EIP_DUTY_TEMP(ID, "Work", IP);

                                    STATUS2 = "Y";
                                    MessageBox.Show("確認 白名單 人員，可以進出:" + textBox2.Text.Trim() + " " + NAME);
                                    textBox2.Text = null;
                                }
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

                if (STATUS2.Equals("N"))
                {
                    MessageBox.Show("查無申請單，不可進出");
                    textBox1.Text = null;
                }
            }
        }

        public void SEARCHHREngFrm001C(string CARDNO)
        {

            try
            {
                connectionString = connectionStringTKGAFFAIRS;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                StringBuilder query = new StringBuilder();

                if (!string.IsNullOrEmpty(CARDNO))
                {
                    sbSql.AppendFormat(@"  
                                        SELECT 
                                         [LEACODENAME]
                                        ,[NAME] AS '申請人'
                                        ,[CRADNO] AS '卡號'
                                        ,[DOC_NBR]  AS '表單編號'
                                        ,[STARTTIME]  AS '預計外出時間'
                                        ,[ENDTIME] AS '預計返廠時間'
                                        ,[TASK_STATUS]
                                        ,[TASK_RESULT]
                                        ,[GROUP_CODE]
                                        ,[APPLICANT]
                                        ,[APPLICANTGUID]
                                        ,[APPLICANTCOMP]
                                        ,[APPLICANTDEPT]
                                        ,[APPLICANTDATE]
                                        ,[LEAEMP]
                                        ,[LEAAGENT]
                                        ,[LEACODE]
                                       
                                        ,[SP_DATE]
                                        ,[SP_NAME]
                                        ,[LEAHOURS]
                                        ,[LEADAYS]
                                        ,[REMARK]
                                        ,[CANCEL_DOC_NBR]
                                        ,[CANCEL_STATUS]
                                        ,[SCSHR]
                                        ,[SCSHRMSG]
                                        FROM [TKGAFFAIRS].[dbo].[Z_SCSHR_LEAVE]
                                        WHERE CONVERT(NVARCHAR,[STARTTIME],112)='{0}'
                                        AND  [CRADNO]='{1}'

                                        ", DateTime.Now.ToString("yyyyMMdd"), CARDNO);
                }



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

        public void ADDTB_EIP_DUTY_TEMP(string CARD_NO,string TYPE, string IP_ADDRESS)
        {
            try
            {

                connectionString = connectionStringUOF;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                              
                sbSql.AppendFormat(@" 
                                    INSERT INTO [UOF].[dbo].[TB_EIP_DUTY_TEMP]
                                    (
                                    [PUNCH_TEMP_ID],[CARD_NO],[PUNCH_TIME],[TYPE],[CREATE_TIME],[IP_ADDRESS],[CLOCK_CODE]
                                    )
                                    VALUES (NEWID(),'{0}',GETDATE(),'{1}',GETDATE(),'{2}','')

                                    ", CARD_NO, TYPE, IP_ADDRESS);

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                    ADD_LOG_TB_EIP_DUTY_TEMP(textBox1.Text.Trim(), "離廠", "記錄失敗");

                }
                else
                {
                    tran.Commit();      //執行交易  


                }
            }
            catch (Exception ex)
            {
                ADD_LOG_TB_EIP_DUTY_TEMP(textBox1.Text.Trim(), "離廠", "記錄失敗"+ ex.Message.ToString());
            }

            finally
            {
                sqlConn.Close();
            }
        }

        private void textBox3_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (!string.IsNullOrEmpty(textBox3.Text.Trim()))
                {

                    //log
                    ADD_LOG_TB_EIP_DUTY_TEMP(textBox3.Text.Trim(), "離廠", "");

                    SEARCHHREngFrm001textBox3(textBox3.Text.Trim());
                }
            }
        }

        private void textBox4_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {

                if (!string.IsNullOrEmpty(textBox4.Text.Trim()))
                {
                    //log
                    ADD_LOG_TB_EIP_DUTY_TEMP(textBox4.Text.Trim(), "回廠", "");

                    SEARCHHREngFrm001textBox4(textBox4.Text.Trim());
                }
            }
        }

        public void SEARCHHREngFrm001textBox3(string CARDNO)
        {
            DataSet ds = new DataSet();

            try
            {
                connectionString = connectionStringTKGAFFAIRS;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                StringBuilder query = new StringBuilder();

                sbSql.AppendFormat(@"  
                                    SELECT 
                                    [LEACODENAME]
                                    ,[NAME] AS '申請人'
                                    ,[CRADNO] AS '卡號'
                                    ,[DOC_NBR]  AS '表單編號'
                                    ,[STARTTIME]  AS '預計外出時間'
                                    ,[ENDTIME] AS '預計返廠時間'
                                    ,[TASK_STATUS]
                                    ,[TASK_RESULT]
                                    ,[GROUP_CODE]
                                    ,[APPLICANT]
                                    ,[APPLICANTGUID]
                                    ,[APPLICANTCOMP]
                                    ,[APPLICANTDEPT]
                                    ,[APPLICANTDATE]
                                    ,[LEAEMP]
                                    ,[LEAAGENT]
                                    ,[LEACODE]
                                    
                                    ,[SP_DATE]
                                    ,[SP_NAME]
                                    ,[LEAHOURS]
                                    ,[LEADAYS]
                                    ,[REMARK]
                                    ,[CANCEL_DOC_NBR]
                                    ,[CANCEL_STATUS]
                                    ,[SCSHR]
                                    ,[SCSHRMSG]
                                    FROM [TKGAFFAIRS].[dbo].[Z_SCSHR_LEAVE]
                                    WHERE [LEACODE] NOT IN (SELECT  [LEACODE]   FROM [TKGAFFAIRS].[dbo].[Z_SCSHR_LEAVELEACODE])
                                    AND [TASK_RESULT]='0'
                                    AND CONVERT(NVARCHAR,[STARTTIME],112)='{0}'
                                    AND [CRADNO]='{1}'
                                     ", DateTime.Now.ToString("yyyyMMdd"), CARDNO);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();


                if (ds.Tables["TEMPds1"].Rows.Count == 0)
                {
                    ////dataGridView1.DataSource = null;

                    //string CRADNO = textBox1.Text.Trim();

                    //CHECKWHITELIST("離開公司");

                    //SEARCHHREngFrm001C(CRADNO);

                    MessageBox.Show("查無申請單，不可進出");
                    textBox3.Text = null;
                }
                else
                {
                    if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        //檢查是否已核單，核過才能出門
                        //TASK_RESULT=0
                        string TASK_RESULT = ds.Tables["TEMPds1"].Rows[0]["TASK_RESULT"].ToString();

                        if (TASK_RESULT.Equals("0"))
                        {
                            string NAME = ds.Tables["TEMPds1"].Rows[0]["申請人"].ToString();
                            string CRADNO = ds.Tables["TEMPds1"].Rows[0]["卡號"].ToString();
                            string APPLICANT = ds.Tables["TEMPds1"].Rows[0]["APPLICANT"].ToString();

                            ADDTB_EIP_DUTY_TEMP(APPLICANT, "Off", IP);

                            SEARCHHREngFrm001B(CARDNO);

                            MessageBox.Show("實際外出時間: " + DateTime.Now.ToString("HH:mm") + " " + NAME);

                            textBox3.Text = null;
                        }
                        else
                        {

                            MessageBox.Show("主管未核準，不允許離廠");
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

        public void SEARCHHREngFrm001textBox4(string CARDNO)
        {
            DataSet ds = new DataSet();
            try
            {
                connectionString = connectionStringTKGAFFAIRS;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                StringBuilder query = new StringBuilder();
                sbSql.AppendFormat(@"  
                                    SELECT 
                                     [LEACODENAME]
                                    ,[NAME] AS '申請人'
                                    ,[CRADNO] AS '卡號'
                                    ,[DOC_NBR]  AS '表單編號'
                                    ,[STARTTIME]  AS '預計外出時間'
                                    ,[ENDTIME] AS '預計返廠時間'
                                    ,[TASK_STATUS]
                                    ,[TASK_RESULT]
                                    ,[GROUP_CODE]
                                    ,[APPLICANT]
                                    ,[APPLICANTGUID]
                                    ,[APPLICANTCOMP]
                                    ,[APPLICANTDEPT]
                                    ,[APPLICANTDATE]
                                    ,[LEAEMP]
                                    ,[LEAAGENT]
                                    ,[LEACODE]
                                    
                                    ,[SP_DATE]
                                    ,[SP_NAME]
                                    ,[LEAHOURS]
                                    ,[LEADAYS]
                                    ,[REMARK]
                                    ,[CANCEL_DOC_NBR]
                                    ,[CANCEL_STATUS]
                                    ,[SCSHR]
                                    ,[SCSHRMSG]
                                    FROM [TKGAFFAIRS].[dbo].[Z_SCSHR_LEAVE]
                                    WHERE [LEACODE] NOT IN (SELECT  [LEACODE]   FROM [TKGAFFAIRS].[dbo].[Z_SCSHR_LEAVELEACODE])
                                    AND [TASK_RESULT]='0'
                                    AND CONVERT(NVARCHAR,[STARTTIME],112)='{0}'
                                    AND [CRADNO]='{1}'
                                     ", DateTime.Now.ToString("yyyyMMdd"), CARDNO);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();


                if (ds.Tables["TEMPds1"].Rows.Count == 0)
                {
                    ////dataGridView1.DataSource = null;
                    //CARDNO = textBox2.Text.Trim();

                    //CHECKWHITELIST("返回公司");

                    //SEARCHHREngFrm001C(CARDNO);

                    MessageBox.Show("查無申請單，不可進出");
                    textBox4.Text = null;
                }
                else
                {
                    if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        string NAME = ds.Tables["TEMPds1"].Rows[0]["申請人"].ToString();
                        string CRADNO = ds.Tables["TEMPds1"].Rows[0]["卡號"].ToString();
                        string APPLICANT = ds.Tables["TEMPds1"].Rows[0]["APPLICANT"].ToString();

                        ADDTB_EIP_DUTY_TEMP(APPLICANT, "Work", IP);

                        SEARCHHREngFrm001B(CARDNO);

                        MessageBox.Show("實際回廠時間: " + DateTime.Now.ToString("HH:mm") + " " + NAME);

                        textBox4.Text = null;

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

        public void ADD_LOG_TB_EIP_DUTY_TEMP(string CARD_NO,string TYPE,string COMMENTS)
        {
            try
            {

                connectionString = connectionStringTKGAFFAIRS;

                using (SqlConnection sqlConn = new SqlConnection(connectionString))
                {
                    sqlConn.Open();

                    using (SqlTransaction tran = sqlConn.BeginTransaction())
                    {
                        using (SqlCommand cmd = new SqlCommand())
                        {
                            cmd.Connection = sqlConn;
                            cmd.Transaction = tran;
                            cmd.CommandTimeout = 60;

                            cmd.CommandText = @"
                                                INSERT INTO [TKGAFFAIRS].[dbo].[LOG_TB_EIP_DUTY_TEMP]
                                                ([CARD_NO],[TYPE],[COMMENTS])
                                                VALUES
                                                (@CARD_NO, @TYPE, @COMMENTS)";

                            // 添加參數
                            cmd.Parameters.AddWithValue("@CARD_NO", CARD_NO);
                            cmd.Parameters.AddWithValue("@TYPE", TYPE);
                            cmd.Parameters.AddWithValue("@COMMENTS", COMMENTS);

                            try
                            {
                                int result = cmd.ExecuteNonQuery();

                                // 提交交易
                                tran.Commit();

                                // 根據需要處理結果
                            }
                            catch (Exception ex)
                            {
                                // 發生錯誤時回滾交易
                                tran.Rollback();

                                // 根據需要處理錯誤
                            }
                        }
                    }
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

        public void PREPARE_MAIL()
        {
            try
            {
                StringBuilder SUBJEST = new StringBuilder();
                StringBuilder BODY = new StringBuilder();

                ////加上附圖
                //string path = System.Environment.CurrentDirectory+@"/Images/emaillogo.jpg";
                //LinkedResource res = new LinkedResource(path);
                //res.ContentId = Guid.NewGuid().ToString();

                SUBJEST.Clear();
                BODY.Clear();


                SUBJEST.AppendFormat(@"系統通知-老楊食品-每日-有請假單未轉入刷卡資料 ，謝謝。 " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));


                //交辨未完成的項目及交辨人回覆狀況
                BODY.AppendFormat("<span style='font-size:12.0pt;font-family:微軟正黑體'> <br>" + "Dear SIR:" + "<br>"
                    + "<br>" + "有請假單未轉入刷卡資料"

                    );


                //if (DSPROOFREAD.Tables[0].Rows.Count > 0)
                //{
                //    BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體'><br>" + "明細");

                //    BODY.AppendFormat(@"<table> ");
                //    BODY.AppendFormat(@"<tr >");
                //    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' width=10% "">交辨開始時間</th>");
                //    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' width=40% "">交辨項目</th>");
                //    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' width=40% "">交辨回覆</th>");
                //    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' width=10% "">被交辨人</th>");
                //    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' width=10% "">交辨狀態</th>");
                //    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' width=10% "">回覆時間</th>");

                //    BODY.AppendFormat(@"</tr> ");

                //    foreach (DataRow DR in DSPROOFREAD.Tables[0].Rows)
                //    {

                //        BODY.AppendFormat(@"<tr >");
                //        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["交辨開始時間"].ToString() + "</td>");
                //        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體'  width=40% "">" + DR["交辨項目"].ToString() + "</td>");
                //        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體'  width=30% "">" + DR["交辨回覆"].ToString() + "</td>");
                //        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["被交辨人"].ToString() + "</td>");
                //        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["交辨狀態"].ToString() + "</td>");
                //        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["回覆時間"].ToString() + "</td>");


                //        BODY.AppendFormat(@"</tr> ");

                //        //BODY.AppendFormat("<span></span>");
                //        //BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體' > <br> " + "品名     " + DR["TD005"].ToString() + "</span>");
                //        //BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體' > <br>" + "採購數量 " + DR["TD008"].ToString() + "</span>");
                //        //BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體' > <br>" + "採購單位 " + DR["TD009"].ToString() + "</span>");
                //        //BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體' > <br>");
                //    }
                //    BODY.AppendFormat(@"</table> ");
                //}
                //else
                //{
                //    BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體'><br>" + "無資料");
                //}
                              


                BODY.AppendFormat(" "
                             + "<br>" + "謝謝"

                             + "</span><br>");



                SENDEMAILUOFPROOFEAD(SUBJEST, BODY);

            }
            catch
            {

            }
            finally
            {

            }
        }
        /// <summary>
        /// 實際寄出
        /// </summary>
        public void SENDEMAILUOFPROOFEAD(StringBuilder Subject, StringBuilder Body)
        {
            try
            {
                string MySMTPCONFIG = "officemail.cloudmax.com.tw";
                string NAME = "tkpublic@tkfood.com.tw";
                string PW = "@@tkmail629";

                System.Net.Mail.MailMessage MyMail = new System.Net.Mail.MailMessage();
                MyMail.From = new System.Net.Mail.MailAddress("tk290@tkfood.com.tw");

                //MyMail.Bcc.Add("密件副本的收件者Mail"); //加入密件副本的Mail          
                //MyMail.Subject = "每日訂單-製令追踨表"+DateTime.Now.ToString("yyyy/MM/dd");
                MyMail.Subject = Subject.ToString();
                //MyMail.Body = "<h1>Dear SIR</h1>" + Environment.NewLine + "<h1>附件為每日訂單-製令追踨表，請查收</h1>" + Environment.NewLine + "<h1>若訂單沒有相對的製令則需通知製造生管開立</h1>"; //設定信件內容
                MyMail.Body = Body.ToString();
                MyMail.IsBodyHtml = true; //是否使用html格式

                //加上附圖
                //string path = System.Environment.CurrentDirectory + @"/Images/emaillogo.jpg";
                //MyMail.AlternateViews.Add(GetEmbeddedImage(path, Body));

                System.Net.Mail.SmtpClient MySMTP = new System.Net.Mail.SmtpClient(MySMTPCONFIG, 25);
                MySMTP.Credentials = new System.Net.NetworkCredential(NAME, PW);


                try
                {
                    //MyMail.To.Add(DR["MAIL"].ToString()); //設定收件者Email，多筆mail
                    MyMail.To.Add("tk290@tkfood.com.tw"); //設定收件者Email
                    MySMTP.Send(MyMail);

                    MyMail.Dispose(); //釋放資源


                }
                catch (Exception ex)
                {
                    MessageBox.Show("有錯誤");

                    //ADDLOG(DateTime.Now, Subject.ToString(), ex.ToString());
                    //ex.ToString();
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
            SEARCHHREngFrm001B("");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ADDZ_SCSHR_LEAVE();
        }



        #endregion


    }
}
