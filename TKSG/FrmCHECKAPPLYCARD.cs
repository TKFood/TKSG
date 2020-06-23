using System;
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
    public partial class FrmCHECKAPPLYCARD : Form
    {
        String connectionStringTKSG = "server=192.168.1.105;database=TKSG;uid=sa;pwd=dsc";
        String connectionStringTKGAFFAIRS = "server=192.168.1.105;database=TKGAFFAIRS;uid=sa;pwd=dsc";
        String connectionStringUOF = "server=192.168.1.223;database=UOF;uid=TKUOF;pwd=TKUOF123456";

        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlDataAdapter adapter2 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder2 = new SqlCommandBuilder();

        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataSet ds2 = new DataSet();

        string tablename = null;
        string EDITID;
        int result;
        Thread TD;

        string TaskId;
        string CARDNO;

        string DB = "UOF";
        //string DB = "UOFTEST";

        //用STATUS來控制在1分鐘內不得連續刷卡
        string STATUS = "Y";

        public FrmCHECKAPPLYCARD()
        {
            InitializeComponent();


            label6.Text = DateTime.Now.ToString("yyyy/MM/dd HH:mm");

            timer1.Enabled = true;
            timer1.Interval = 1000 * 1;
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

        }


        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox1.Text.Trim()))
            {
                SEARCHHREngFrm001(textBox1.Text.Trim());
            }
        }

        public void SEARCHHREngFrm001(string CARDNO)
        {

            try
            {
                connectionString = connectionStringTKGAFFAIRS;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                StringBuilder query = new StringBuilder();

                sbSql.AppendFormat(@"  SELECT ");
                sbSql.AppendFormat(@"  [HREngFrm001User] AS '申請人',[HREngFrm001Rank] AS '職級',[HREngFrm001OutDate] AS '外出日期',[HREngFrm001Transp] AS '交通工具',[HREngFrm001LicPlate] AS '車牌',[HREngFrm001DefOutTime] AS '預計外出時間',[HREngFrm001OutTime] AS '實際外出時間',[HREngFrm001DefBakTime] AS '預計返廠時間',[HREngFrm001BakTime] AS '實際返廠時間'");
                sbSql.AppendFormat(@"  ,[TaskId] AS 'TaskId',[HREngFrm001SN] AS '表單編號',[HREngFrm001Date] AS '申請日期',[HREngFrm001UsrDpt] AS '部門',[HREngFrm001Location] AS '外出地點',[HREngFrm001Agent] AS '代理人',[HREngFrm001Cause] AS '外出原因',[HREngFrm001FF] AS '是否由公司出發',[HREngFrm001CH] AS '是否回廠',[CRADNO] AS '卡號'");
                sbSql.AppendFormat(@"  FROM [TKGAFFAIRS].[dbo].[HREngFrm001]");
                sbSql.AppendFormat(@"  WHERE [HREngFrm001OutDate]='{0}' AND [CRADNO]='{1}'", DateTime.Now.ToString("yyyy/MM/dd"), CARDNO);
                sbSql.AppendFormat(@"  AND ((ISNULL([HREngFrm001OutTime],'')='' AND [HREngFrm001FF]='是' ) OR (ISNULL([HREngFrm001BakTime],'')='' AND [HREngFrm001CH]='是' ))");
                sbSql.AppendFormat(@"  ORDER BY [HREngFrm001DefOutTime]");
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();


                if (ds.Tables["TEMPds1"].Rows.Count == 0)
                {
                    dataGridView1.DataSource = null;

                    CHECKWHITELIST();

                }
                else
                {
                    if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        dataGridView1.DataSource = ds.Tables["TEMPds1"];
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

        public void SEARCHHREngFrm001B(string CARDNO)
        {

            try
            {
                connectionString = connectionStringTKGAFFAIRS;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                StringBuilder query = new StringBuilder();

                if(string.IsNullOrEmpty(CARDNO))
                {
                    sbSql.AppendFormat(@"  SELECT ");
                    sbSql.AppendFormat(@"  [HREngFrm001User] AS '申請人',[HREngFrm001Rank] AS '職級',[HREngFrm001OutDate] AS '外出日期',[HREngFrm001Transp] AS '交通工具',[HREngFrm001LicPlate] AS '車牌',[HREngFrm001DefOutTime] AS '預計外出時間',[HREngFrm001OutTime] AS '實際外出時間',[HREngFrm001DefBakTime] AS '預計返廠時間',[HREngFrm001BakTime] AS '實際返廠時間'");
                    sbSql.AppendFormat(@"  ,[TaskId] AS 'TaskId',[HREngFrm001SN] AS '表單編號',[HREngFrm001Date] AS '申請日期',[HREngFrm001UsrDpt] AS '部門',[HREngFrm001Location] AS '外出地點',[HREngFrm001Agent] AS '代理人',[HREngFrm001Cause] AS '外出原因',[HREngFrm001FF] AS '是否由公司出發',[HREngFrm001CH] AS '是否回廠',[CRADNO] AS '卡號'");
                    sbSql.AppendFormat(@"  FROM [TKGAFFAIRS].[dbo].[HREngFrm001]");
                    sbSql.AppendFormat(@"  WHERE [HREngFrm001OutDate]='{0}' ", DateTime.Now.ToString("yyyy/MM/dd"));
                    sbSql.AppendFormat(@"  ORDER BY [HREngFrm001User],[HREngFrm001DefOutTime]");
                }
                else
                {
                    sbSql.AppendFormat(@"  SELECT ");
                    sbSql.AppendFormat(@"  [HREngFrm001User] AS '申請人',[HREngFrm001Rank] AS '職級',[HREngFrm001OutDate] AS '外出日期',[HREngFrm001Transp] AS '交通工具',[HREngFrm001LicPlate] AS '車牌',[HREngFrm001DefOutTime] AS '預計外出時間',[HREngFrm001OutTime] AS '實際外出時間',[HREngFrm001DefBakTime] AS '預計返廠時間',[HREngFrm001BakTime] AS '實際返廠時間'");
                    sbSql.AppendFormat(@"  ,[TaskId] AS 'TaskId',[HREngFrm001SN] AS '表單編號',[HREngFrm001Date] AS '申請日期',[HREngFrm001UsrDpt] AS '部門',[HREngFrm001Location] AS '外出地點',[HREngFrm001Agent] AS '代理人',[HREngFrm001Cause] AS '外出原因',[HREngFrm001FF] AS '是否由公司出發',[HREngFrm001CH] AS '是否回廠',[CRADNO] AS '卡號'");
                    sbSql.AppendFormat(@"  FROM [TKGAFFAIRS].[dbo].[HREngFrm001]");
                    sbSql.AppendFormat(@"  WHERE [HREngFrm001OutDate]='{0}' AND [CRADNO]='{1}'", DateTime.Now.ToString("yyyy/MM/dd"), CARDNO);
                    sbSql.AppendFormat(@"  ORDER BY [HREngFrm001DefOutTime]");
                }
                
                sbSql.AppendFormat(@"  ");
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

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow != null && !string.IsNullOrEmpty(textBox1.Text))
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];

                    string TaskId = row.Cells["TaskId"].Value.ToString();
                    string HREngFrm001User = row.Cells["申請人"].Value.ToString();
                    string HREngFrm001OutTime = row.Cells["實際外出時間"].Value.ToString();
                    string HREngFrm001BakTime = row.Cells["實際返廠時間"].Value.ToString();
                    string HREngFrm001FF = row.Cells["是否由公司出發"].Value.ToString();
                    string HREngFrm001CH = row.Cells["是否回廠"].Value.ToString();

                    if (!string.IsNullOrEmpty(textBox1.Text)&&STATUS.Equals("Y"))
                    {
                        CEHCK(TaskId, HREngFrm001User, HREngFrm001OutTime, HREngFrm001BakTime, HREngFrm001FF, HREngFrm001CH);
                    }


                    //if (STATUS.Equals("Y") && ds.Tables["TEMPds1"].Rows.Count == 1)
                    //{
                    //    CEHCK(TaskId, HREngFrm001User, HREngFrm001OutTime, HREngFrm001BakTime, HREngFrm001FF, HREngFrm001CH);
                    //}

                }
                else
                {


                }
            }
        }

        public void CEHCK(string TaskId, string HREngFrm001User, string HREngFrm001OutTime, string HREngFrm001BakTime, string HREngFrm001FF, string HREngFrm001CH)
        {
            if (STATUS.Equals("Y") && HREngFrm001FF.Equals("是") && string.IsNullOrEmpty(HREngFrm001OutTime))
            {
                INSERTHREngFrm001HREngFrm001OutTime(TaskId, HREngFrm001User, "實際外出時間");
                INSERTUOFHREngFrm001HREngFrm001OutTime(TaskId);

                STATUS = "N";

                if (!string.IsNullOrEmpty(textBox1.Text.Trim()))
                {
                    SEARCHHREngFrm001B(textBox1.Text.Trim());
                    //textBox1.Text = null;
                }

                //MessageBox.Show("實際外出時間"+ TaskId+" "+ HREngFrm001User);
            }
            else if (STATUS.Equals("Y") && !HREngFrm001FF.Equals("是") && HREngFrm001CH.Equals("是") && string.IsNullOrEmpty(HREngFrm001BakTime))
            {
                INSERTHREngFrm001HREngFrm001BakTime(TaskId, HREngFrm001User, "1實際返廠時間");
                INSERTUOFHREngFrm001HREngFrm001BakTime(TaskId);

                STATUS = "N";

                if (!string.IsNullOrEmpty(textBox1.Text.Trim()))
                {
                    SEARCHHREngFrm001(textBox1.Text.Trim());
                    //textBox1.Text = null;
                }


                //MessageBox.Show("1實際返廠時間" + TaskId + " " + HREngFrm001User);
            }
            else if (STATUS.Equals("Y") && HREngFrm001FF.Equals("是") && !string.IsNullOrEmpty(HREngFrm001OutTime) && HREngFrm001CH.Equals("是") && string.IsNullOrEmpty(HREngFrm001BakTime))
            {
                INSERTHREngFrm001HREngFrm001BakTime(TaskId, HREngFrm001User, "2實際返廠時間");
                INSERTUOFHREngFrm001HREngFrm001BakTime(TaskId);

                STATUS = "N";

                if (!string.IsNullOrEmpty(textBox1.Text.Trim()))
                {
                    SEARCHHREngFrm001(textBox1.Text.Trim());
                    //textBox1.Text = null;
                }


                //MessageBox.Show("2實際返廠時間" + TaskId + " " + HREngFrm001User);
            }




        }

        public void INSERTHREngFrm001HREngFrm001OutTime(string TaskId, string MODIFYUSR, string MODIFYCASUE)
        {
            if (!string.IsNullOrEmpty(TaskId))
            {
                UPDATEHREngFrm001HREngFrm001OutTime(TaskId, DateTime.Now.ToString("HH:mm"), MODIFYUSR, MODIFYCASUE);
            }
        }

        public void UPDATEHREngFrm001HREngFrm001OutTime(string TaskId, string HREngFrm001OutTime, string MODIFYUSR, string MODIFYCASUE)
        {
            try
            {

                //add ZWAREWHOUSEPURTH
                connectionString = connectionStringTKGAFFAIRS;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" UPDATE [TKGAFFAIRS].[dbo].[HREngFrm001]");
                sbSql.AppendFormat(" SET [HREngFrm001OutTime]='{0}',[MODIFYUSR]='{1}',[MODIFYCASUE]='{2}',[MODIFYTIME]='{3}'", HREngFrm001OutTime, MODIFYUSR, MODIFYCASUE, DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
                sbSql.AppendFormat(" WHERE TaskId='{0}'", TaskId);
                sbSql.AppendFormat(" ");

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


        public void INSERTHREngFrm001HREngFrm001BakTime(string TaskId, string MODIFYUSR, string MODIFYCASUE)
        {
            if (!string.IsNullOrEmpty(TaskId))
            {
                UPDATEHREngFrm001HREngFrm001BakTime(TaskId, DateTime.Now.ToString("HH:mm"), MODIFYUSR, MODIFYCASUE);
            }
        }

        public void UPDATEHREngFrm001HREngFrm001BakTime(string TaskId, string HREngFrm001OutTime, string MODIFYUSR, string MODIFYCASUE)
        {
            try
            {

                //add ZWAREWHOUSEPURTH
                connectionString = connectionStringTKGAFFAIRS;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" UPDATE [TKGAFFAIRS].[dbo].[HREngFrm001]");
                sbSql.AppendFormat(" SET [HREngFrm001BakTime]='{0}',[MODIFYUSR]='{1}',[MODIFYCASUE]='{2}',[MODIFYTIME]='{3}'", HREngFrm001OutTime, MODIFYUSR, MODIFYCASUE, DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
                sbSql.AppendFormat(" WHERE TaskId='{0}'", TaskId);
                sbSql.AppendFormat(" ");

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

        public void INSERTUOFHREngFrm001HREngFrm001OutTime(string TaskId)
        {
            DataSet ds = new DataSet();

            try
            {
                connectionString = connectionStringUOF;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  SELECT [TASK_ID],[TASK_SEQ],[BEGIN_TIME],[END_TIME],[TASK_STATUS],[TASK_RESULT],[DOC_NBR],[FLOW_TYPE],[FLOW_ID],[FORM_VERSION_ID],[SOURCE_DOC_ID],[CURRENT_DOC_ID],[FORM_STATUS],[USER_GUID],[USER_GROUP_ID],[USER_JOB_TITLE_ID],[ATTACH_ID],[URGENT_LEVEL],[CURRENT_SIGNER],[LOCK_STATUS],[CURRENT_DOC],[FILING_STATUS],[CURRENT_SITE_ID],[IS_APPLICANT_GETBACK],[APPLICANT_COMMENT],[DISPLAY_TITLE],[MESSAGE_CONTENT],[DEFAULT_IQY_USERS],[AGENT_USER],[CANCEL_FORM_REASON],[CANCEL_USER],[JSON_DISPLAY]");
                sbSql.AppendFormat(@"  FROM [{0}].[dbo].[TB_WKF_TASK]", DB);
                sbSql.AppendFormat(@"  WHERE TASK_ID='{0}'", TaskId);
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

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
                        XmlDocument Xmldoc = new XmlDocument();
                        Xmldoc.LoadXml(ds.Tables["ds"].Rows[0]["CURRENT_DOC"].ToString());

                        XmlNode node = Xmldoc.SelectSingleNode("Form/FormFieldValue/FieldItem[@fieldId='HREngFrm001OutTime']");
                        XmlElement element = (XmlElement)node;
                        element.SetAttribute("fieldValue", DateTime.Now.ToString("HH:mm"));

                        UPDATETUOFHREngFrm001HREngFrm001OutTime(TaskId, Xmldoc);
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

        public void UPDATETUOFHREngFrm001HREngFrm001OutTime(string TaskId, XmlDocument Xmldoc)
        {
            SqlCommand cmd = new SqlCommand();

            try
            {

                //add ZWAREWHOUSEPURTH
                connectionString = connectionStringUOF;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" UPDATE [{0}].[dbo].[TB_WKF_TASK]", DB);
                sbSql.AppendFormat(" SET  CURRENT_DOC=@CURRENT_DOC");
                sbSql.AppendFormat(" WHERE TASK_ID='{0}'", TaskId);
                sbSql.AppendFormat(" ");

                cmd.Parameters.AddWithValue("@CURRENT_DOC", Xmldoc.OuterXml);

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

        public void INSERTUOFHREngFrm001HREngFrm001BakTime(string TaskId)
        {
            DataSet ds = new DataSet();

            try
            {
                connectionString = connectionStringTKGAFFAIRS;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  SELECT [TASK_ID],[TASK_SEQ],[BEGIN_TIME],[END_TIME],[TASK_STATUS],[TASK_RESULT],[DOC_NBR],[FLOW_TYPE],[FLOW_ID],[FORM_VERSION_ID],[SOURCE_DOC_ID],[CURRENT_DOC_ID],[FORM_STATUS],[USER_GUID],[USER_GROUP_ID],[USER_JOB_TITLE_ID],[ATTACH_ID],[URGENT_LEVEL],[CURRENT_SIGNER],[LOCK_STATUS],[CURRENT_DOC],[FILING_STATUS],[CURRENT_SITE_ID],[IS_APPLICANT_GETBACK],[APPLICANT_COMMENT],[DISPLAY_TITLE],[MESSAGE_CONTENT],[DEFAULT_IQY_USERS],[AGENT_USER],[CANCEL_FORM_REASON],[CANCEL_USER],[JSON_DISPLAY]");
                sbSql.AppendFormat(@"  FROM [{0}].[dbo].[TB_WKF_TASK]", DB);
                sbSql.AppendFormat(@"  WHERE TASK_ID='{0}'", TaskId);
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

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
                        XmlDocument Xmldoc = new XmlDocument();
                        Xmldoc.LoadXml(ds.Tables["ds"].Rows[0]["CURRENT_DOC"].ToString());

                        XmlNode node = Xmldoc.SelectSingleNode("Form/FormFieldValue/FieldItem[@fieldId='HREngFrm001BakTime']");
                        XmlElement element = (XmlElement)node;
                        element.SetAttribute("fieldValue", DateTime.Now.ToString("HH:mm"));

                        UPDATETUOFHREngFrm001HREngFrm001OutTime(TaskId, Xmldoc);
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

        public void UPDATETUOFHREngFrm001HREngFrm001BakTime(string TaskId, XmlDocument Xmldoc)
        {
            SqlCommand cmd = new SqlCommand();

            try
            {

                //add ZWAREWHOUSEPURTH
                connectionString = connectionStringUOF;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" UPDATE [{0}].[dbo].[TB_WKF_TASK]", DB);
                sbSql.AppendFormat(" SET  CURRENT_DOC=@CURRENT_DOC");
                sbSql.AppendFormat(" WHERE TASK_ID='{0}'", TaskId);
                sbSql.AppendFormat(" ");

                cmd.Parameters.AddWithValue("@CURRENT_DOC", Xmldoc.OuterXml);

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

        public void CHECKWHITELIST()
        {
            if(!string.IsNullOrEmpty(textBox1.Text))
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

                    sbSql.AppendFormat(@"  SELECT [ID],[CARDNO],[NAME] FROM [TKGAFFAIRS].[dbo].[WHITELIST]");
                    sbSql.AppendFormat(@"  ");
                    sbSql.AppendFormat(@"  ");
                    sbSql.AppendFormat(@"  ");
       

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
                                if(dr["CARDNO"].ToString().Trim().Equals(textBox1.Text.Trim()))
                                {
                                    ADDTOHREngFrm001(dr["ID"].ToString().Trim(), dr["CARDNO"].ToString().Trim(), dr["NAME"].ToString().Trim());
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
            }
        }

        public void ADDTOHREngFrm001(string ID,string CARDNO,string NAME)
        {
            SqlCommand cmd = new SqlCommand();

            try
            {
            
                connectionString = connectionStringTKGAFFAIRS;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" INSERT INTO [TKGAFFAIRS].[dbo].[HREngFrm001]");
                sbSql.AppendFormat(" ([TaskId],[HREngFrm001SN],[HREngFrm001Date],[HREngFrm001User],[HREngFrm001UsrDpt],[HREngFrm001Rank],[HREngFrm001OutDate],[HREngFrm001Location],[HREngFrm001Agent],[HREngFrm001Transp],[HREngFrm001LicPlate],[HREngFrm001Cause],[HREngFrm001DefOutTime],[HREngFrm001FF],[HREngFrm001OutTime],[HREngFrm001DefBakTime],[HREngFrm001CH],[HREngFrm001BakTime],[CRADNO],[MODIFYUSR],[MODIFYCASUE],[MODIFYTIME])");
                sbSql.AppendFormat(" VALUES");
                sbSql.AppendFormat(" (@TaskId,@HREngFrm001SN,@HREngFrm001Date,@HREngFrm001User,@HREngFrm001UsrDpt,@HREngFrm001Rank,@HREngFrm001OutDate,@HREngFrm001Location,@HREngFrm001Agent,@HREngFrm001Transp,@HREngFrm001LicPlate,@HREngFrm001Cause,@HREngFrm001DefOutTime,@HREngFrm001FF,@HREngFrm001OutTime,@HREngFrm001DefBakTime,@HREngFrm001CH,@HREngFrm001BakTime,@CRADNO,@MODIFYUSR,@MODIFYCASUE,@MODIFYTIME)");
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" ");

                cmd.Parameters.AddWithValue("@TaskId", Guid.NewGuid());
                cmd.Parameters.AddWithValue("@HREngFrm001SN", "");
                cmd.Parameters.AddWithValue("@HREngFrm001Date",DateTime.Now.ToString("yyyy/MM/dd") );
                cmd.Parameters.AddWithValue("@HREngFrm001User", NAME+ ID);
                cmd.Parameters.AddWithValue("@HREngFrm001UsrDpt","" );
                cmd.Parameters.AddWithValue("@HREngFrm001Rank","" );
                cmd.Parameters.AddWithValue("@HREngFrm001OutDate", DateTime.Now.ToString("yyyy/MM/dd"));
                cmd.Parameters.AddWithValue("@HREngFrm001Location", "");
                cmd.Parameters.AddWithValue("@HREngFrm001Agent", "");
                cmd.Parameters.AddWithValue("@HREngFrm001Transp", "");
                cmd.Parameters.AddWithValue("@HREngFrm001LicPlate","" );
                cmd.Parameters.AddWithValue("@HREngFrm001Cause", "可自由外出人員" );
                cmd.Parameters.AddWithValue("@HREngFrm001DefOutTime", DateTime.Now.ToString("HH:mm"));
                cmd.Parameters.AddWithValue("@HREngFrm001FF", "否");
                cmd.Parameters.AddWithValue("@HREngFrm001OutTime", DateTime.Now.ToString("HH:mm"));
                cmd.Parameters.AddWithValue("@HREngFrm001DefBakTime", DateTime.Now.ToString("HH:mm"));
                cmd.Parameters.AddWithValue("@HREngFrm001CH", "否");
                cmd.Parameters.AddWithValue("@HREngFrm001BakTime", DateTime.Now.ToString("HH:mm"));
                cmd.Parameters.AddWithValue("@CRADNO",CARDNO );
                cmd.Parameters.AddWithValue("@MODIFYUSR", NAME + ID);
                cmd.Parameters.AddWithValue("@MODIFYCASUE", "");
                cmd.Parameters.AddWithValue("@MODIFYTIME", DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));

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

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SEARCHHREngFrm001B("");
        }

        #endregion


    }
}
