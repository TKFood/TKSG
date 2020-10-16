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
    public partial class FrmUOFGAFrm : Form
    {
        String connectionStringTKSG = "server=192.168.1.105;database=TKSG;uid=sa;pwd=dsc";
        String connectionStringTKGAFFAIRS = "server=192.168.1.105;database=TKGAFFAIRS;uid=sa;pwd=dsc";
        String connectionStringUOF = "server=192.168.1.223;database=UOF;uid=TKUOF;pwd=TKUOF123456";


        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
    

        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds1 = new DataSet();


        int result;
        Thread TD;

        public FrmUOFGAFrm()
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

            SEARCHUOFGAFrm(DateTime.Now.ToString("yyyy/MM/dd"));
        }

        public void SEARCHUOFGAFrm(string GAFrm004OD)
        {

            try
            {
                connectionString = connectionStringTKGAFFAIRS;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                StringBuilder query = new StringBuilder();

                sbSql.AppendFormat(@"  
                                    SELECT 
                                    [GAFrm004SN] AS '單號'
                                    ,[GAFrm004OD] AS '日期'
                                    ,[GAFrm004SI] AS '申請人'
                                    ,[GAFrm004CM] AS '資產編號'
                                    ,[GAFrm004DN] AS '設備名稱'
                                    ,[GAFrm004NB] AS '設備數量'
                                    ,[GAFrm004ID] AS '保管部門'
                                    ,[GAFrm004ER] AS '異常原呌'
                                    ,[GAFrm004S0ND] AS '原因及處理'
                                    ,[GAFrm004PS] AS '是否出廠'
                                    ,[GAFrm004PID] AS '外送時間'
                                    ,[GAFrm004RD] AS '回廠時間'
                                    ,[TaskId] 
                                    ,[SERNO] 

                                    FROM [TKGAFFAIRS].[dbo].[UOFGAFrm]
                                    WHERE [GAFrm004OD]='{0}'
                                    ORDER  BY [GAFrm004SN],[SERNO] 
                                    ", GAFrm004OD);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count == 0)
                {
                    dataGridView1.DataSource = null;

                }
                else
                {
                    if (ds1.Tables["ds1"].Rows.Count >= 1)
                    {
                        dataGridView1.DataSource = ds1.Tables["ds1"];
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
            SEARCHUOFGAFrm(DateTime.Now.ToString("yyyy/MM/dd"));
        }
        #endregion

       
    }
}
