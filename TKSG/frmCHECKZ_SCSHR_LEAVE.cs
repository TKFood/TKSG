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
    public partial class frmCHECKZ_SCSHR_LEAVE : Form
    {

        String connectionStringTKSG = "server=192.168.1.105;database=TKSG;uid=sa;pwd=dsc";
        String connectionStringTKGAFFAIRS = "server=192.168.1.105;database=TKGAFFAIRS;uid=sa;pwd=dsc";
        String connectionStringUOF = "server=192.168.1.223;database=UOF;uid=TKUOF;pwd=TKUOF123456";

        string DB = "UOF";

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


        //用STATUS來控制在1分鐘內不得連續刷卡
        string STATUS = "Y";

        public frmCHECKZ_SCSHR_LEAVE()
        {
            InitializeComponent();
        }


        #region FUNCTION

        #endregion

        #region BUTTON
        private void button2_Click(object sender, EventArgs e)
        {

        }
        #endregion
    }
}
