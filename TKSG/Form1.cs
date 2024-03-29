﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Configuration;
using System.Reflection;
using System.Net;
using System.Net.Sockets;

namespace TKSG
{
    public partial class Form1 : Form
    {
        String connectionString = "server=192.168.1.105;database=TKSG;uid=sa;pwd=dsc";
        public Form1()
        {
            InitializeComponent();
        }

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            LOGIN();
        }

        #endregion

        #region LOGIN
        public void LOGIN()
        {
            if (txt_UserName.Text == "" || txt_Password.Text == "")
            {
                MessageBox.Show("請輸入帳號、密碼");
                return;
            }
            try
            {
                //Create SqlConnection
              
                SqlConnection conn;
               
                conn = new SqlConnection(connectionString);
                SqlCommand cmd = new SqlCommand("Select * from MNU_Login where UserName=@username and Password=@password", conn);
                cmd.Parameters.AddWithValue("@username", txt_UserName.Text);
                cmd.Parameters.AddWithValue("@password", txt_Password.Text);
                conn.Open();
                SqlDataAdapter adapt = new SqlDataAdapter(cmd);
                DataSet ds = new DataSet();
                adapt.Fill(ds);
                conn.Close();
                int count = ds.Tables[0].Rows.Count;
                //If count is equal to 1, than show frmMain form
                if (count == 1)
                {
                    //ADD USED LOG
                    List<string> IPAddress = GetHostIPAddress();
                    //MessageBox.Show(IPAddress[0].ToString());    
                    ADDTKSYSLOGIN(MethodBase.GetCurrentMethod().DeclaringType.Namespace, txt_UserName.Text.Trim(), IPAddress[0].ToString(), "SUCCESS");

                    //MessageBox.Show("登入成功!");

                    FrmParent fm = new FrmParent(txt_UserName.Text.ToString());
                    fm.Show();
                    this.Hide();
                }
                else
                {
                    //ADD USED LOG
                    List<string> IPAddress = GetHostIPAddress();
                    //MessageBox.Show(IPAddress[0].ToString());    
                    ADDTKSYSLOGIN(MethodBase.GetCurrentMethod().DeclaringType.Namespace, txt_UserName.Text.Trim(), IPAddress[0].ToString(), "FAIL");

                    MessageBox.Show("登入失敗!");

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        public void ADDTKSYSLOGIN(string SYSTEMNAME, string USEDID, string USEDIP, string LOGIN)
        {
            SqlConnection sqlConn = new SqlConnection();
            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();
            int result;
            StringBuilder sbSql = new StringBuilder();


            sqlConn = new SqlConnection(connectionString);


            sqlConn.Close();
            sqlConn.Open();
            tran = sqlConn.BeginTransaction();

            sbSql.Clear();



            sbSql.AppendFormat(@" 
                                INSERT INTO [TKIT].[dbo].[TKSYSLOGIN]
                                ([SYSTEMNAME],[USEDDATES],[USEDID],[USEDIP],[LOGIN])
                                VALUES
                                (@SYSTEMNAME,@USEDDATES,@USEDID,@USEDIP,@LOGIN)
                                ");


            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                SqlCommand command = new SqlCommand(sbSql.ToString(), connection);
                command.Parameters.AddWithValue("@SYSTEMNAME", SYSTEMNAME);
                command.Parameters.AddWithValue("@USEDDATES", DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
                command.Parameters.AddWithValue("@USEDID", USEDID);
                command.Parameters.AddWithValue("@USEDIP", USEDIP);
                command.Parameters.AddWithValue("@LOGIN", LOGIN);
                try
                {
                    connection.Open();
                    Int32 rowsAffected = command.ExecuteNonQuery();

                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }

                finally
                {
                    sqlConn.Close();
                }
            }


        }

        // <summary>
        /// 取得本機 IP Address
        /// </summary>
        /// <returns></returns>
        private List<string> GetHostIPAddress()
        {
            List<string> lstIPAddress = new List<string>();
            IPHostEntry IpEntry = Dns.GetHostEntry(Dns.GetHostName());
            foreach (IPAddress ipa in IpEntry.AddressList)
            {
                if (ipa.AddressFamily == AddressFamily.InterNetwork)
                {
                    lstIPAddress.Add(ipa.ToString());
                    //MessageBox.Show(ipa.ToString());
                }

            }
            return lstIPAddress; // result: 192.168.1.17 ......
        }
        #endregion

        #region FUNCTION
        private void txt_Password_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                LOGIN();
            }
        }

        private void txt_UserName_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txt_Password.Focus();
            }
        }

        #endregion
    }
}
