using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExclHelp
{
    public partial class Form1 : Form
    {

        public DataTable resultOne = new DataTable();
        public DataTable resultTwo = new DataTable();

        public List<string> columnsNameOne = new List<string>();
        public List<string> columnsNameTwo = new List<string>();

        public List<string> SheetNameOne = new List<string>();
        public List<string> SheetNameTwo = new List<string>();
        public string sheetNameOne = string.Empty;
        public string sheetNameTwo = string.Empty;

        public DataTable result = new DataTable();

        public string fileNameOne = string.Empty;
        public string fileNameTwo = string.Empty;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.pb.Value = 0;
            this.pb.Minimum = 0;
            this.pb.Maximum = 100;
        }

        private void btn_FileChoseOne_Click(object sender, EventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Multiselect = true;
            fileDialog.Title = "请选择文件";
            fileDialog.Filter = "所有文件(*.*)|*.*";
            string fileName = string.Empty;
            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                string[] names = fileDialog.FileNames;

                foreach (string file in names)
                {
                    fileName = file;
                }
                if (!string.IsNullOrWhiteSpace(fileName))
                {
                    fileNameOne = fileName;
                    this.txt_FileChoseone.Text = fileName;
                    this.pb.Visible = true;
                    Task a = new Task(ReadExcelToTable,1);
                    a.Start();
                    while (!a.IsCompleted)
                    {
                        this.pb.Value = 0;
                        int ab=0;
                        for (int i = 0; i < 100; i++)
                        {
                            this.pb.Increment(ab++);
                            Thread.Sleep(10);
                        }
                    }
                    a.Wait();
                    this.pb.Visible = false;
                    this.dt_result.DataSource = resultOne;
                    this.cb_sheetNameOne.DataSource = SheetNameOne;
                    this.cb_sheetNameOne.DisplayMember = string.Join("\n", SheetNameOne);
                }

            }

        }

        private void btn_FileChoseTwo_Click(object sender, EventArgs e)
        {
            List<string> columnsNameList = new List<string>();
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Multiselect = true;
            fileDialog.Title = "请选择文件";
            fileDialog.Filter = "所有文件(*.*)|*.*";
            string fileName = string.Empty;
            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                string[] names = fileDialog.FileNames;

                foreach (string file in names)
                {
                    fileName = file;
                }
                if (!string.IsNullOrWhiteSpace(fileName))
                {
                    fileNameTwo = fileName;
                    this.txt_FileChoseTwo.Text = fileNameTwo;
                    this.pb.Visible = true;
                    Task a = new Task(ReadExcelToTable, 2);
                    a.Start();
                    while (!a.IsCompleted)
                    {
                        this.pb.Value = 0;
                        int ab = 0;
                        for (int i = 0; i < 100; i++)
                        {
                            this.pb.Increment(ab++);
                            Thread.Sleep(10);
                        }
                    }
                    a.Wait();
                    this.pb.Visible = false;
                    this.dt_result.DataSource = resultTwo;
                    this.cb_sheetNameTwo.DataSource = SheetNameTwo;
                    this.cb_sheetNameTwo.DisplayMember = string.Join("\n", SheetNameTwo);
                }
            }
        }


        private void  ReadExcelToTable(object param)
        {
            string path = string.Empty;
            string firstSheetName = string.Empty;
            if (param.ToString() == "1")
            {
                path=fileNameOne;
                try
                {
                    firstSheetName = sheetNameOne;
                }
                catch (Exception ex)
                { }
            }
            else
            {
                path = fileNameTwo;
                try
                {
                    firstSheetName = sheetNameTwo;
                }
                catch (Exception ex)
                { }
            }
            //连接字符串
            string connstring = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 8.0;HDR=NO;IMEX=1';"; // Office 07及以上版本 不能出现多余的空格 而且分号注意
            //string connstring = Provider=Microsoft.JET.OLEDB.4.0;Data Source=" + path + ";Extended Properties='Excel 8.0;HDR=NO;IMEX=1';"; //Office 07以下版本 因为本人用Office2010 所以没有用到这个连接字符串 可根据自己的情况选择 或者程序判断要用哪一个连接字符串
           List<string> columnsName = new List<string>();
           List<string> sheetName = new List<string>();
            using (OleDbConnection conn = new OleDbConnection(connstring))
            {
                conn.Close();
                conn.Open();
                System.Data.DataTable sheetsName = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "Table" }); //得到所有sheet的名字
                if (String.IsNullOrWhiteSpace(firstSheetName))
                {
                    firstSheetName = sheetsName.Rows[0][2].ToString(); //得到第一个sheet的名字
                    foreach (DataRow dr in sheetsName.Rows)
                    {
                        sheetName.Add(dr[2].ToString());
                    }
                }
                string sql = string.Format("SELECT * FROM [{0}]", firstSheetName); //查询字符串
                OleDbDataAdapter ada = new OleDbDataAdapter(sql, connstring);
                DataSet set = new DataSet();
                ada.Fill(set);
                System.Data.DataTable dt = set.Tables[0];
                if (dt != null && dt.Columns.Count > 0)
                {
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        columnsName.Add(dt.Rows[0][i].ToString());
                    }
                }
                try
                {
                    if (param.ToString() == "1")
                    {
                        resultOne = dt;
                        columnsNameOne = columnsName;
                        if (sheetName.Any())
                        {
                            SheetNameOne = sheetName;
                        }
                    }
                    else
                    {
                        resultTwo = dt;
                        columnsNameTwo = columnsName;
                        if (sheetName.Any())
                        {
                            SheetNameTwo = sheetName;
                        }
                    }
                }
                catch (Exception ex)
                { 
                
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.pb.Visible = true;
            result = new DataTable("yc");
            DataColumn dc;
            DataRow dr;
            int a = Convert.ToInt32(this.txt_indexOne.Text.ToString().Trim());
            int b= Convert.ToInt32(this.txt_indexTwo.Text.ToString().Trim());
            if (columnsNameOne.Any() && columnsNameTwo.Any())
            {
                for (int i = 0; i < columnsNameOne.Count; i++)
                {
                    dc = new DataColumn(columnsNameOne[i] + "One", Type.GetType("System.String"));
                    result.Columns.Add(dc);
                }
                for (int i = 0; i < columnsNameTwo.Count; i++)
                {
                    dc = new DataColumn(columnsNameTwo[i] + "Two", Type.GetType("System.String"));
                    result.Columns.Add(dc);
                }
            }
            if (resultOne != null && resultOne.Rows.Count > 0 && resultTwo != null && resultTwo.Rows.Count > 0)
            {
                try
                {
                    this.pb.Maximum = resultOne.Rows.Count-1;
                    for (int i = 1; i < resultOne.Rows.Count; i++)
                    {
                        for (int j = 1; j < resultTwo.Rows.Count; j++)
                        {
                            if (resultOne.Rows[i][a].ToString().Trim() == resultTwo.Rows[j][b].ToString().Trim())
                            {
                                dr = result.NewRow();
                                for (int m = 0; m < resultOne.Columns.Count; m++)
                                {
                                    dr[columnsNameOne[m] + "One"] = resultOne.Rows[i][m];
                                }
                                for (int m = 0; m < resultTwo.Columns.Count; m++)
                                {
                                    dr[columnsNameTwo[m] + "Two"] = resultTwo.Rows[j][m];
                                }
                                result.Rows.Add(dr);
                                break;
                            }
                            else
                            {
                                dr = result.NewRow();
                                for (int m = 0; m < resultOne.Columns.Count; m++)
                                {
                                    dr[columnsNameOne[m] + "One"] = resultOne.Rows[i][m];
                                }
                            }
                          
                        }
                        this.pb.Value = i;
                        Thread.Sleep(100);
                    }
                }
                catch (Exception ex)
                {
                    var ss = result;
                }
            }
            this.dt_result.DataSource = result;
            this.pb.Visible = false;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DataSet ds = new DataSet();
            ds.Tables.Add(result);
            ToExcel(this.dt_result);
        }

        public void ToExcel(DataGridView dataGridView1)
        {
            try
            {
                //没有数据的话就不往下执行  
                if (dataGridView1.Rows.Count == 0)
                    return;
                //实例化一个Excel.Application对象  
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                //让后台执行设置为不可见，为true的话会看到打开一个Excel，然后数据在往里写  
                excel.Visible = true;
                //新增加一个工作簿，Workbook是直接保存，不会弹出保存对话框，加上Application会弹出保存对话框，值为false会报错  
                excel.Application.Workbooks.Add(true);
                //生成Excel中列头名称  
                for (int i = 0; i < dataGridView1.Columns.Count; i++)
                {
                    if (this.dt_result.Columns[i].Visible == true)
                    {
                        excel.Cells[1, i + 1] = dataGridView1.Columns[i].HeaderText;
                    }

                }
                //把DataGridView当前页的数据保存在Excel中  
                for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                {
                    System.Windows.Forms.Application.DoEvents();
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        if (this.dt_result.Columns[j].Visible == true)
                        {
                            if (dataGridView1[j, i].ValueType == typeof(string))
                            {
                                excel.Cells[i + 2, j + 1] = "'" + dataGridView1[j, i].Value.ToString();
                            }
                            else
                            {
                                excel.Cells[i + 2, j + 1] = dataGridView1[j, i].Value.ToString();
                            }
                        }

                    }
                }
                //设置禁止弹出保存和覆盖的询问提示框  
                excel.DisplayAlerts = true;
                excel.AlertBeforeOverwriting = true;
                //保存工作簿  
                excel.Application.Workbooks.Add(true).Save();
                //保存excel文件  
                excel.Save("D:" + "\\yc" + DateTime.Now.Ticks + ".xls");
                //确保Excel进程关闭  
                excel.Quit();
                excel = null;
                GC.Collect();//如果不使用这条语句会导致excel进程无法正常退出，使用后正常退出
                MessageBox.Show(this, "文件已经成功导出！", "信息提示");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误提示");
            }
        }

        private void cb_sheetNameOne_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.pb.Visible = true;
            sheetNameOne = this.cb_sheetNameOne.SelectedValue.ToString();
            Task a = new Task(ReadExcelToTable, 1);
            a.Start();
            while (!a.IsCompleted)
            {
                this.pb.Value = 0;
                int ab = 0;
                for (int i = 0; i < 100; i++)
                {
                    this.pb.Increment(ab++);
                    Thread.Sleep(10);
                }
            }
            a.Wait();
            this.pb.Visible = false;
            this.dt_result.DataSource = resultOne;
        }

        private void cb_sheetNameTwo_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.pb.Visible = true;
            sheetNameTwo = this.cb_sheetNameTwo.SelectedValue.ToString();
            Task a = new Task(ReadExcelToTable,2);
            a.Start();
            while (!a.IsCompleted)
            {
                this.pb.Value = 0;
                int ab = 0;
                for (int i = 0; i < 100; i++)
                {
                    this.pb.Increment(ab++);
                    Thread.Sleep(10);
                }
            }
            a.Wait();
            this.pb.Visible = false;
            this.dt_result.DataSource = resultTwo;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            result=new DataTable();
            this.pb.Visible = true;
             DataRow dr;
             DataColumn dc;
             int a = Convert.ToInt32(this.txt_indexOne.Text.ToString().Trim());
            //连接字符串
            string connstring = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=./Area.xls;Extended Properties='Excel 8.0;HDR=NO;IMEX=1';"; // Office 07及以上版本 不能出现多余的空格 而且分号注意

            List<string> columnsName = new List<string>();
            List<string> sheetName = new List<string>();
            using (OleDbConnection conn = new OleDbConnection(connstring))
            {
                conn.Close();
                conn.Open();
                System.Data.DataTable sheetsName = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "Table" }); //得到所有sheet的名字
                string sql = string.Format("SELECT * FROM [{0}]", sheetsName.Rows[0][2].ToString()); //查询字符串
                OleDbDataAdapter ada = new OleDbDataAdapter(sql, connstring);
                DataSet set = new DataSet();
                ada.Fill(set);
                System.Data.DataTable dt = set.Tables[0];
                if (dt != null && dt.Columns.Count > 0)
                {
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        columnsName.Add(dt.Rows[0][i].ToString());
                    }
                }

                //初始化列明
                if (columnsNameOne.Any())
                {
                    for (int i = 0; i < columnsNameOne.Count; i++)
                    {
                        dc = new DataColumn(columnsNameOne[i], Type.GetType("System.String"));
                        result.Columns.Add(dc);
                    }
                     dc = new DataColumn("省份", Type.GetType("System.String"));
                     result.Columns.Add(dc);
                }

                bool flag = false;
                //根据城市名称进行排序
                if (resultOne != null && resultOne.Rows.Count > 0)
                {
                    try
                    {
                        this.pb.Maximum = resultOne.Rows.Count - 1;
                        for (int i = 1; i < resultOne.Rows.Count; i++)
                        {
                            this.pb.Maximum = resultOne.Rows.Count - 1;
                            //DataRow GetRows = dt.Select("name like '%" + resultOne.Rows[i][a].ToString().Trim() + "'% or shortCnName like '%" + resultOne.Rows[i][a].ToString().Trim() + "%'").FirstOrDefault();
                            //if (GetRows != null)
                            //{
                            //    dr = result.NewRow();
                            //    for (int m = 0; m < resultOne.Columns.Count; m++)
                            //    {
                            //        dr[columnsNameOne[m]] = resultOne.Rows[i][m];
                            //    }
                            // //   dr["省份"] = GetRows..Rows[j][0].ToString().Trim().Split(',')[2];
                            //    result.Rows.Add(dr);
                            //    break;
                            //}
                            for (int j = 1; j < dt.Rows.Count; j++)
                            {
                                flag = false;
                                if (resultOne.Rows[i][a].ToString().Trim().Contains(dt.Rows[j][1].ToString().Trim()) || resultOne.Rows[i][a].ToString().Trim().Contains(dt.Rows[j][2].ToString().Trim()))
                                {
                                    dr = result.NewRow();
                                    for (int m = 0; m < resultOne.Columns.Count; m++)
                                    {
                                        dr[columnsNameOne[m]] = resultOne.Rows[i][m];
                                    }
                                    dr["省份"] = dt.Rows[j][0].ToString().Trim().Split(',')[2];
                                    result.Rows.Add(dr);
                                    flag = true;
                                    break;
                                }
                            }
                            if (!flag)
                            {
                                dr = result.NewRow();
                                for (int m = 0; m < resultOne.Columns.Count; m++)
                                {
                                    dr[columnsNameOne[m]] = resultOne.Rows[i][m];
                                }
                                dr["省份"] = string.Empty;
                                result.Rows.Add(dr);
                            }
                            this.pb.Value = i;
                          //  Thread.Sleep(100);
                        }
                    }
                    catch (Exception ex)
                    {
                        var ss = result;
                    }
                }
               
                this.dt_result.DataSource = result;
                this.pb.Visible = false;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
           
        }
    }




    public class ModelConvertHelper<T> where T : new()  // 此处一定要加上new()
    {

        public static IList<T> ConvertToModel(DataTable dt)
        {

            IList<T> ts = new List<T>();// 定义集合
            Type type = typeof(T); // 获得此模型的类型
            string tempName = "";
            foreach (DataRow dr in dt.Rows)
            {
                T t = new T();
                PropertyInfo[] propertys = t.GetType().GetProperties();// 获得此模型的公共属性
                foreach (PropertyInfo pi in propertys)
                {
                    tempName = pi.Name;
                    if (dt.Columns.Contains(tempName))
                    {
                        if (!pi.CanWrite) continue;
                        object value = dr[tempName];
                        if (value != DBNull.Value)
                            pi.SetValue(t, value, null);
                    }
                }
                ts.Add(t);
            }
            return ts;
        }
    }
}
