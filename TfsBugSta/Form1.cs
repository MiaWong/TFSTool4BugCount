using System;
using System.IO;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using Microsoft.TeamFoundation.ProcessConfiguration.Client;
//using Microsoft.Office.Interop.Excel;
//using Excel = Microsoft.Office.Interop.Excel;
using Aspose.Cells;



namespace TfsBugSta
{

    public partial class Form1 : Form
    {
        #region
        private WorkItemStore workstore;
            
        private TfsTeamProjectCollection server;

        private TeamSettingsConfigurationService configSvc;

        private TfsTeamService teamService;

        private WorkItemCollection queryResults;

        /// <summary>
        /// 查询条件-产品名称
        /// </summary>
        private string productname = string.Empty;


        /// <summary>
        /// 定时器
        /// </summary>
        System.Timers.Timer tmSeconds;

        /// <summary>
        /// 秒表
        /// </summary>
        private int Seconds = 0;
        #endregion

        public String TfsUri { get; set; }


        public Form1()
        {
            #region
            InitializeComponent();
            //跨线程调用窗体
            Control.CheckForIllegalCrossThreadCalls = false;
            #endregion
        }



        /// <summary>
        /// 初始化TFSServer
        /// </summary>
        /// <param name="model"></param>
        public void TFSServerDto(string tfsUri)
        {
            #region
            try
            {
                this.TfsUri = tfsUri;
                Uri uri = new Uri(TfsUri);
                server = new TfsTeamProjectCollection(uri);
                workstore = (WorkItemStore)server.GetService(typeof(WorkItemStore));
                configSvc = server.GetService<TeamSettingsConfigurationService>();
                teamService = server.GetService<TfsTeamService>();
                if (string.IsNullOrEmpty(workstore.ToString()))
                {
                    MessageBox.Show("获取团队失败");
                    this.tmSeconds.Stop();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("ConnectServer Fail-- "+ex.Message.ToString());
            }
            #endregion
        }


        /// <summary>
        /// 获取项目集合
        /// </summary>
        /// <returns></returns>
        public ProjectCollection GetProjectList()
        {
            #region
            return workstore.Projects;
            #endregion
        }

        /// <summary>
        /// 获取项目ID
        /// </summary>
        /// <returns></returns>
        public Project GetProject(int projectId)
        {
            #region
            return workstore.Projects.GetById(projectId);
            #endregion
        }


        /// <summary>
        /// 获取模块
        /// </summary>
         private void GetWorkItemCollectionforlevel(string proname)
        {
            #region    
        

           string condition1 = textBox2.Text;

           string condition2 = textBox3.Text;

           string condition3 = textBox4.Text;

           string condition4 = textBox5.Text;

           string condition5 = textBox7.Text;

           string condition6 = textBox8.Text;

           string condition7 = textBox9.Text;

           string condition8 = textBox10.Text;

            string sql = @"Select [System.Title] From WorkItems Where [System.WorkItemType] ='Bug'  and [System.Title]  contains '{0}'";
            sql = string.Format(sql, proname);

            try
            {
                if (!string.IsNullOrEmpty(condition1))
                {
                    BugCountforlevel(sql, condition1, 1);
                }
                if (!string.IsNullOrEmpty(condition2))
                {
                    BugCountforlevel(sql, condition2, 1);
                }
                if (!string.IsNullOrEmpty(condition3))
                {
                    BugCountforlevel(sql, condition3, 1);
                }
                if (!string.IsNullOrEmpty(condition4))
                {
                    BugCountforlevel(sql, condition4, 1);
                }
                if (!string.IsNullOrEmpty(condition5))
                {
                    BugCountforlevel(sql, condition5, 1);
                }
                if (!string.IsNullOrEmpty(condition6))
                {
                    BugCountforlevel(sql, condition6, 1);
                }
                if (!string.IsNullOrEmpty(condition7))
                {
                    BugCountforlevel(sql, condition7, 1);
                }
                if (!string.IsNullOrEmpty(condition8))
                {
                    BugCountforlevel(sql, condition8, 1);
                }
                BugCountforlevel(sql,string.Empty, 0);
            }
             catch(Exception ex)
            {
                MessageBox.Show("Query Fail-- " + ex.Message.ToString());
             }
            #endregion
        }


         /// <summary>
         /// sql按模块条件查询
         /// </summary>
         private void BugCountforlevel(string sql, string condition, int number)
         {
             int currentcount = 0;

             int countotal = 0;

             string execsql = string.Empty;

             ListViewItem item = new ListViewItem();

             string[] arr = condition.Split(new char[] { ';', ' ', '；', '-' }, StringSplitOptions.RemoveEmptyEntries);

             if (number==0)
             {
                 item.Text = "总计";
             }
             else
             {
                 item.Text = condition;
             }

             //按条件补全sql
             for (int i = 0; i < arr.Length ;i++ )
             {
                 sql += "and [System.Title] contains '" + arr[i] + "'";
             }

             //按bug的4个等级查询bug数
             for (int lev = 1; lev < 5; lev++)
             {
                 execsql = sql+ "and [Microsoft.VSTS.Common.Priority] = '" + lev.ToString() + "'";
                 try
                 {
                     queryResults = workstore.Query(execsql);
                     currentcount = queryResults.Count;
                     countotal += currentcount;
                     item.SubItems.Add(Convert.ToString(currentcount));
                 }
                 catch (Exception ex)
                 {
                     MessageBox.Show("sql exec fail-- " + ex.Message.ToString());
                 }
             }
             item.SubItems.Add(Convert.ToString(countotal));
             listView1.Items.Add(item);
         }


         /// <summary>
         /// sql按BUG等级查询
         /// </summary>
         private void BugQueryforlevel()
         {
             #region
             productname = textBox1.Text;

             Seconds = 0;

             if (!string.IsNullOrEmpty(productname))
             {
                 //启动定时器显示秒表
                 button1.Enabled = false;
                 tmSeconds = new System.Timers.Timer();
                 tmSeconds.Elapsed += TimerTick_Elapsed;
                 tmSeconds.Interval = 1000;
                 tmSeconds.Enabled = true;

                 listView1.Items.Clear();
                 TFSServerDto("http://tfs2018-web.winning.com.cn:8080/tfs/WN_HIS");

                 GetWorkItemCollectionforlevel(productname);

                 WriteExcel(1, listView1);      
                 this.tmSeconds.Stop();
                 button1.Enabled = true;
             }
             else
             {
                 MessageBox.Show("请输入产品名称！");
             }
             #endregion
         }


        /// <summary>
        /// 按输入模块条件查询线程
        /// </summary>
        private void BtClick(object sender, EventArgs e)
         {
             #region
             new Thread(new ThreadStart(this.BugQueryforlevel)).Start();
             #endregion
         }


        /// <summary>
        /// 计时器
        /// </summary>
        private void TimerTick_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            #region
            try
            {
                if (Seconds >= Int32.MaxValue)
                {
                    Seconds = 0;
                }
                Seconds++;
                this.Invoke(new MethodInvoker(delegate()
                {
                    label6.Text = Seconds.ToString();
                    label9.Text = Seconds.ToString();
                }));
            }
            catch (Exception ex)
            {
                MessageBox.Show("TimerTick_Elapsed-" + ex.Message.ToString());
            }
            #endregion
        }


        /// <summary>
        /// 重置BUG统计表页面
        /// </summary>
        private void ResttabControl1(object sender, EventArgs e)
        {
            #region 
            textBox1.Text = string.Empty;
            textBox2.Text = string.Empty;
            textBox3.Text = string.Empty;
            textBox4.Text = string.Empty;
            listView1.Items.Clear();
            #endregion 
        }


        /// <summary>
        /// 重置BUG增长趋势表页面
        /// </summary>
        private void ResttabControl2(object sender, EventArgs e)
        {
            #region
            textBox6.Text = string.Empty;
            listView2.Items.Clear();
            #endregion
        }


        /// <summary>
        /// 查询结果写excel
        /// </summary>
        public static void WriteExcel( int number ,ListView listview)
        {
            #region 
            string str;

            try
            {
                string path =string.Empty;
                if (number == 1)
                {
                    path = System.Windows.Forms.Application.StartupPath + "\\Data\\Bug模块查询";
                }
                else
                {
                    path = System.Windows.Forms.Application.StartupPath + "\\Data\\Bug增长你趋势查询";
                }
                string sname = "\\" + DateTime.Now.ToString("yyyyMMddhhmmss") + ".xlsx";
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
                Aspose.Cells.License li = new Aspose.Cells.License();
                Aspose.Cells.Workbook wk = new Aspose.Cells.Workbook();
                Worksheet ws = wk.Worksheets[0];
                for (int i = 0; i < listview.Items.Count; i++)
                {

                    ListViewItem item = listview.Items[i];
                    for (int j = 0; j < item.SubItems.Count; j++)
                    {
                        str = item.SubItems[j].Text;
                        ws.Cells[i, j].PutValue(str);
                    }
                }
                wk.Save(path+sname);
            }
            catch(Exception ex)
            {
                MessageBox.Show("Excel写入失败,请检查是否安装了office---"+ex.Message.ToString());
            }
            #endregion
        }

        /// <summary>
        /// 获取BUG按日期查询
        /// </summary>
        private int GetWorkItemCollectionfordate(string proname, DateTime date)
        {
            #region
            DateTime enddate = date.AddDays(1);
            string sql = @"Select [System.Title] From WorkItems Where [System.WorkItemType] ='Bug' and [System.Title] contains '{0}' and [System.CreatedDate] >= '{1}'and [System.CreatedDate] < '{2}'";
            sql = string.Format(sql, proname, date, enddate);
            try
            {
                queryResults = workstore.Query(sql);
            }
            catch(Exception ex)
            {
                MessageBox.Show("Query exec Fail-- "+ex.ToString());
            }
            return queryResults.Count;
            #endregion
        }


        /// <summary>
        /// sql按BUG日期查询
        /// </summary>
        private void BugQueryfordate()
        {
            #region
            int totalbug = 0;

            DateTime startdate = DateTime.Parse(dateTimePicker1.Text);

            DateTime enddate = DateTime.Parse(dateTimePicker2.Text);

            int days = (enddate-startdate).Days;

            productname = textBox6.Text;

            Seconds = 0;

            if (!string.IsNullOrEmpty(productname))
            {
                button4.Enabled = false;
                //启动定时器显示秒表
                tmSeconds = new System.Timers.Timer();
                tmSeconds.Elapsed += TimerTick_Elapsed;
                tmSeconds.Interval = 1000;
                tmSeconds.Enabled = true;

                listView2.Items.Clear();

                TFSServerDto("http://tfs2018-web.winning.com.cn:8080/tfs/WN_HIS");

                

                for (int i = 0; i < days+1; i++ )
                {
                    int currentcount = GetWorkItemCollectionfordate(productname, startdate.AddDays(i));

                    ListViewItem item = new ListViewItem();
                    item.Text = startdate.AddDays(i).ToShortDateString();
                    item.SubItems.Add((currentcount + totalbug).ToString());
                    if (totalbug == 0)
                    {
                        item.SubItems.Add("-");
                    }
                    else
                    {
                        item.SubItems.Add(((float)currentcount / totalbug).ToString("p0"));
                    }
                    listView2.Items.Add(item);

                    totalbug += currentcount;
                }

                ListViewItem item1 = new ListViewItem();
                item1.Text = "总计";
                item1.SubItems.Add(totalbug.ToString());
                item1.SubItems.Add(" ");
                listView2.Items.Add(item1);

                WriteExcel(2, listView2);
                button4.Enabled = true;
                this.tmSeconds.Stop();
            }
            else
            {
                MessageBox.Show("请输入产品名称！");
            }
            #endregion
        }


        /// <summary>
        /// BUG增长趋势查询线程
        /// </summary>
        private void TrendClick(object sender, EventArgs e)
        {
            #region
            new Thread(new ThreadStart(this.BugQueryfordate)).Start();
            #endregion
        }

    }
}
