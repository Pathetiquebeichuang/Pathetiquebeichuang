using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using MediumSoft.Serialization;
using System.Windows.Forms;
using System.Threading;

namespace Agency.DataExchange
{
    public partial class Exchange : ServiceBase
    {
        public Exchange()
        {
            InitializeComponent();
        }
        private System.Text.RegularExpressions.Regex reg = new System.Text.RegularExpressions.Regex("\\{([A-Za-z_0-9]+?)\\}");
        protected override void OnStart(string[] args)
        {
            Common.Quit = false;
            System.Threading.Thread th = new System.Threading.Thread(new System.Threading.ThreadStart(Publish));
            th.Start();
            //Publish();
        }

        protected override void OnStop()
        {
            Common.Quit = true;
            //ProcessStartInfo p = new ProcessStartInfo(Application.StartupPath + "\\Message.exe", "服务停止了");
            //Process.Start(p);
        }

        protected override void OnShutdown()
        {
            Common.Quit = true;
            base.OnShutdown();
        }

        public void Publish()
        {
            var start = DateTime.Now;
            DateTime end;
            int count = 0, success = 0;
            string error = "";
            System.Text.Encoding encode = System.Text.Encoding.UTF8;
            try
            {
                WebService.DataService service = new WebService.DataService();
                var re = service.GetPublishHouse();
                if (re.IsSuccess)
                {
                    System.Net.WebClient c = new System.Net.WebClient();
                    c.Encoding = encode;
                    count = re.Data.Length;
                    re.Data.ToList().ForEach(fangyid =>
                           {
                               var pub_str = c.DownloadString(string.Format(System.Configuration.ConfigurationManager.AppSettings["PublishUrl"], fangyid));
                               var pub_re = Deserialize<ClientResult>(pub_str);
                               if (pub_re.result == ActionResultClient.Success)
                               {
                                   success += 1;
                               }
                               else
                               {
                                   error += string.Format("房源id=>{0}发布失败，发生错误：{1}\n", fangyid, pub_re.message);
                               }
                               service.SetPublishState(fangyid);
                           });
                }
            }
            catch (Exception ex)
            {
                error = ex.ToString();
            }
            end = DateTime.Now;
            DateTime next = start.AddMinutes(Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["Delay"]));
            var ticks = next - end;
            if (ticks.TotalMilliseconds < 0)
            {
                ticks = new TimeSpan(1000000);
                next = end.AddSeconds(10);
            }
            string path = string.Format(System.Configuration.ConfigurationManager.AppSettings["log_path"], end);
            if (!System.IO.Directory.Exists(path))
                System.IO.Directory.CreateDirectory(path);
            string file = System.IO.Path.Combine(path, "log.txt");
            System.IO.File.AppendAllText(file, string.Format("---------------------------------\r\n开始时间:{0:yyyy-MM-dd HH:mm:ss}\r\n结束时间:{1:yyyy-MM-dd HH:mm:ss}\r\n房源数:{2}\r\n成功:{3}\r\n失败:{4}\r\n下次执行：{5:yyyy-MM-dd HH:mm:ss}\r\n间隔：{6:0.00}分钟\r\n---------------------------------\r\n", start, end, count, success, error, next, ticks.Minutes), encode);
            if (Sleep(ticks, 2))
            {
                System.Threading.Thread th = new System.Threading.Thread(new System.Threading.ThreadStart(Publish));
                th.Start();
            }

        }

        public void Gapz()
        {
            DateTime start = DateTime.Now;
            StringBuilder log_build = new StringBuilder();
            Action<string> addLog = (c) =>
            {
                log_build.Append(string.Format("{0:HH:mm:ss}:\r\n{1}\r\n", DateTime.Now));

            };
            try
            {
                addLog("启动获取待人房核验数据");
                var data = Common.GA_Query_List();
                if (data.code == 200)
                {
                    addLog(string.Format("公安接口调用成功，返回{0}条待人房核验数据,稍候启动人房核验", data.data.Count()));
                    int i = 1;
                    data.data.ToList().ForEach(d =>
                    {
                        addLog(string.Format("第{0}条，开始调用房源核验,内容:{1}", i++, Common.Serialize(d)));
                        ga_service.ForGA service = new ga_service.ForGA();
                        var check_result = service.CheckHouse(d.XM, d.SFZH, d.QZH, d.ZJLX);
                        
                    });
                }
            }
            catch (Exception ex)
            {
                addLog(string.Format("发生错误:{0}", ex.ToString()));
            }

            DateTime end = DateTime.Now;
            DateTime next = start.AddMinutes(Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["ga_Delay"]));
            var ticks = next - end;
            if (ticks.TotalMilliseconds < 0)
            {
                ticks = new TimeSpan(1000000);
                next = end.AddSeconds(10);
            }
            string path = string.Format(System.Configuration.ConfigurationManager.AppSettings["ga_log_path"], end);
            if (!System.IO.Directory.Exists(path))
                System.IO.Directory.CreateDirectory(path);
            string file = System.IO.Path.Combine(path, string.Format("log_{0:yyyyMMddHHmmss}.txt", DateTime.Now));
            addLog(string.Format("结束执行，下次执行时间预计{0}分钟后", ticks.Minutes));
            if (Sleep(ticks, 2))
            {
                System.Threading.Thread th = new System.Threading.Thread(new System.Threading.ThreadStart(Publish));
                th.Start();
            }
        }
        /// <summary>
        /// step为秒
        /// </summary>
        /// <param name="TotalSpan"></param>
        /// <param name="Step"></param>
        /// <returns></returns>
        private bool Sleep(TimeSpan TotalSpan, int Step)
        {
            TimeSpan start = new TimeSpan(0);
            TimeSpan step = new TimeSpan(Step * 1000 * 10000);
            while (start <= TotalSpan && Common.Quit == false)
            {
                Thread.Sleep(step);
                start += step;
            }
            if (start > TotalSpan)
                return true;
            else
                return false;

        }
        public class PublishData
        {
            public int id { get; set; }
            public int fangyid { get; set; }
            public DateTime Valid_Date { get; set; }
        }
        public class ClientResult
        {
            public ActionResultClient result { get; set; }
            public string html { get; set; }
            public object data { get; set; }
            public string message { get; set; }
            public string script { get; set; }
            public bool ReLoadData { get; set; }
            public object OpenNew { get; set; }
            public dynamic ModelState { get; set; }
        }
        public enum ActionResultClient
        {
            Success,
            Error,
            ReLogin,
            Warning
        }

        private T Deserialize<T>(string json)
        {
            JavaScriptSerializer Serializer = new JavaScriptSerializer();
            return Serializer.Deserialize<T>(json);
        }
    }
}
