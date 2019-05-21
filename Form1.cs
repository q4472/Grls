using ExcelDataReader;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using System.Xml;

namespace Grls
{
    public partial class Form1 : Form
    {
        delegate void LogPopLineCallback1(String msg);
        delegate void LogPopLineCallback2(String msg);
        delegate void LogPopLineCallback3(String msg);
        private Grls grls;
        private Grlp grlp;
        public Form1()
        {
            InitializeComponent();

            grls = new Grls(this);
            grlp = new Grlp(this);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            grls.StopFlag = false;
            Thread thread = new Thread(grls.Download);
            thread.Start();
        }
        private void button4_Click(object sender, EventArgs e)
        {
            grlp.StopFlag = false;
            Thread thread = new Thread(grlp.Download);
            thread.Start();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            grls.StopFlag = true;
        }
        private void button3_Click(object sender, EventArgs e)
        {
            grlp.StopFlag = true;
        }
        public void LogPopLine1(String msg)
        {
            if (this.richTextBox1.InvokeRequired)
            {
                LogPopLineCallback1 d = new LogPopLineCallback1(LogPopLine1);
                this.Invoke(d, new object[] { msg });
            }
            else
            {
                richTextBox1.Text = msg + "\n" + richTextBox1.Text;
                richTextBox1.Refresh();
            }
        }
        public void LogPopLine2(String msg)
        {
            if (this.richTextBox2.InvokeRequired)
            {
                LogPopLineCallback2 d = new LogPopLineCallback2(LogPopLine2);
                this.Invoke(d, new object[] { msg });
            }
            else
            {
                richTextBox2.Text = msg + "\n" + richTextBox2.Text;
                richTextBox2.Refresh();
            }
        }
        public void LogPopLine3(String msg)
        {
            if (this.richTextBox3.InvokeRequired)
            {
                LogPopLineCallback3 d = new LogPopLineCallback3(LogPopLine3);
                this.Invoke(d, new object[] { msg });
            }
            else
            {
                richTextBox3.Text = msg + "\n" + richTextBox3.Text;
                richTextBox3.Refresh();
            }
        }

        private void buttonExcel_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                LogPopLine3(openFileDialog1.FileName);
                DataSet grlsExcel = null;
                FileStream stream = File.Open(openFileDialog1.FileName, FileMode.Open, FileAccess.Read);
                try
                {
                    IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                    grlsExcel = excelReader.AsDataSet();
                    excelReader.Close();
                }
                catch (Exception exc) { LogPopLine3(exc.ToString()); }
                if (grlsExcel != null && grlsExcel.Tables.Count > 0)
                {
                    DataTable table = grlsExcel.Tables[0];
                    LogPopLine3(table.Rows.Count.ToString());

                    Int32.TryParse(textBox4.Text, out Int32 ColIndex);
                    Int32.TryParse(textBox2.Text, out Int32 StartIndex);
                    Int32.TryParse(textBox3.Text, out Int32 StopIndex);

                    List<String> nums = new List<String>();
                    foreach (DataRow row in table.Rows)
                    {
                        String num = row[ColIndex] as String;
                        if (!String.IsNullOrWhiteSpace(num) && num.Length > 3 && num.Substring(0, 3) != "ФС-" && !nums.Contains(num))
                        {
                            nums.Add(num);
                        }
                    }

                    grls.RegNums = nums;
                    grls.StartIndex = StartIndex;
                    grls.StopIndex = (StopIndex == 0) ? nums.Count : StopIndex;

                    grlp.RegNums = nums;
                    grlp.StartIndex = StartIndex;
                    grlp.StopIndex = (StopIndex == 0) ? nums.Count : StopIndex;

                    LogPopLine3($"{nums.Count}, {ColIndex}, {StartIndex}, {StopIndex}");
                }
                else { LogPopLine3("Excel read as DataSet error."); }
            }
        }
    }

    public class Grls
    {
        private static String cnString =
                "Data Source=192.168.135.14;" +
                "Initial Catalog=Grls;" +
                "Integrated Security=True";
        private static SqlConnection cn;
        private static SqlCommand cmdWriteErr;
        private static SqlCommand cmdGetList;
        private static SqlCommand cmdDelRg;
        private static SqlCommand cmdAddRg { get; set; }
        private static SqlCommand cmdAddF;
        private static SqlCommand cmdAddU;
        private static SqlCommand cmdAddM;

        private static DateTime StartRequestTime;
        private static String CookieSet;

        public List<String> RegNums;
        public Int32 StartIndex;
        public Int32 StopIndex;
        public Boolean StopFlag;
        private Form1 form1;
        public Grls(Form1 _f1)
        {
            cn = new SqlConnection(cnString);
            StopFlag = false;
            cmdWriteErr = WriteErr();
            cmdGetList = GetList();
            cmdDelRg = DelRg();
            cmdAddRg = AddRg();
            cmdAddF = AddF();
            cmdAddU = AddU();
            cmdAddM = AddM();

            StartRequestTime = DateTime.Now;
            CookieSet = null;

            form1 = _f1;
        }
        public void Download()
        {
            if (RegNums == null)
            {
                DataTable regNumList = new DataTable();
                SqlDataAdapter da = new SqlDataAdapter(cmdGetList);
                da.Fill(regNumList);

                if (regNumList.Rows.Count == 0) { return; }

                // отчёт
                form1.LogPopLine1(regNumList.Rows.Count.ToString());

                RegNums = new List<String>();
                for (int ri = 0; ri < regNumList.Rows.Count; ri++)
                {
                    String regNum = regNumList.Rows[ri][0] as String;
                    RegNums.Add(regNum);
                }
                StartIndex = 0;
                StopIndex = RegNums.Count;
            }

            try
            {
                CreateSession();
            }
            catch (Exception ex) { form1.LogPopLine1(ex.ToString()); return; }

            for (int ri = StartIndex; ri < Math.Min(StopIndex, RegNums.Count); ri++)
            {
                String regNum = RegNums[ri];

                DownloadRg(regNum);

                // отчёт
                form1.LogPopLine1(ri.ToString() + ") " + regNum);
                if (StopFlag) { break; }
            }

            // отчёт
            form1.LogPopLine1("--------------");
        }
        private static void CreateSession()
        {
            Uri uri = new Uri("http://grls.rosminzdrav.ru/GRLS.aspx");
            String receivedString = GetResponse(uri);
        }
        private static SqlCommand WriteErr()
        {
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = cn;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "[dbo].[Ошибки при загрузке Добавить]";
            cmd.Parameters.Add("@reg_num", SqlDbType.NVarChar, 50);
            cmd.Parameters.Add("@level", SqlDbType.Int);
            return cmd;
        }
        private static SqlCommand GetList()
        {
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = cn;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "[dbo].[Список несоответствий Grls]";
            return cmd;
        }
        private static SqlCommand DelRg()
        {
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = cn;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "[dbo].[Регистрационные записи Удалить]";
            cmd.Parameters.Add("@Код_в_источнике", SqlDbType.NVarChar, 36);
            return cmd;
        }
        private static SqlCommand AddRg()
        {
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = cn;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "[dbo].[Регистрационные записи Добавить]";
            cmd.Parameters.Add("@Код_в_источнике", SqlDbType.NVarChar, 36);
            cmd.Parameters.Add("@Номер", SqlDbType.NVarChar, 4000);
            cmd.Parameters.Add("@Дата_регистрации", SqlDbType.NVarChar, 4000);
            cmd.Parameters.Add("@Дата_переоформления", SqlDbType.NVarChar, 4000);
            cmd.Parameters.Add("@Дата_окончания_действия", SqlDbType.NVarChar, 4000);
            cmd.Parameters.Add("@Срок", SqlDbType.NVarChar, 4000);
            cmd.Parameters.Add("@Дата_аннулирования", SqlDbType.NVarChar, 4000);
            cmd.Parameters.Add("@Владелец_РУ", SqlDbType.NVarChar, 4000);
            cmd.Parameters.Add("@Страна", SqlDbType.NVarChar, 4000);
            cmd.Parameters.Add("@Торговое_наименование", SqlDbType.NVarChar, 4000);
            cmd.Parameters.Add("@Мнн", SqlDbType.NVarChar, 4000);
            cmd.Parameters.Add("@ЖНВЛП", SqlDbType.NVarChar, 4000);
            cmd.Parameters.Add("@Нарко", SqlDbType.NVarChar, 4000);
            cmd.Parameters.Add("@hfIdReg", SqlDbType.NVarChar, 4000);
            cmd.Parameters.Add("@Код_АТХ", SqlDbType.NVarChar, 4000);
            return cmd;
        }
        private static SqlCommand AddF()
        {
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = cn;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "[dbo].[Формы выпуска Добавить]";
            cmd.Parameters.Add("@Код_РЗ", SqlDbType.Int);
            cmd.Parameters.Add("@Лекарственная_форма", SqlDbType.NVarChar, 4000);
            cmd.Parameters.Add("@Дозировка", SqlDbType.NVarChar, 4000);
            cmd.Parameters.Add("@Срок_годности", SqlDbType.NVarChar, 4000);
            cmd.Parameters.Add("@Условия_хранения", SqlDbType.NVarChar, 4000);
            return cmd;
        }
        private static SqlCommand AddU()
        {
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = cn;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "[dbo].[Упаковки Добавить]";
            cmd.Parameters.Add("@Код_ФВ", SqlDbType.Int);
            cmd.Parameters.Add("@Упаковка", SqlDbType.NVarChar, 4000);
            return cmd;
        }
        private static SqlCommand AddM()
        {
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = cn;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "[dbo].[Стадии производства Добавить]";
            cmd.Parameters.Add("@Код_РЗ", SqlDbType.Int);
            cmd.Parameters.Add("@Стадия_производства", SqlDbType.NVarChar, 4000);
            cmd.Parameters.Add("@Производитель", SqlDbType.NVarChar, 4000);
            cmd.Parameters.Add("@Адрес", SqlDbType.NVarChar, 4000);
            cmd.Parameters.Add("@Страна", SqlDbType.NVarChar, 4000);
            return cmd;
        }
        private void DownloadRg(String regNum)
        {
            String[] ids = null;
            try
            {
                ids = GetIds(regNum);
            }
            catch (Exception e)
            {
                // отчёт
                form1.LogPopLine1(e.Message);
                cmdWriteErr.Parameters["@reg_num"].Value = regNum;
                cmdWriteErr.Parameters["@level"].Value = 1;
                cn.Open();
                cmdWriteErr.ExecuteNonQuery();
                cn.Close();
            }
            if (ids == null) { return; }

            foreach (String id in ids)
            {
                String receivedString = null;
                try
                {
                    receivedString = GetId(id);
                }
                catch (Exception e)
                {
                    // отчёт
                    form1.LogPopLine1(e.Message);
                    cmdWriteErr.Parameters["@reg_num"].Value = regNum;
                    cmdWriteErr.Parameters["@level"].Value = 2;
                    cn.Open();
                    cmdWriteErr.ExecuteNonQuery();
                    cn.Close();
                }
                XmlDocument doc = ParseXml(receivedString);
                if (doc == null) { continue; }

                cn.Open();
                UpsertRg(id, doc);
                cn.Close();

                // отчёт
                form1.LogPopLine1("- " + id.ToString());
            }
        }
        private void UpsertRg(String sid, XmlDocument doc)
        {
            cmdDelRg.Parameters["@Код_в_источнике"].Value = sid;
            cmdDelRg.ExecuteNonQuery();

            cmdAddRg.Parameters["@Код_в_источнике"].Value = sid;
            cmdAddRg.Parameters["@Номер"].Value = GetValue(doc.SelectSingleNode("/.//*[@id='ctl00_plate_RegNr']"));
            cmdAddRg.Parameters["@Дата_регистрации"].Value = GetValue(doc.SelectSingleNode("/.//*[@id='ctl00_plate_RegDate']"));
            cmdAddRg.Parameters["@Дата_переоформления"].Value = GetValue(doc.SelectSingleNode("/.//*[@id='ctl00_plate_DChange']"));
            cmdAddRg.Parameters["@Дата_окончания_действия"].Value = GetValue(doc.SelectSingleNode("/.//*[@id='ctl00_plate_RegDateFin']"));
            cmdAddRg.Parameters["@Срок"].Value = GetValue(doc.SelectSingleNode("/.//*[@id='ctl00_plate_txtCirculationPeriod']"));
            cmdAddRg.Parameters["@Дата_аннулирования"].Value = GetValue(doc.SelectSingleNode("/.//*[@id='ctl00_plate_Annul']"));
            cmdAddRg.Parameters["@Владелец_РУ"].Value = GetInnerText(doc.SelectSingleNode("/.//*[@id='ctl00_plate_MnfClNmR']"));
            cmdAddRg.Parameters["@Страна"].Value = GetInnerText(doc.SelectSingleNode("/.//*[@id='ctl00_plate_CountryClR']"));
            cmdAddRg.Parameters["@Торговое_наименование"].Value = GetValue(doc.SelectSingleNode("/.//*[@id='ctl00_plate_TradeNmR']"));
            cmdAddRg.Parameters["@Мнн"].Value = GetInnerText(doc.SelectSingleNode("/.//*[@id='ctl00_plate_Innr']"));
            cmdAddRg.Parameters["@ЖНВЛП"].Value = GetValue(doc.SelectSingleNode("/.//*[@id='ctl00_plate_txtNecessary']"));
            cmdAddRg.Parameters["@Нарко"].Value = GetValue(doc.SelectSingleNode("/.//*[@id='ctl00_plate_txtNarco']"));
            cmdAddRg.Parameters["@hfIdReg"].Value = GetValue(doc.SelectSingleNode("/.//*[@id='ctl00_plate_hfIdReg']"));
            cmdAddRg.Parameters["@Код_АТХ"].Value = GetInnerText(doc.SelectSingleNode("/.//*[@id='ctl00_plate_grATC']//tr[2]//td[1]"));
            Int32 id = (Int32)cmdAddRg.ExecuteScalar();

            XmlNode drugforms = doc.SelectSingleNode("/.//*[@id='ctl00_plate_drugforms']");
            Int32 fi = 1;
            XmlNode f = drugforms.SelectSingleNode(".//table//tr[" + ((fi * 2) + 1) + "]");
            while (f != null)
            {
                cmdAddF.Parameters["@Код_РЗ"].Value = id;
                cmdAddF.Parameters["@Лекарственная_форма"].Value = GetInnerText(f.ChildNodes[0]);
                cmdAddF.Parameters["@Дозировка"].Value = GetInnerText(f.ChildNodes[1]);
                cmdAddF.Parameters["@Срок_годности"].Value = GetInnerText(f.ChildNodes[2]);
                cmdAddF.Parameters["@Условия_хранения"].Value = GetInnerText(f.ChildNodes[3]);

                Int32 fId = (Int32)cmdAddF.ExecuteScalar();

                XmlNodeList lis = drugforms.SelectNodes(".//table//tr[" + ((fi * 2) + 2) + "]/td/ul/li");
                if (lis != null)
                {
                    foreach (XmlNode li in lis)
                    {
                        cmdAddU.Parameters["@Код_ФВ"].Value = fId;
                        cmdAddU.Parameters["@Упаковка"].Value = GetInnerText(li);

                        cmdAddU.ExecuteNonQuery();
                    }
                }

                // отчёт
                form1.LogPopLine1("-f- " + fId.ToString());

                fi++;
                f = drugforms.SelectSingleNode(".//table//tr[" + ((fi * 2) + 1) + "]");
            }

            XmlNode grMnf = doc.SelectSingleNode("/.//*[@id='ctl00_plate_gr_mnf']");
            if (grMnf != null)
            {
                Int32 si = 1;
                XmlNode s = grMnf.SelectSingleNode(".//tr[" + (si + 1) + "]");
                while (s != null)
                {
                    cmdAddM.Parameters["@Код_РЗ"].Value = id;
                    cmdAddM.Parameters["@Стадия_производства"].Value = GetInnerText(s.ChildNodes[1]);
                    cmdAddM.Parameters["@Производитель"].Value = GetInnerText(s.ChildNodes[2]);
                    cmdAddM.Parameters["@Адрес"].Value = GetInnerText(s.ChildNodes[3]);
                    cmdAddM.Parameters["@Страна"].Value = GetInnerText(s.ChildNodes[4]);

                    Int32 sId = (Int32)cmdAddM.ExecuteScalar();

                    // отчёт
                    form1.LogPopLine1("-s- " + sId.ToString());

                    si++;
                    s = grMnf.SelectSingleNode(".//tr[" + (si + 1) + "]");
                }
            }
        }
        private String GetValue(XmlNode n)
        {
            String v = "";
            if (n != null)
            {
                XmlAttribute a = n.Attributes["value"];
                if (a != null)
                {
                    v = a.InnerText;
                }
            }
            return Trim(v);
        }
        private String GetInnerText(XmlNode n)
        {
            String v = (n == null) ? "" : n.InnerText;
            return Trim(v);
        }
        private String Trim(String s)
        {
            return ((new Regex(@"^[\s\x0A]*|[\s\x0A]*$")).Replace(s, ""));
        }
        private String[] GetIds(String regNum)
        {
            List<String> ids = new List<String>();

            Uri uri = new Uri("http://grls.rosminzdrav.ru/GRLS.aspx" +
                "?RegNumber=" + regNum.Replace(" ", "%20").Replace("/", "%2F") +
                "&MnnR=" +
                "&lf=" +
                "&TradeNmR=" +
                "&OwnerName=" +
                "&MnfOrg=" +
                "&MnfOrgCountry=" +
                "&isfs=0" +
                "&isND=-1" +
                "&regtype=1" +
                "&pageSize=50" +
                "&order=RegDate" +
                "&orderType=desc" +
                "&pageNum=1");

            String receivedString = GetResponse(uri);

            if (!String.IsNullOrWhiteSpace(receivedString))
            {
                // в ответе может появиться captcha - тогда надо сбросить сессию и запросить по новой
                if (receivedString.IndexOf("ctl00_plate_dCaptcha") > 0)
                {
                    form1.LogPopLine1("ctl00_plate_dCaptcha");
                    CookieSet = null;
                    CreateSession();
                    receivedString = GetResponse(uri);
                }

                Int32 s = receivedString.IndexOf("\"det(");
                while (s >= 0)
                {
                    Int32 e = receivedString.IndexOf(',', s);
                    if (e >= 0)
                    {
                        String str = receivedString.Substring(s + 5, e - s - 5);
                        if (str[0] == '\'')
                        {
                            ids.Add(str.Substring(1, str.Length - 2));
                        }
                        else if (str.Substring(0, 5) == "&#39;")
                        {
                            ids.Add(str.Substring(5, str.Length - 10));
                        }
                        else
                        {
                            ids.Add(str);
                        }
                    }
                    s = receivedString.IndexOf("\"det(", e);
                }
            }
            return ids.ToArray();
        }
        private String GetId(String id)
        {
            Uri uri;
            if (id.Length < 36)
            {
                uri = new Uri("http://grls.rosminzdrav.ru/Grls_View_v2.aspx?idReg=" + id);
            }
            else
            {
                uri = new Uri("http://grls.rosminzdrav.ru/Grls_View_v2.aspx?routingGuid=" + id + "&t=");
            }
            String receivedString = GetResponse(uri);
            return receivedString;
        }
        private static String GetResponse(Uri uri)
        {
            DateTime now = DateTime.Now;
            TimeSpan timeout = (new TimeSpan(0, 0, 1)).Subtract(now - StartRequestTime);
            if (timeout.Milliseconds > 0) { Thread.Sleep(timeout); }
            StartRequestTime = DateTime.Now;

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(uri);
            request.Method = "GET";
            //request.ContentLength = 0;
            request.Credentials = CredentialCache.DefaultCredentials;
            request.Timeout = 10000; // 10 sec.

            HttpWebResponse response = null;
            String receivedString = null;
            //try
            {
                if (!String.IsNullOrWhiteSpace(CookieSet))
                {
                    String cName = "ASP.NET_SessionId";
                    String cValue = "";
                    String[] cs = CookieSet.Split(';');
                    if (cs.Length > 0)
                    {
                        foreach (String c in cs)
                        {
                            String[] nv = c.Split('=');
                            if (nv.Length == 2 && nv[0] == cName)
                            {
                                cValue = nv[1];
                                if (!String.IsNullOrWhiteSpace(cValue))
                                {
                                    request.CookieContainer = new CookieContainer();
                                    request.CookieContainer.Add(new Uri("http://grls.rosminzdrav.ru"), new Cookie(cName, cValue));
                                    break;
                                }
                            }
                        }
                    }
                }
                response = (HttpWebResponse)request.GetResponse();
                if (String.IsNullOrWhiteSpace(CookieSet))
                {
                    CookieSet = response.Headers["Set-Cookie"];
                }
            }
            //catch (Exception e)
            {
                //throw new Exception(e.Message);
            }
            if ((response != null) && (response.StatusCode == HttpStatusCode.OK))
            {
                Stream receivedStream = response.GetResponseStream();
                StreamReader readStream = new StreamReader(receivedStream, Encoding.UTF8);

                receivedString = readStream.ReadToEnd();

                readStream.Close();
                response.Close();
            }
            return receivedString;
        }
        private XmlDocument ParseXml(String receivedString)
        {
            XmlDocument doc = null;
            if (String.IsNullOrWhiteSpace(receivedString)) { return null; }

            Int32 startIndex = receivedString.IndexOf("<body");

            if (startIndex < 0) { return null; }

            Int32 endIndex = receivedString.IndexOf("</body>");
            String tempDoc = receivedString.Substring(startIndex, endIndex - startIndex + 7);

            int s = tempDoc.IndexOf("<script");
            int e = tempDoc.IndexOf("</script>");
            while (s >= 0)
            {
                tempDoc = tempDoc.Substring(0, s) + tempDoc.Substring(e + 9);
                s = tempDoc.IndexOf("<script");
                e = tempDoc.IndexOf("</script>");
            }


            tempDoc = tempDoc
                .Replace("&nbsp;", " ")
                .Replace("<br />", " ")
                .Replace("<br>", " ")
                .Replace("</br>", "")
                .Replace("=ts1", "='ts1'")
                .Replace("=all", "='all'")
                .Replace("=hdr_flat", "='hdr_flat'")
                .Replace("=2", "='2'")
                .Replace("=hi_sys", "='hi_sys'")
                .Replace("&", "&amp;")
                .Replace("//<![CDATA[", "")
                .Replace("//]]>", "")
                .Replace("<!-- Повторите попытку.- ->", "")
                ;

            ArrayList tempLines = new ArrayList();

            /*
            String[] splitLines = tempDoc.Split(new String[] { "\r\n" }, StringSplitOptions.None);
            for (int i = 0; i < splitLines.Length; i++)
            {
                tempLines.Add(i.ToString() + "> " + splitLines[i] + "\r\n");
            }

            richTextBox1.Lines = (String[])tempLines.ToArray(typeof(String));
            */

            doc = new XmlDocument();
            try
            {
                doc.LoadXml(tempDoc);
            }
            catch (Exception ex) { form1.LogPopLine1(ex.Message); return null; }

            return doc;
        }
    }
    public class Grlp
    {
        private static String cnString =
                "Data Source=192.168.135.14;" +
                "Initial Catalog=Grls;" +
                "Integrated Security=True";
        private static SqlConnection cn;
        private static DateTime StartRequestTime;
        private static SqlCommand cmdWriteErr;
        private static SqlCommand cmdGetList;
        private static SqlCommand cmdDelRg;
        private static SqlCommand cmdAddRg;

        public List<String> RegNums;
        public Int32 StartIndex;
        public Int32 StopIndex;
        public Boolean StopFlag;
        private Form1 form1;
        public Grlp(Form1 _f1)
        {
            cn = new SqlConnection(cnString);
            StopFlag = false;
            StartRequestTime = DateTime.Now;
            cmdWriteErr = WriteErr();
            cmdGetList = GetList();
            cmdDelRg = DelRg();
            cmdAddRg = AddRg();
            form1 = _f1;
        }
        public void Download()
        {
            if (RegNums == null)
            {
                DataTable regNumList = new DataTable();
                SqlDataAdapter da = new SqlDataAdapter(cmdGetList);
                da.Fill(regNumList);

                if (regNumList.Rows.Count == 0) { return; }

                // отчёт
                form1.LogPopLine2(regNumList.Rows.Count.ToString());

                RegNums = new List<String>();
                for (int ri = 0; ri < regNumList.Rows.Count; ri++)
                {
                    String regNum = regNumList.Rows[ri][0] as String;
                    RegNums.Add(regNum);
                }
                StartIndex = 0;
                StopIndex = RegNums.Count;
            }

            try
            {
                //CreateSession();
            }
            catch (Exception ex) { form1.LogPopLine1(ex.ToString()); return; }


            for (int ri = StartIndex; ri < Math.Min(StopIndex, RegNums.Count); ri++)
            {
                String regNum = RegNums[ri];

                DownloadRg(regNum);

                // отчёт
                form1.LogPopLine2(ri.ToString() + ") " + regNum);
                if (StopFlag) { break; }
            }

            // отчёт
            form1.LogPopLine2("--------------");
        }
        private static void CreateSession()
        {
            Uri uri = new Uri("http://grls.rosminzdrav.ru/PriceLims.aspx");
            String receivedString = GetResponse(uri);
        }
        private static SqlCommand WriteErr()
        {
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = cn;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "[dbo].[Ошибки при загрузке Добавить]";
            cmd.Parameters.Add("@reg_num", SqlDbType.NVarChar, 50);
            cmd.Parameters.Add("@level", SqlDbType.Int);
            return cmd;
        }
        private static SqlCommand GetList()
        {
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = cn;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "[dbo].[Список несоответствий Grlp]";
            return cmd;
        }
        private static SqlCommand DelRg()
        {
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = cn;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "[dbo].[Регистрационные записи Удалить Grlp]";
            cmd.Parameters.Add("@Номер", SqlDbType.NVarChar, 4000);
            return cmd;
        }
        private static SqlCommand AddRg()
        {
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = cn;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "[dbo].[Регистрационные записи Добавить Grlp]";
            cmd.Parameters.Add("@Мнн", SqlDbType.NVarChar, 4000);
            cmd.Parameters.Add("@Торговое_наименование", SqlDbType.NVarChar, 4000);
            cmd.Parameters.Add("@Лек_форма_дозировка_упаковка", SqlDbType.NVarChar, 4000);
            cmd.Parameters.Add("@Владелец_РУ", SqlDbType.NVarChar, 4000);
            // Код АТХ
            cmd.Parameters.Add("@Количество_в_потреб_упаковке", SqlDbType.NVarChar, 4000);
            cmd.Parameters.Add("@Предельная_цена_руб_без_НДС", SqlDbType.NVarChar, 4000);
            cmd.Parameters.Add("@Номер", SqlDbType.NVarChar, 4000);
            cmd.Parameters.Add("@Дата_регистрации_цены", SqlDbType.NVarChar, 4000);
            cmd.Parameters.Add("@Номер_решения", SqlDbType.NVarChar, 4000);
            cmd.Parameters.Add("@Штрих_код_EAN13", SqlDbType.NVarChar, 4000);
            return cmd;
        }
        private void DownloadRg(String regNum)
        {
            String plateGrTable = null;
            try
            {
                plateGrTable = GetPlateGrTable(regNum);
            }
            catch (Exception e)
            {
                // отчёт
                form1.LogPopLine2(e.Message);
                cmdWriteErr.Parameters["@reg_num"].Value = regNum;
                cmdWriteErr.Parameters["@level"].Value = 1;
                cn.Open();
                cmdWriteErr.ExecuteNonQuery();
                cn.Close();
            }
            if (plateGrTable == null) { return; }
            UpsertRg(regNum, plateGrTable);
        }
        private String GetPlateGrTable(String regNum)
        {
            Uri uri = new Uri("http://grls.rosminzdrav.ru/PriceLims.aspx?Torg=&Mnn=" +
                "&RegNum=" + regNum.Replace(" ", "%20").Replace("/", "%2F") +
                "&Mnf=&Barcode=&Order=&isActual=0&All=0&PageSize=128&orderby=pklimprice&orderType=desc&pagenum=1");

            String receivedString = GetResponse(uri);

            Match match = new Regex("id=\"ctl00_plate_gr\"[^>]*>").Match(receivedString);
            Int32 bi = match.Index + match.Length;
            Int32 ei = receivedString.IndexOf("</table>", bi);
            String table = String.Format("<table>{0}</table>", receivedString.Substring(bi, ei - bi));
            table = table.Replace("&nbsp", " ");
            return table;
        }
        private static String GetResponse(Uri uri)
        {
            DateTime now = DateTime.Now;
            TimeSpan timeout = (new TimeSpan(0, 0, 1)).Subtract(now - StartRequestTime);
            if (timeout.Milliseconds > 0) { Thread.Sleep(timeout); }
            StartRequestTime = DateTime.Now;

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(uri);
            request.Credentials = CredentialCache.DefaultCredentials;
            request.Timeout = 10000; // 10 sec.

            HttpWebResponse response = null;
            String receivedString = null;
            //try
            {
                response = (HttpWebResponse)request.GetResponse();
            }
            //catch (Exception e)
            {
                //throw new Exception(e.Message);
            }
            if ((response != null) && (response.StatusCode == HttpStatusCode.OK))
            {
                Stream receivedStream = response.GetResponseStream();
                StreamReader readStream = new StreamReader(receivedStream, Encoding.UTF8);

                receivedString = readStream.ReadToEnd();

                readStream.Close();
                response.Close();
            }
            return receivedString;
        }
        private void UpsertRg(String regNum, String plateGrTable)
        {
            XmlDocument doc = new XmlDocument();
            try
            {
                doc.LoadXml(plateGrTable);
            }
            catch (Exception ex) { form1.LogPopLine2(ex.Message); return; }

            XmlNodeList trs = doc.SelectNodes("./table/tr");
            cn.Open();
            cmdDelRg.Parameters["@Номер"].Value = regNum;
            cmdDelRg.ExecuteNonQuery();
            for (int ri = 1; ri < trs.Count; ri++)
            {
                XmlNode tr = trs[ri];
                XmlNodeList tds = tr.SelectNodes("td");
                cmdAddRg.Parameters["@Мнн"].Value = tds[1].InnerText;
                cmdAddRg.Parameters["@Торговое_наименование"].Value = tds[2].InnerText;
                cmdAddRg.Parameters["@Лек_форма_дозировка_упаковка"].Value = tds[3].InnerText;
                cmdAddRg.Parameters["@Владелец_РУ"].Value = tds[4].InnerText;
                // ATX tds[5].InnerText;
                cmdAddRg.Parameters["@Количество_в_потреб_упаковке"].Value = tds[6].InnerText;
                cmdAddRg.Parameters["@Предельная_цена_руб_без_НДС"].Value = tds[7].InnerText;
                // цена за перв уп tds[8].InnerText;
                cmdAddRg.Parameters["@Номер"].Value = tds[9].InnerText;
                cmdAddRg.Parameters["@Дата_регистрации_цены"].Value = tds[10].InnerText.TrimStart().Substring(0, 10);
                cmdAddRg.Parameters["@Номер_решения"].Value = tds[10].InnerText.TrimStart().Substring(10).Trim();
                cmdAddRg.Parameters["@Штрих_код_EAN13"].Value = tds[11].InnerText;
                cmdAddRg.ExecuteNonQuery();
                // отчёт
                form1.LogPopLine2("- " + regNum + " - " + ri.ToString());
            }
            cn.Close();
        }
    }
}
