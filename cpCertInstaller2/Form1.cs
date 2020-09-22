using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Principal;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Diagnostics;

namespace cpCertInstaller2
{
    public partial class Form1 : Form
    {
        public static string appdatapath= $@"{Environment.GetEnvironmentVariable("APPDATA")}\dln\cpci2";
        public static journal log = new journal($@"{appdatapath}\cpci2.log");
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;

            if (!Directory.Exists(appdatapath))
            {
                Directory.CreateDirectory(appdatapath);
            }

            regConf conf = new regConf();

            log.writeLog($"Cpci2 agent started.\r\n");
            log.writeLog($"Checking configuration...\r\n");
            try
            {
                string sid = WindowsIdentity.GetCurrent().User.Value;
                toolStripStatusLabel1.Text = $"firstRun: {conf.getParam("firstRun", "")}";
                toolStripStatusLabel2.Text = $"SID: {sid}";
                timer1.Interval = conf.getParam("timerInterval", 999);
                toolStripStatusLabel3.Text = $"timerInterval: {timer1.Interval}";
                log.writeLog($"\t {toolStripStatusLabel1.Text}\r\n");
                log.writeLog($"\t {toolStripStatusLabel2.Text}\r\n");
                log.writeLog($"\t timerInterval: {timer1.Interval}\r\n");
                log.writeLog($"\t baseKey: {conf.getParam("baseKey", "")}\r\n");
                log.writeLog($"\t cpKeys: {conf.getParam("cpKeys", "")}\r\n");
                log.writeLog($"\t pathCSPTest: {conf.getParam("pathCSPTest", "")}\r\n");
                label1.Text=$@"{Environment.GetEnvironmentVariable("APPDATA")}\dln\cpci2\cpci2.log";
    }
            catch(Exception ex)
            {
                log.writeLog(FormatErrorStr("Error while checking configuration!\r\n Details:\r\n", ex));
            }
            timer1.Start();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            regConf conf = new regConf();
            string sid = WindowsIdentity.GetCurrent().User.Value;

            cpRegFile rf = new cpRegFile("input.reg");

            bool res = rf.importRegFile();
            if (!res) MessageBox.Show("error importing");

            cpContainers tmpcontainers = new cpContainers($@"{conf.getParam("baseKey", "")}\{sid}\{conf.getParam("cpKeys", "")}", $@"{appdatapath}\cpKey.list");
            listBox1.Items.Clear();

            List<string> toinstall = tmpcontainers.compareList();
            if (toinstall.Count > 0)
            {
                foreach(string item in toinstall)
                {
                    if (tmpcontainers.instCert(item))
                    {
                        listBox1.Items.Add(item);
                    }
                }
            }
        }

        public static string FormatErrorStr(string message, Exception ex)
        {
            string res = "";
            res += $"{message}\r\n";
            res += $"ECode: {ex.HResult}\r\n Src: {ex.Source}\r\n Trg: {ex.TargetSite}\r\n Msg: {ex.Message}\r\n Stk: {ex.StackTrace}";
            return res;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            string inputpath = $@"{appdatapath}\input.reg";
            bool res = false;
            regConf conf = new regConf();
            string sid = WindowsIdentity.GetCurrent().User.Value;
            timer1.Stop();

            if (File.Exists(inputpath))
            {
                log.writeLog($"{inputpath} found.\r\n");
                cpRegFile rf = new cpRegFile(inputpath);
                log.writeLog($"Trying to import {inputpath}\r\n");
                try
                {
                    res = rf.importRegFile();
                    if (res)
                    {
                        log.writeLog($"{inputpath} successfully imported.\r\n");
                    }
                    else
                    {
                        log.writeLog($"Unknown error while importing {inputpath}!\r\n");
                    }
                }
                catch(Exception ex)
                {
                    log.writeLog(FormatErrorStr($"Error while importing {inputpath}!\r\n Details:\r\n", ex));
                }
            }

            cpContainers cont = new cpContainers($@"{conf.getParam("baseKey", "")}\{sid}\{conf.getParam("cpKeys", "")}", $@"{appdatapath}\cpKey.list");
            try
            {
                List<string> toinstall = cont.compareList();
                List<string> installed = new List<string>();
                if (toinstall.Count > 0)
                {
                    foreach (string item in toinstall)
                    {
                        try
                        {
                            log.writeLog($"Trying to install certificate from {item} container...\r\n");
                            if (cont.instCert(item))
                            {
                                installed.Add(item);
                                log.writeLog($"OK Certificate from {item} container installed successfully!\r\n");
                            }
                            else
                            {
                                log.writeLog($"ERROR while installing certificate {item}! No certificate in key containers?\r\n");
                                log.writeLog($"Try to run {conf.getParam("pathCSPTest", "")} {item} for details.\r\n");
                            }
                        }
                        catch(Exception ex)
                        {
                            log.writeLog(FormatErrorStr($"ERROR while installing certificate {item}!\r\n Details:\r\n", ex));
                        }
                    }

                    if (installed.Count > 0)
                    {
                        try
                        {
                            File.WriteAllLines($@"{appdatapath}\lasttimeinstalled.txt", installed.ToArray());
                        }
                        catch
                        {
                            //
                        }

                        notifyIcon1.BalloonTipTitle = "ЭЦП Установлены!";
                        notifyIcon1.BalloonTipText = $"Количество установленных ЭЦП: {installed.Count}\r\n Нажмите для просмотра списка.";
                        notifyIcon1.ShowBalloonTip(900 * 1000);
                    }
                }
            }
            catch(Exception ex)
            {
                log.writeLog(FormatErrorStr($"Error while comparing containers list!\r\n Details:\r\n", ex));
            }

            timer1.Start();
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            log.writeLog("Cpci2 agent stopped.\r\n\r\n");
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
            this.Hide();
        }

        private void notifyIcon1_BalloonTipClicked(object sender, EventArgs e)
        {
            using(Process notepad = new Process())
            {
                try
                {
                    notepad.StartInfo.FileName = $@"{Environment.GetEnvironmentVariable("WINDIR")}\notepad.exe";
                    notepad.StartInfo.Arguments = $@"{appdatapath}\lasttimeinstalled.txt";
                    notepad.Start();
                }
                catch
                {
                    //
                }
            }
        }
    }

    public class journal
    {
        public string path { get; }

        public journal(string path)
        {
            this.path = path;
        }

        public int writeLog(string text)
        {
            try
            {
                using (FileStream fs = new FileStream(path, FileMode.OpenOrCreate, FileAccess.Write))
                {
                    long eof = fs.Length;
                    string ttw = $"{DateTime.Now.ToString("yyyy.MM.dd HH:mm:ss")} {text}";
                    byte[] buf = Encoding.UTF8.GetBytes(ttw.ToCharArray());
                    fs.Seek(eof, SeekOrigin.Begin);
                    fs.Write(buf, 0, buf.Length);
                }
            }
            catch
            {
                return 1;
            }

            return 0;
        }
    }

    public class regConf
    {
        public string softwareName { get; }

        public regConf(string softwareName=@"dln\cpCertInstaller2")
        {
            this.softwareName = softwareName;
        }

        public bool isParamExists(string name)
        {
            bool res = false;

            try
            {
                using (RegistryKey HKCU = Registry.CurrentUser)
                {
                    using (RegistryKey settings = HKCU.OpenSubKey($@"SOFTWARE\{softwareName}\"))
                    {
                        if (settings.GetValue(name) != null)
                        {
                            res = true;
                        }
                        else
                        {
                            res = false;
                        }
                    }
                }
            }
            catch
            {
                res = false;
            }
            return res;
        }

        public bool isKeyExists(string name)
        {
            bool res = false;
            using (RegistryKey HKCU = Registry.CurrentUser)
            {
                try
                {
                    RegistryKey settings = HKCU.OpenSubKey($@"SOFTWARE\{softwareName}\");
                    settings.GetValue("firstRun");
                    res = true;
                }
                catch
                {
                    using (RegistryKey settings = HKCU.CreateSubKey($@"SOFTWARE\{softwareName}\", true))
                    {
                        try
                        {
                            defaultAction(settings);
                            res = true;
                        }
                        catch
                        {
                            res = false;
                        }
                        
                    }
                }
            }
            

            return res;
        }

        public string getParam(string name, string nothing="")
        {
            string res = null;

            if (isKeyExists(softwareName))
            {
                if (isParamExists(name))
                {
                    using(RegistryKey HKCU = Registry.CurrentUser)
                    {
                        using(RegistryKey settings = HKCU.CreateSubKey($@"SOFTWARE\{softwareName}\"))
                        {
                            RegistryValueKind type = settings.GetValueKind(name);
                            if(type == RegistryValueKind.String |
                               type == RegistryValueKind.MultiString |
                               type == RegistryValueKind.ExpandString)
                            {
                                res = (string) settings.GetValue(name);
                            }
                        }
                    }
                }
                else
                {
                    throw new Exception($"No such parameter: {name}");
                }
            }
            else
            {
                throw new Exception($"No such software key: {softwareName}");
            }

            return res;
        }

        public int getParam(string name, int nothing = 0)
        {
            int res = 0;

            if (isKeyExists(softwareName))
            {
                if (isParamExists(name))
                {
                    using (RegistryKey HKCU = Registry.CurrentUser)
                    {
                        using (RegistryKey settings = HKCU.CreateSubKey($@"SOFTWARE\{softwareName}\"))
                        {
                            RegistryValueKind type = settings.GetValueKind(name);
                            if (type == RegistryValueKind.DWord)
                            {
                                res = (int) settings.GetValue(name);
                            }
                        }
                    }
                }
                else
                {
                    throw new Exception($"No such parameter: {name}");
                }
            }
            else
            {
                throw new Exception($"No such software key: {softwareName}");
            }

            return res;
        }

        public byte[] getParam(string name, byte nothing = 0)
        {
            byte[] res = { };

            if (isKeyExists(softwareName))
            {
                if (isParamExists(name))
                {
                    using (RegistryKey HKCU = Registry.CurrentUser)
                    {
                        using (RegistryKey settings = HKCU.CreateSubKey($@"SOFTWARE\{softwareName}\"))
                        {
                            RegistryValueKind type = settings.GetValueKind(name);
                            if (type == RegistryValueKind.Binary)
                            {
                                res = (byte[]) settings.GetValue(name);
                            }
                        }
                    }
                }
                else
                {
                    throw new Exception($"No such parameter: {name}");
                }
            }
            else
            {
                throw new Exception($"No such software key: {softwareName}");
            }

            return res;
        }

        public void setParam(string name, string value)
        {
            if (isKeyExists(softwareName))
            {
                using (RegistryKey HKCU = Registry.CurrentUser)
                {
                    using (RegistryKey settings = HKCU.CreateSubKey($@"SOFTWARE\{softwareName}\", true))
                    {
                        if (isParamExists(name))
                        {
                            RegistryValueKind type = settings.GetValueKind(name);
                            if (type == RegistryValueKind.String |
                               type == RegistryValueKind.MultiString |
                               type == RegistryValueKind.ExpandString)
                            {
                                settings.SetValue(name, value, type);
                            }
                            else
                            {
                                throw new Exception("Wrong parameter type!");
                            }
                        }
                        else
                        {
                            settings.SetValue(name, value, RegistryValueKind.String);
                        }
                    }
                }
            }
            else
            {
                throw new Exception($"No such software key: {softwareName}");
            }
        }

        public void setParam(string name, uint value)
        {
            if (isKeyExists(softwareName))
            {
                using (RegistryKey HKCU = Registry.CurrentUser)
                {
                    using (RegistryKey settings = HKCU.CreateSubKey($@"SOFTWARE\{softwareName}\", true))
                    {
                        if (isParamExists(name))
                        {
                            RegistryValueKind type = settings.GetValueKind(name);
                            if (type == RegistryValueKind.DWord)
                            {
                                settings.SetValue(name, value, type);
                            }
                            else
                            {
                                throw new Exception("Wrong parameter type!");
                            }
                        }
                        else
                        {
                            settings.SetValue(name, value, RegistryValueKind.DWord);
                        }
                    }
                }
            }
            else
            {
                throw new Exception($"No such software key: {softwareName}");
            }
        }

        public void setParam(string name, byte[] value)
        {
            if (isKeyExists(softwareName))
            {
                using (RegistryKey HKCU = Registry.CurrentUser)
                {
                    using (RegistryKey settings = HKCU.CreateSubKey($@"SOFTWARE\{softwareName}\", true))
                    {
                        if (isParamExists(name))
                        {
                            RegistryValueKind type = settings.GetValueKind(name);
                            if (type == RegistryValueKind.Binary)
                            {
                                settings.SetValue(name, value, type);
                            }
                            else
                            {
                                throw new Exception("Wrong parameter type!");
                            }
                        }
                        else
                        {
                            settings.SetValue(name, value, RegistryValueKind.Binary);
                        }
                    }
                }
            }
            else
            {
                throw new Exception($"No such software key: {softwareName}");
            }
        }

        void defaultAction(RegistryKey settings)
        {
            string ttw = $"{DateTime.Now.ToString("yyyy.MM.dd HH:mm:ss")}";
            settings.SetValue("firstRun", ttw, RegistryValueKind.String);
            settings.SetValue("baseKey", @"SOFTWARE\WOW6432Node\Crypto Pro\Settings\Users\", RegistryValueKind.String);
            settings.SetValue("cpKeys", @"\Keys\", RegistryValueKind.String);
            settings.SetValue("pathCSPTest", @"C:\Program Files\Crypto Pro\CSP\csptest.exe", RegistryValueKind.String);
            settings.SetValue("timerInterval", 5000, RegistryValueKind.DWord);
        }

    }

    public class cpRegFile
    {
        public string path { get; }

        public cpRegFile(string path)
        {
            this.path = path;
        }

        public bool parceRegFile()
        {
            bool res = true;
            string[] regFileData = { };

            try
            {
                regFileData = File.ReadAllLines(path, Encoding.Unicode);
            }
            catch
            {
                return false;
            }

            if (regFileData[0] != "Windows Registry Editor Version 5.00")
            {
                return false;
            }

            uint cnt = 0;
            uint i = 0;
            foreach (string item in regFileData)
            {
                if(item.StartsWith("[") && item.EndsWith("]"))
                {
                    if(item.Contains(new regConf().getParam("baseKey", "")))
                    {
                        string sid = WindowsIdentity.GetCurrent().User.Value;
                        regFileData[i] = Regex.Replace(item, @"S-(\d+-){6}\d+", sid);
                    }
                    else
                    {
                        return false;
                    }
                    cnt++;
                }
                i++;
            }

            if (cnt < 1) return false;

            File.Delete(path);
            File.WriteAllLines(path, regFileData, Encoding.Unicode);
            return res;
        }

        public bool importRegFile()
        {
            bool res = false;
            bool rres = parceRegFile();
            if (rres)
            {
                FileInfo rf = new FileInfo(path);
                System.Diagnostics.Process psRegexe = new System.Diagnostics.Process();
                psRegexe.StartInfo.FileName = $@"{Environment.GetEnvironmentVariable("WINDIR")}\system32\reg.exe";
                string psArgs = $"import {rf.FullName}";
                psRegexe.StartInfo.Arguments = psArgs;
                psRegexe.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;

                psRegexe.Start();
                psRegexe.WaitForExit();
                
                if(psRegexe.ExitCode == 0)
                {
                    File.Delete(path);
                    res = true;
                }
                else
                {
                    res = false;
                }
            }
            else
            {
                res = rres;
            }

            return res;
        }    
    }

    public class cpContainers
    {
        public string basePath { get; }
        public string tmpfilepath { get; }

        public cpContainers(string basePath, string tmpfilepath)
        {
            this.basePath = basePath;
            this.tmpfilepath = tmpfilepath;
        }

        public List<string> compareList()
        {
            List<string> res = new List<string>();
            string[] regContainers = { };
            string[] fileContainers = { };
            string tmpfile = tmpfilepath;

            using (RegistryKey HKLM = Registry.LocalMachine)
            {
                using(RegistryKey containers = HKLM.OpenSubKey(basePath))
                {
                    try
                    {
                        regContainers = containers.GetSubKeyNames();
                    }
                    catch
                    {
                        //
                    }
                }
            }

            if (regContainers.Length > 0)
            {
                if (File.Exists(tmpfile))
                {
                    fileContainers = File.ReadAllLines(tmpfile);
                }
                else
                {
                    File.WriteAllLines(tmpfile, regContainers);
                    res = regContainers.ToList<string>();
                    return res;
                }

                int i = 0, j = 0;
                if (fileContainers.Length > 0)
                {
                    foreach (string regItem in regContainers)
                    {
                        j = 0;
                        foreach(string fileItem in fileContainers)
                        {
                            if (regItem.Equals(fileItem))
                            {
                                j++;
                                break;
                            }
                        }

                        if (j > 0)
                        {
                            continue;
                        }
                        else
                        {
                            res.Add(regItem);
                        }

                        i++;
                    }
                }
                else
                {
                    File.WriteAllLines(tmpfile, regContainers);
                    res = regContainers.ToList<string>();
                    return res;
                }
            }

            File.WriteAllLines(tmpfile, regContainers);

            return res;
        }

        public bool instCert(string cpcName)
        {
            bool res = false;
            regConf conf = new regConf();
            string sid = WindowsIdentity.GetCurrent().User.Value;

            using (Process psCspt = new Process())
            {
                psCspt.StartInfo.FileName = conf.getParam("pathCsptest", "");
                string rArgs = $"-property -cinstall -container \"{cpcName}\"";

                psCspt.StartInfo.Arguments = rArgs;
                psCspt.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                psCspt.Start();
                psCspt.WaitForExit();

                if (psCspt.ExitCode == 0)
                {
                    res = true;
                }
                else
                {
                    res = false;
                }
            }

            return res;
        }

    }
}