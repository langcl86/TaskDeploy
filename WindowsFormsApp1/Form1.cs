using System;
using System.Collections.ObjectModel;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Management;
using System.Threading;
using System.Windows.Forms;
using System.Xml;
using System.IO;
using System.Diagnostics;
using System.Text.RegularExpressions;
using System.Management.Automation;

namespace WindowsFormsApp1
{    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            MakeButtons();
            verifyTasks();
        }

        private static string getCIM(string cimclass, string value, string filter = ".")
        {
            List<string> info = new List<string>();

            var thread = new Thread(() =>
            {

            try
            {
                ManagementClass mc = new ManagementClass(cimclass);
                ManagementObjectCollection moc = mc.GetInstances();
                if (moc.Count != 0)
                {
                    foreach (ManagementObject obj in moc)
                    {
                        Match matches = Regex.Match(obj.ToString(), filter, RegexOptions.IgnoreCase);
                        if (!matches.Success) { continue; }

                        info.Add(obj[value].ToString());

                    }
                }

            }

            catch
            {
                info.Add("CIM Query Failed");
            }
            });
            thread.Start();
            thread.Join();
            
            return info.ToArray().FirstOrDefault();
        }


        private void GetDevInfo()
        {
            string cimclass = "Win32_ComputerSystem";
            string name = getCIM(cimclass, "Name");
            string make = getCIM (cimclass, "Manufacturer");
            string model = getCIM (cimclass, "Model");
            string ram = getCIM(cimclass, "TotalPhysicalMemory");
            string dom = getCIM(cimclass, "Domain");
            string cpu = getCIM("Win32_Processor", "Name");
            string bios = getCIM("Win32_BIOS", "Name");
            string serial = getCIM("Win32_BIOS", "SerialNumber");
            string hdd = getCIM("Win32_DiskDrive", "Model", "drive0");
            string os = getCIM("Win32_OperatingSystem", "Caption");
            string osver = getCIM("Win32_OperatingSystem", "Version");
            string build = getCIM("Win32_OperatingSystem", "BuildNumber");
          
            string osinfo = String.Format("\r\n{0}\r\n Ver: {1} \r\n Build: {2}", os, osver, build);

            // Convert RAM bytes to GB      Bytes / 1024^3
            var toGB = float.Parse(ram) / Math.Pow(1024, 3);
            ram = Convert.ToInt32(Math.Round(toGB)).ToString() + " GB";

            List<string> deviceInfo = new List<string>();
            deviceInfo.Add("Hostame: " + name);
            deviceInfo.Add("Domain: " + dom);
            deviceInfo.Add("Manufacturer: " + make);
            deviceInfo.Add("Model: " + model); ;
            deviceInfo.Add("CPU: " + cpu);
            deviceInfo.Add("RAM: " + ram);
            deviceInfo.Add("HDD: " + hdd);
            deviceInfo.Add("BIOS: " + bios);
            deviceInfo.Add("Serial: " + serial);
            deviceInfo.Add(osinfo);

            deviceInfo_textBox.Text = String.Join(Environment.NewLine, deviceInfo);
            NewStatusInfo(String.Join(" ", name, "Is Ready"));
        }
    
        static void FileNotFound (FileInfo file)
        {
            string title = String.Concat(file.Extension, " file not found.");
            string error = String.Concat(file.Name, " file not found.");
            MessageBox.Show(error, title, MessageBoxButtons.OK, MessageBoxIcon.Error);
            //throw new FileNotFoundException(error);
        }

        public void AddNewHistoryEvent (string info)
        {

            if (string.IsNullOrEmpty(info)) { return; }

            MethodInvoker m = new MethodInvoker(() => 
            {
                richTextBox1.Select(0, 0);
                richTextBox1.SelectedText = info + Environment.NewLine;
            });

            try
            {
                if (richTextBox1.InvokeRequired == true)
                {
                    richTextBox1.Invoke(m);
                }
                else
                {
                    Invoke(m);
                }
            }
            catch { }
        }

        private void NewStatusInfo (string info)
        {
            // Status Bar Message
            int strChars = info.Length;
            int maxChars = 35;

            if (strChars > maxChars)
            {
                int r = strChars - maxChars;
                strChars = strChars - r;
                info = info.Substring(0, strChars) + "...";
            }

            MethodInvoker m = new MethodInvoker(() =>
            {
                tooStripStatusLabel1.Text = info;
            });

            if (InvokeRequired)
            {
                tableLayoutPanel2.Invoke(m);
            }
            else
            {
                m.Invoke();
            }
            
            AddNewHistoryEvent(info);
        }

        public static Tool[] xmlTools()
        {
            FileInfo file = new FileInfo("tools.xml");
            
            if (!file.Exists)
            {
                FileNotFound(file);
                throw new FileNotFoundException();
            }

            Tool[] tools;
            try
            {
                XmlDocument xmlfile = new XmlDocument();
                xmlfile.Load(file.FullName);
                XmlNode root = xmlfile.DocumentElement;
                XmlNodeList toolNodes = root.SelectNodes("tool");
                tools = new Tool[toolNodes.Count];
                for (int i = 0; i < toolNodes.Count; i++)
                {
                    XmlNode node = toolNodes[i];
                    XmlNode Name = node.SelectSingleNode("name");
                    XmlNode Path = node.SelectSingleNode("path");

                    tools[i] = new Tool(i, Name.InnerText);
                    tools[i].tPath = Path.InnerText;
                    tools[i].tType = getAttr(Path, "type");
                    tools[i].verify = getAttr(Name, "verify");
                    tools[i].level = getAttr(Name, "level");
                    tools[i].silent = getAttr(Path, "silent") == null;
                    tools[i].status = false;

                }
            }
            catch (Exception ex)
            {
                string error = ex.Message;
                MessageBox.Show(error, "Failed to read XML data", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw new Exception("Invalid XML File");
            }

            return tools;
        }

        private static string getAttr (XmlNode node, string attr)
        {
            if (node.Attributes.Count > 0)
            {
                foreach (XmlAttribute a in node.Attributes)
                {
                    if (a.Name == attr)
                    {
                        return node.Attributes[a.Name].Value;
                    }
                }

            }

            return null;
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private List<Button> getButtons ()
        {
            List<Button> buttons = new List<Button>();
            foreach (Control c in tableLayoutPanel2.Controls.OfType<Button>())
            {
                c.Dock  = DockStyle.Fill;
                buttons.Add((Button)c);
            }

            return buttons;
        }

        Tool[] toolBox = xmlTools();
        private void MakeButtons()
        {
            int c = 0;
            var buttons = getButtons();
    
            foreach (Button b in buttons)
            {
                int id = buttons.IndexOf(b);
                
                try
                {
                    var tool = toolBox[c];
                    tool.buttonName = b.Name;
                    b.Text = tool.tName;
                    b.AutoEllipsis = true;
                    b.Click += (s, e) => button_Click(tool);
                }

                catch
                {
                    b.Text = "empty";
                    b.Enabled = false;
                }
                
               c++;
            }

        }

        private string _currentProcess;
        public string CurrentProcess
        {
            get { return _currentProcess; }
            set { _currentProcess = value; }
        }
        private object pwrShll(string pscmd)
        {  
            // https://docs.microsoft.com/en-us/dotnet/api/system.management.automation.powershell?view=powershellsdk-1.1.0
           
            using (PowerShell ps = PowerShell.Create())
            {
                ps.AddScript("Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process");
                ps.AddScript(pscmd);

                PSDataCollection<PSObject> output = new PSDataCollection<PSObject>();

                output.DataAdded += new EventHandler<DataAddedEventArgs>(Output_DataAdded);
                ps.InvocationStateChanged += new EventHandler<PSInvocationStateChangedEventArgs>(Ps_InvocationStateChanged);

                try
                {
                    IAsyncResult result = ps.BeginInvoke<PSObject, PSObject>(null, output);
                                       
                    while (!result.IsCompleted)
                    {
                        // Wait for exit
                        Thread.Sleep(1000);
                    }

                    if (ps.Streams.Error != null)
                    {
                        foreach (ErrorRecord e in ps.Streams.Error)
                        {
                            NewStatusInfo("FAILED " + CurrentProcess);
                            AddNewHistoryEvent(e.ToString());
                            return null;
                        }
                    }

                    if (ps.InvocationStateInfo.State == PSInvocationState.Completed)
                    {
                        verifyTasks();
                    }

                   return output.LastOrDefault();                   
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Powershell request failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return null;
                }
            }
        }

        private void Ps_InvocationStateChanged(object sender, PSInvocationStateChangedEventArgs e)
        {
            PSInvocationState state = e.InvocationStateInfo.State;
            string statusStr = String.Format("{0} {1}", state.ToString(), CurrentProcess);
            NewStatusInfo(statusStr);

                switch (state)
                {
                case PSInvocationState.Running:
                    FormReady(false);
                    break;

                case PSInvocationState.Completed:
                case PSInvocationState.Stopped:
                case PSInvocationState.Failed:
                    FormReady(true);
                    break;

                default:
                    return;
                }
        }
        private void Output_DataAdded(object sender, DataAddedEventArgs e)
        {
            PSDataCollection<PSObject> psobj = (PSDataCollection<PSObject>)sender;
            Collection<PSObject> results = psobj.ReadAll();
            foreach (PSObject result in results)
            {
                if(result == null ) continue;

                Invoke((MethodInvoker)delegate
                {
                    richTextBox1.Select(0, 0);
                });

                AddNewHistoryEvent(result.ToString());
            }
        }

        private ProcessStartInfo BaseStartInfo ()
        {
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.UseShellExecute = false;
            startInfo.Verb = "RunAs";
            startInfo.RedirectStandardOutput = true;
            startInfo.RedirectStandardError = true;
            return startInfo;
        }
        private List<string> StartNewProcess(Process p)
        {
            List<string> output = new List<string>();

            var thread = new Thread(() =>
            {
                p.Start();
                while (!p.StandardOutput.EndOfStream)
                {
                    output.Add(p.StandardOutput.ReadLine());  
                }

                while (!p.StandardError.EndOfStream)
                {
                    output.Add(p.StandardError.ReadLine());
                }
            });

            thread.Start();
            thread.Join();
            return output;
        }

        private void button_Click(Tool tool)
        {
            Process p = new Process();
            string filename;
            string args = null;
            CurrentProcess = tool.tName;

            string type = null;
            if (tool.tType != null)
            {
                type = tool.tType.ToUpper();
            }

            string pscmds = null;
            var t = new Thread(() =>
            {
                pwrShll(pscmds);
            });

            switch (type)
            {
                case "PS":
                    if (!tool.silent)
                    {
                        filename = "C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe";
                        args = "-ep bypass -file " + tool.tPath;
                        break;
                    }

                    try
                    {
                        FileInfo path = new FileInfo(tool.tPath);
                        if (path.Extension != ".ps1") { throw new Exception(); }
                        pscmds = path.FullName.Replace(" ", "` ");
                    }
                    catch
                    {
                        pscmds = tool.tPath;
                    }
                    t.Start();
                    return;

                case "CMD":
                    string cmdexe = Path.Combine(Environment.SystemDirectory.ToString(), "cmd.exe");
                    pscmds = String.Format("{0} /c '{1}'; if ($LASTEXITCODE -GT 0) {{ Write-Error \"Exit Code: $LASTEXITCODE\"; throw; }}", cmdexe, tool.tPath);
                    t.Start();
                    return;
                    //filename = "C:\\Windows\\System32\\cmd.exe";
                    //args = "/c " + tool.tPath;
                    //break;

                case null:
                default:
                    FileInfo fi = new FileInfo(tool.tPath);
                    if (!fi.Exists)
                    {
                        FileNotFound(fi);
                    }
                    else
                    {
                        pscmds = String.Format("Start-Process -FilePath \"{0}\" -Wait; if ($LASTEXITCODE -GT 0) {{ Write-Error \"Exit Code: $LASTEXITCODE\"; throw; }}", tool.tPath);
                        t.Start();
                    }
                    return;
            }
          
            ProcessStartInfo startInfo = BaseStartInfo();
            startInfo.FileName = filename;
            startInfo.Arguments = args;
            p.StartInfo = startInfo;

            try
            {
                FormReady(false);
                string eventInfo = String.Join(" ", "Running process", tool.tName);
                NewStatusInfo(eventInfo);

                List<string> output = StartNewProcess(p);
                if (p.HasExited)
                {
                    AddNewHistoryEvent(String.Join(Environment.NewLine, output));
                    eventInfo = eventInfo.Replace("Running", "Exited");
                    NewStatusInfo(eventInfo);
                    Thread.Sleep(1800);
                    FormReady(true);
                }

                if (p.ExitCode != 0)
                {
                    string message = String.Format("{0} closed abnormally. \r\n\r\n Arguments: {1} \r\n\r\n Output: {2} \r\n\r\n Exit Code: {3}", Path.GetFileName(filename), args, string.Join("\n", output), p.ExitCode); ;
                    MessageBox.Show(message, "DOh!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                else
                {
                    verifyTasks();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Failed to launch " + tool.tName, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private List<string> PS_RunCMD (string pscmd, string script = null)
        {
            string args = pscmd;
            if (script != null)
            {
                // If script is null run PS command ~ else run script with pscmd as arguments
                args = String.Format("powershell -ep bypass -file \"{0}\" {1}", script.Replace(" ", "` "), pscmd);
            }

            Process p = new Process();
            ProcessStartInfo startInfo = BaseStartInfo();
            startInfo.FileName = "powershell.exe";
            startInfo.Arguments = args;
            startInfo.CreateNoWindow = true;
            p.StartInfo = startInfo;

            try
            {   
                List<string> output = StartNewProcess(p);
                string outString = String.Join(" ", output);
                AddNewHistoryEvent(String.Join(" - ", pscmd, outString));

                if (p.ExitCode != 0)
                {
                    string message = String.Format("Powershell closed abnormally. \r\n\r\n Arguments: {0} \r\n\r\n Output: {1} \r\n\r\n Exit Code: {2}", pscmd, string.Join("\n",output), p.ExitCode); ;
                    MessageBox.Show(message, "DOh!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return null;
                }
                return output;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Failed to launch " + pscmd, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }
        public void setButtons ()
        {

            foreach (Tool tool in toolBox)
            {
                Control b;


                try
                {
                    b = Controls.Find(tool.buttonName, true)[0];
                }
                catch
                {
                    continue;
                }

                bool priority = tool.level == "high";

                // Default properties
                b.BackColor = DefaultBackColor;
                b.Text = tool.tName;
                b.Enabled = true;

                if (priority)
                {
                    b.BackColor = Color.Tomato;
                }
                //

                if (tool.status != false)
                {
                    // Task Not Completed
                    b.BackColor = Color.LimeGreen;
                    b.Enabled = false;
                }

                if (priority)
                {
                    // Button Always Enabled
                    b.Enabled = true;
                }

            }

        }

        private Tool getTool(string str)
        {
            Tool t = toolBox.Where(x => x.tName.ToLower().Contains(str.ToLower())).FirstOrDefault();
            return t;
        }
        public void FormReady (bool state)
        {
            if (tableLayoutPanel2.InvokeRequired == true)
            {
                tableLayoutPanel2.Invoke((MethodInvoker)delegate
                {
                    tableLayoutPanel2.Enabled = state; 
                });
            }

            else
            {
                tableLayoutPanel2.Enabled = state;
            }
        }

        private void verifyTasks ()
        {
            new Thread(() =>
           {
               FileInfo script = new FileInfo("verify.ps1");   // Powershell Script  Performing Task Verification

               if (!script.Exists) { FileNotFound(script); }

               try
               {
                   NewStatusInfo("Running Checks...");
                   FormReady(false);

                   foreach (Tool t in toolBox)
                   {
                       if (t.verify != null)
                       {
                           var checkTool = PS_RunCMD(t.verify, script.FullName).ToArray();

                           //AddNewHistoryEvent(t.verify + "...");
                           //var checkTool = pwrShll(scriptStr, t.verify);
                           string output = String.Join("\n", checkTool);
                           string info = Regex.Replace(output, "\\[..+]", "");

                           // On  passing verification PS script returns [DONE]  

                           if (info.Length > 0)
                           {
                               t.tName = info;              // Send Output To Button When Info != null
                           }

                           if (!output.ToLower().Contains("[done]"))
                           {
                               t.status = false;           // Task Not Completed
                           }
                           else
                           {
                               t.status = true;            // Task Completed
                           }
                       }
                   }

                   Invoke(new MethodInvoker(setButtons));
                   Invoke(new MethodInvoker(GetDevInfo));
                   FormReady(true);
               }
               catch { }
           }).Start();

        }

        private void refreshStatusToolStripMenuItem_Click(object sender, EventArgs e)
        {
            verifyTasks();
            NewStatusInfo("Refresh Complete");
        }

        private void exitToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            this.Close();
        }

        private void statusPanel_lic_Paint(object sender, PaintEventArgs e)
        {
        }

        private void refreshStatusToolStripMenuItem_MouseDown(object sender, MouseEventArgs e)
        {
        }
        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void testToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var t = new Thread(() =>
            {
                FileInfo fi = new FileInfo("verify.ps1");
                string pscmd = string.Join(" ", fi.FullName, "checkPnP");
                pwrShll("Start-Process notepad.exe -Wait");
            });
            t.Start();
        }
    }
}
