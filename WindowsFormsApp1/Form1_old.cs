using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Management;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Runtime;
using System.Reflection;
using System.IO;
using System.Diagnostics;
using System.Text.RegularExpressions;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            tooStripStatusLabel1.Text = DeviceInfo();
        //    setButtons();
            MakeButtons();
            statusCheck();

            Tool[] tools = xmlTools();
            foreach(Tool tool in tools)
            {
                if (tool == null) { continue; }
            }

        }

        private string DeviceInfo ()
        {
            string result = null;
            // create management class object
            ManagementClass mc = new ManagementClass("Win32_ComputerSystem");
            //collection to store all management objects
            ManagementObjectCollection moc = mc.GetInstances();
            if (moc.Count != 0)
            {
                foreach (ManagementObject obj in mc.GetInstances())
                {
                    result = String.Join(" ", obj["manufacturer"].ToString().Split()[0], obj["Model"].ToString());
                }
            }

            string hostname = Environment.MachineName.ToString();

            result = String.Format("{00} on {01}", hostname , result);

            return result;
        }

        public static Tool[] xmlTools ()
        {
            FileInfo file = new FileInfo("tools.xml");

            if (!file.Exists)
            {
                string error = "XML File Not Found!";
                MessageBox.Show(error, "Doh!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw new FileNotFoundException(error);
            }

            Dictionary<string, string> items = new Dictionary<string, string>();
            try
            {
                XmlDocument xmlfile = new XmlDocument();
                xmlfile.Load(file.FullName);
                XmlNode root = xmlfile.DocumentElement;
                XmlNodeList toolNodes = root.SelectNodes("tool");
                Tool[] tools = new Tool[toolNodes.Count];
                for (int i = 0; i < toolNodes.Count; i++)
                {
                    string name = toolNodes[i].SelectSingleNode("name").InnerText;
                    string path = toolNodes[i].SelectSingleNode("path").InnerText;
                    tools[i] = new Tool(i, name, path);
                }

                return tools;
            }
            catch (Exception ex)
            {
                string error = ex.Message;
                MessageBox.Show(error, "Failed to read XML data", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw new Exception("Invalid XML File");
            }

            return null;
        }
        
        private Dictionary<string, string> readXML ()
        {
            FileInfo file = new FileInfo("tools.xml");

            if (!file.Exists)
            {
                string error = "XML File Not Found!";
                MessageBox.Show(error, "Doh!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw new FileNotFoundException(error);
            }

            Dictionary<string, string> items = new Dictionary<string, string>();
            try
            {
                XmlDocument xmlfile = new XmlDocument ();
                xmlfile.Load(file.FullName);
                XmlNode root = xmlfile.DocumentElement;
                XmlNodeList toolNodes = root.SelectNodes("tool");
                
                foreach (XmlNode node in toolNodes)
                {
                    // search for child nodes to collect name & path
                    string name = node.SelectSingleNode("name").InnerText;
                    string path = node.SelectSingleNode("path").InnerText;
                    items.Add(name, path);

                }
                
            }
            catch (Exception ex)
            {
                string error = ex.Message;
                MessageBox.Show(error, "Failed to read XML data", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw new Exception("Invalid XML File");
            }

            return items;
          
        }
        
        private string ToolName (int id)
        {
            string[] names = readXML().Keys.ToArray();

            try
            {
                return names[id];
            }
            catch
            {
                return null;
            }
        }

        private string ToolPath (int id)
        {
            string[] paths = readXML().Values.ToArray();

            try
            {
                return paths[id];
            }
            catch
            {
                return null;
            }
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

        private void MakeButtons()
        {
            Tool[] toolBox = xmlTools();
            var buttons = getButtons();
            int i = 0;
            foreach (Button b in buttons)
            {
                int id = buttons.IndexOf(b);
                var tool = toolBox[i];

               if (String.IsNullOrEmpty(tool.tName))
                {
                    b.Text = "empty";
                    b.Enabled = false;
                }

               else
                {
                    b.Text = tool.tName;
                    b.Name = tool.buttonName;
                    b.Click += (s, e) => button_Click(tool.tPath);
                }
               i++;
            }

        }
        
        private void setButtons ()
        {
            List<Button> buttons = getButtons();
            foreach (Button b in buttons)
            {
                int id = buttons.IndexOf(b);
                string name = ToolName(id);

                if (string.IsNullOrEmpty(name))
                {
                    b.Text = "empty";
                    b.Enabled = false;
                }

                else
                {
                    b.Text = name;
//                    b.Click += (s, e) => { button_Click(id); };

                }

            }
            
        }
        
        private void button_Click(string path)
        {
            //FileInfo fileInfo = new FileInfo(ToolPath(id));

            FileInfo fileInfo = new FileInfo(path);

            if (fileInfo.Exists)
            {
                Process p = new Process();

                /*
                //
                */

                int timeout;
                string filename;
                string args = null;
                bool silent = false;

                if (fileInfo.Extension == ".ps1")
                {
                    filename = "C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe";
                    args = "-ep bypass -file " + fileInfo.FullName;
                    silent = true;
                    timeout = 300000;
                }

                else
                {
                    filename = fileInfo.FullName;
                    timeout = 900000;
                }

                p.StartInfo.FileName = filename;
                p.StartInfo.Arguments = args;
                p.StartInfo.CreateNoWindow = silent;
                p.StartInfo.UseShellExecute = false;
                p.StartInfo.RedirectStandardOutput = false;
                p.StartInfo.RedirectStandardError = false;

                p.Start();
                p.WaitForExit(timeout);

                if (p.HasExited)
                {
                    statusCheck();
                }
                
            }

            else
            {
                MessageBox.Show("\"" + fileInfo.Name + "\" was not found", "File not found", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private List<string> PS_RunCMD (string pscmd)
        {
            Process p = new Process();
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.FileName = "powershell.exe";
            startInfo.UseShellExecute = false;
            startInfo.RedirectStandardOutput = true;
            startInfo.RedirectStandardError = true;
            startInfo.Arguments = pscmd;
            startInfo.CreateNoWindow = true;

            p.StartInfo = startInfo;
            p.Start();
            List<string> output = new List<string>();
            while (!p.StandardOutput.EndOfStream)
            {
                output.Add(p.StandardOutput.ReadLine());
            }

            return output;

        }

        private void checkDrivers()
        {
            List<string> output = PS_RunCMD("(Get-PnPDevice -Status 'Error').Caption");

            int c = output.Count;
            if (c > 0)
            {
                richTextBox1.BackColor = Color.Tomato;
                richTextBox1.Text = c + " Driver Issues Found\r\n\r";
                foreach (string line in output)
                {
                    richTextBox1.Text += line;
                }
            }

            else
            {
                richTextBox1.BackColor = Color.LimeGreen;
                richTextBox1.Text = "No Driver Issues Found";
            }
        }

        private void enableButton (string keyword, bool state)
        {
            foreach (Control c in tableLayoutPanel2.Controls.OfType<Button>())
            {
                string ButtonText = c.Text.ToLower();
                keyword = keyword.ToLower();

                if (ButtonText.Contains(keyword))
                {
                    c.Enabled = state;
                }
            }
        }

        private void statusCheck ()
        {
            
            string domain = System.Net.NetworkInformation.IPGlobalProperties.GetIPGlobalProperties().DomainName;
            string[] SolarWinds = PS_RunCMD("(Get-ChildItem ${env:ProgramFiles(x86)} | Where-Object Name -match 'Take Control').Name").ToArray();
            string[] BitDefender = PS_RunCMD("(Get-ChildItem $env:ProgramFiles | Where-Object Name -match 'BitDefender').Name").ToArray();

            checkDrivers();
            
            if (domain.Length > 0)
            {
                enableButton("domain", false);
                statusLabel_domain.Text = statusLabel_domain.Text.Replace("Not", "Is");
                statusPanel_domain.BackColor = Color.LimeGreen;
            }

            if (SolarWinds.Length > 0)
            {
                enableButton("SolarWinds", false);
                statusLabel_SW.Text = statusLabel_SW.Text.Replace("Not", "Is");
                statusPanel_SW.BackColor = Color.LimeGreen;
            }

            if (BitDefender.Length > 0)
            {
                enableButton("BitDefender", false);
                statusLabel_BEST.Text = statusLabel_BEST.Text.Replace("Not", "Is");
                statusPanel_BEST.BackColor = Color.LimeGreen;
            }

            var slmgr = PS_RunCMD("cscript /NoLogo 'c:\\Windows\\System32\\slmgr.vbs' /dli");
            string[] licinfo = slmgr.ToArray().Where(x => !string.IsNullOrEmpty(x)).ToArray();
            bool activated = slmgr[4].Contains("Licensed");

            statusLabel_lic.Text = String.Format("{0} \r {1}", slmgr[3], slmgr[4]);

            if (activated)
            {
                statusPanel_lic.BackColor = Color.LimeGreen;
            }
            else
            {
                statusPanel_lic.BackColor = Color.Tomato;
            }
          }

    }
}
