using System;
using System.DirectoryServices;
using System.DirectoryServices.ActiveDirectory;
using System.DirectoryServices.AccountManagement;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Net.NetworkInformation;
using System.Linq;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AddToWiFi
{
    public partial class Form1 : Form
    {
        private string LDAPDomainName = string.Empty;

        public Form1()
        {
            InitializeComponent();

            if (!CheckIfInDomain())
            {
                MessageBox.Show("You are not part of a domain.\nThis application will now terminate.");
                Environment.Exit(1);
            }

            LDAPDomainName = GetDomainDN(IPGlobalProperties.GetIPGlobalProperties().DomainName);

            openFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        }
        public bool CheckIfInDomain()
        {
            bool isInDomain = true;

            try
            {
                Domain.GetComputerDomain();
            }
            catch (ActiveDirectoryObjectNotFoundException)
            {
                isInDomain = false;
            }

            return isInDomain;
        }
        private string GetDomainDN(string domain)
        {
            DirectoryContext context = new DirectoryContext(DirectoryContextType.Domain, domain);
            Domain d = Domain.GetDomain(context);
            DirectoryEntry de = d.GetDirectoryEntry();
            return de.Properties["DistinguishedName"].Value.ToString();
        }

        private void btnOpenFile_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
        }

        private void AddUserToWiFi(string[] Usernames)
        {
            try
            {
                using (PrincipalContext pc = new PrincipalContext(ContextType.Domain, "INTERNET.LOCAL"))
                {
                    GroupPrincipal group = GroupPrincipal.FindByIdentity(pc, "Guest-WiFi");
                    foreach (string user in Usernames)
                    {
                        try
                        {
                            group.Members.Add(pc, IdentityType.SamAccountName, user);

                            //only runs if no exception caught
                            Output(string.Format("+ {0} blev tilføjet til Guest-WiFi", user));
                        }
                        catch (PrincipalExistsException)
                        {
                            Output(string.Format("- {0} er allerede medlem af Guest-WiFi", user));
                        }
                        catch (NoMatchingPrincipalException)
                        {
                            Output(string.Format("% {0} findes ikke", user));
                        }
                    }
                    group.Save();
                }
            }
            catch (System.DirectoryServices.DirectoryServicesCOMException E)
            {
                Output(E.Message.ToString());
            }
        }

        private void Output(string message)
        {
            txtOutput.AppendText(message + Environment.NewLine);
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            ReadFile(openFileDialog1.FileName);
        }

        private void ReadFile(string FileName)
        {
            switch (Path.GetExtension(FileName))
            {
                case ".csv":
                    List<string> lines = new List<string>(File.ReadAllLines(FileName));
                    List<string> userList = new List<string>();
                    lines.RemoveAt(0);

                    foreach (string line in lines)
                    {
                        int indexOfSeperator = line.IndexOf(';');

                        if (indexOfSeperator != -1)
                        {
                            userList.Add(line.Substring(0, indexOfSeperator));
                        }
                        else
                        {
                            userList.Add(line);
                        }
                    }

                    AddUserToWiFi(userList.ToArray());
                    break;
                case ".txt":
                    AddUserToWiFi(File.ReadAllLines(FileName));
                    break;
                default:
                    break;
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            Environment.Exit(1);
        }

        private void Form1_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var files = (string[])e.Data.GetData(DataFormats.FileDrop);

                foreach (var file in files)
                {
                    if (Path.GetExtension(file) == ".txt" || Path.GetExtension(file) == ".csv")
                    {
                        ReadFile(file);
                    }
                }
            }
        }

        private void Form1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.Copy;
            else
                e.Effect = DragDropEffects.None;
        }
    }
}