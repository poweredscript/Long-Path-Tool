using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Principal;
using System.Windows.Forms;
using IWshRuntimeLibrary;
using Microsoft.Win32;
using Ookii.Dialogs;

namespace Long_Path_Tool
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            if (IsRunAsAdministrator())
            {
                RegistyCheck();
                Close();
                //Environment.Exit(0);
                return;
            }
            var commandLineArgs = Environment.GetCommandLineArgs().ToList();
            if (commandLineArgs.Count > 2)
            {
                FileOrFolderList.AddRange(commandLineArgs.GetRange(1, commandLineArgs.Count - 1));
            }
            else if (commandLineArgs.Count == 2)
            {
                FileOrFolderList.Add(commandLineArgs[1]);
            }
            InitTimer();
            //MessageBox.Show(FileOrFolderList.Count.ToString());
        }
        private System.Windows.Forms.Timer _timer1;
        public void InitTimer()
        {
            _timer1 = new System.Windows.Forms.Timer();
            _timer1.Tick += new EventHandler(timer1_Tick);
            _timer1.Interval = 1000; // in miliseconds
            _timer1.Start();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            label1.Text = @"Item Count: " + FileOrFolderList.Count;
        }
        private static readonly List<string> FileOrFolderList = new List<string>();
        private void buttonMove_Click(object sender, EventArgs e)
        {
            string tagetPath = "";
            using (var dialog = new VistaFolderBrowserDialog())
            {
                dialog.SelectedPath = "";
                dialog.UseDescriptionForTitle = true;
                dialog.ShowNewFolderButton = true;
                dialog.Description = @"Select target folder";
                //dialog.RootFolder = Environment.SpecialFolder.Desktop;
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    tagetPath = dialog.SelectedPath;
                }
            }
            if (Pri.LongPath.Directory.Exists(tagetPath))
            {
                foreach (var s in FileOrFolderList)
                {
                    if (Pri.LongPath.Directory.Exists(s) || Pri.LongPath.File.Exists(s))
                    {
                        MoveTo(s, tagetPath);
                    }
                }
                MessageBox.Show(@"Success", @"Success move " + FileOrFolderList.Count + @" item", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            Close();
            //Environment.Exit(0);
        }

        private void buttonDelete_Click(object sender, EventArgs e)
        {
            foreach (var s in FileOrFolderList)
            {
                if (Pri.LongPath.Directory.Exists(s) || Pri.LongPath.File.Exists(s))
                {
                    Delete(s);
                }
            }
            MessageBox.Show(@"Success", @"Success delete " + FileOrFolderList.Count + @" item", MessageBoxButtons.OK, MessageBoxIcon.Information);
            Close();
            //Environment.Exit(0);
        }


     //   string _folderName = "c:\\dinoch";
        //private void buttonBrowse_Click(object sender, EventArgs e)
        //{
        //    var browseForFolder = new BrowseForFolder();
        //    var selectFolder = browseForFolder.SelectFolder("Select a file/folder", "", Handle);
        //    if (!string.IsNullOrEmpty(selectFolder))
        //    {
        //        FileOrFolderList.Add(selectFolder);
        //    }
        //}


        public void MoveTo(string sourcePath, string tagetPath)
        {
            if (IsFolder(sourcePath))
            {
                Pri.LongPath.DirectoryInfo directoryInfo = new Pri.LongPath.DirectoryInfo(sourcePath);
                directoryInfo.MoveTo(Pri.LongPath.Path.Combine(tagetPath, directoryInfo.Name));
            }
            else
            {
                Pri.LongPath.FileInfo fileInfo = new Pri.LongPath.FileInfo(sourcePath);
                fileInfo.MoveTo(Pri.LongPath.Path.Combine(tagetPath, fileInfo.Name));
            }
        }


        public void Delete(string sourcePath)
        {
            if (IsFolder(sourcePath))
            {
                Pri.LongPath.DirectoryInfo directoryInfo = new Pri.LongPath.DirectoryInfo(sourcePath);
                directoryInfo.Delete(true);
            }
            else
            {
                Pri.LongPath.FileInfo fileInfo = new Pri.LongPath.FileInfo(sourcePath);
                fileInfo.Delete();
            }
        }
        public bool IsFolder(string path)
        {
            return Pri.LongPath.Directory.Exists(path);
        }

        public static void RegistyCheck()
        {
            var fullName = System.Reflection.Assembly.GetExecutingAssembly().Location;
            var thisPath = Directory.GetCurrentDirectory();
            //var iconPath = Path.Combine(thisPath, "Note.ico");
            var iconPathReg = "\"" + fullName + "\"" + "";
            var exePathReg = "\"" + fullName + "\"" + " \"%1\"";
            var exePathReg2 = "\"" + fullName + "\"" + " \"%V\"";
            //File
            Registry.SetValue(@"HKEY_CLASSES_ROOT\*\shell\Long Path Tool", "Icon", iconPathReg, RegistryValueKind.String);
            Registry.SetValue(@"HKEY_CLASSES_ROOT\*\shell\Long Path Tool\command", "", exePathReg, RegistryValueKind.String);
            //Folder
            Registry.SetValue(@"HKEY_CLASSES_ROOT\Directory\shell\Long Path Tool", "Icon", iconPathReg, RegistryValueKind.String);
            Registry.SetValue(@"HKEY_CLASSES_ROOT\Directory\shell\Long Path Tool\command", "", exePathReg, RegistryValueKind.String);
            //Inside folder
            Registry.SetValue(@"HKEY_CLASSES_ROOT\Directory\Background\shell\Long Path Tool", "Icon", iconPathReg, RegistryValueKind.String);
            Registry.SetValue(@"HKEY_CLASSES_ROOT\Directory\Background\shell\Long Path Tool\command", "", exePathReg2, RegistryValueKind.String);

            var schPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Microsoft\Windows\SendTo\Long Path Tool.lnk";
            //MessageBox.Show(fullName);
            //MessageBox.Show(thisPath);
            //MessageBox.Show(schPath);
            //MessageBox.Show(fullName);
            if (System.IO.File.Exists(schPath)) System.IO.File.Delete(schPath);
            WshShellClass wsh = new WshShellClass();
            IWshRuntimeLibrary.IWshShortcut shortcut = wsh.CreateShortcut(schPath) as IWshRuntimeLibrary.IWshShortcut;
            shortcut.Arguments = "";
            shortcut.TargetPath = fullName;
            // not sure about what this is for
            shortcut.WindowStyle = 1;
            shortcut.Description = "Long Path Tool";
            shortcut.WorkingDirectory = thisPath;
            shortcut.IconLocation = fullName;
            shortcut.Save();
            MessageBox.Show("Success Install", "Success Install", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private static bool IsRunAsAdministrator()
        {
            WindowsIdentity wi = WindowsIdentity.GetCurrent();
            WindowsPrincipal wp = new WindowsPrincipal(wi);

            return wp.IsInRole(WindowsBuiltInRole.Administrator);
        }
    }
}
