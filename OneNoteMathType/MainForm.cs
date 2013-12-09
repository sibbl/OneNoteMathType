using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using MTSDKDN;
using on = Microsoft.Office.Interop.OneNote;

namespace OneNoteMathType
{
    public partial class MainForm : Form
    {
        #region Configuration
        public const string EquationStartEndString = "$$";
        public const int HotkeyModifier = MOD_WIN;
        public const Keys Hotkey = Keys.Y;
        #endregion

        #region Hooks
        [DllImport("user32.dll", EntryPoint = "FindWindowEx")]
        public static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);
        [DllImport("User32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int uMsg, int wParam, string lParam);

        public IntPtr FindMathTypeWindow()
        {
            Process[] notepads = Process.GetProcessesByName("notepad");
            if (notepads.Length == 0) return IntPtr.Zero;
            if (notepads[0] != null)
            {
                IntPtr child = FindWindowEx(notepads[0].MainWindowHandle, new IntPtr(0), "Edit", null);
                if (child != IntPtr.Zero) return child;
            }
            return IntPtr.Zero;
        }
        #endregion

        #region UI
        public MainForm()
        {
            InitializeComponent();
            Hide();
        }
        private void MainForm_Load(object sender, EventArgs e)
        {
            activeToolStripMenuItem.Checked = true;
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            UnbindHotkey();
        }
        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
        private void activeToolStripMenuItem_CheckedChanged(object sender, EventArgs e)
        {
            if (activeToolStripMenuItem.Checked) BindHotkey();
            else UnbindHotkey();
        }
        #endregion

        #region Translate equations in OneNote with MathType

        private static XNamespace ns;
        public static void TranslateCurrentOneNotePage()
        {
            var onenoteApp = new on.Application();

            string notebookXml;
            onenoteApp.GetHierarchy(null, on.HierarchyScope.hsPages, out notebookXml);
            var doc = XDocument.Parse(notebookXml);
            ns = doc.Root.Name.Namespace;
            var pageNode = doc.Descendants(ns + "Page").Where(n =>
              n.Attribute("isCurrentlyViewed") != null && n.Attribute("isCurrentlyViewed").Value == "true").FirstOrDefault();
            
            if (pageNode == null) return;
            var existingPageId = pageNode.Attribute("ID").Value;
            
            string content;
            onenoteApp.GetPageContent(existingPageId, out content, on.PageInfo.piAll);

            var xDoc = XDocument.Parse(content);
            var textElements = xDoc.Descendants(ns + "T").ToList();
            foreach(var item in textElements)
            {
                if(item != null) ReplaceEquationsInXElement(item);
            }

            onenoteApp.UpdatePageContent(xDoc.ToString(), DateTime.MinValue);
        }

        public static string MMLStart = "<mml:math";
        public static string MMLEnd = "</mml:math>";

        public static void ReplaceEquationsInXElement(XElement content)
        {
            var equationParts = content.Value.Split(new string[] { EquationStartEndString }, StringSplitOptions.None);
            if (equationParts.Count() < 2) return;

            var newContent = new StringBuilder();
            var count = equationParts.Count();
            var isEquation = true; //set true, which will be set false before checking first item
            var isUnfinishedEquation = false;
            for (var i = 0; i < count; i++)
            {
                if (i >= 1 && equationParts[i - 1].EndsWith("\\")) continue;
                isEquation = !isEquation;
                if (isEquation && i == count - 1)
                    isUnfinishedEquation = true;

                if (isEquation && !isUnfinishedEquation)
                {
                    var equationStr = TranslateEquation(equationParts[i]);

                    //use string modification to get only "<math:mml .... </math:mml>"
                    var startIndex = equationStr.IndexOf(MMLStart, StringComparison.InvariantCultureIgnoreCase);
                    var endIndex = equationStr.IndexOf(MMLEnd, StringComparison.InvariantCultureIgnoreCase);
                    equationStr = equationStr.Substring(startIndex, endIndex - startIndex + MMLEnd.Length);
                    newContent.AppendFormat("</span><!--[if mathML]>{0}<![endif]--><span>", equationStr);
                }
                else if (isUnfinishedEquation)
                    newContent.AppendFormat(EquationStartEndString + equationParts[i]);
                else
                    newContent.AppendFormat(equationParts[i]);
            }
            content.Value = "<span>"+newContent.ToString()+"</span>";
        } 

        public static string TranslateEquation(string eq)
        {
            var inputFile = Path.GetTempPath() + Guid.NewGuid() + ".tex";
            using (var fs = new StreamWriter(inputFile,false))
            {
                fs.Write("$$"+eq+"$$");
            }
            var outputFile = Path.GetTempPath() + Guid.NewGuid() + ".mml";
            var ce = new ConvertEquation();
            ce.Convert(new EquationInputFileText(inputFile, ClipboardFormats.cfTeX),
                       new EquationOutputFileText(outputFile, "MathML2 (DataObject).tdl"));
            string result;
            using (var fs = new StreamReader(outputFile, true))
            {
                 result = fs.ReadToEnd();
            }
            File.Delete(inputFile);
            File.Delete(outputFile);
            return result;
        }
        #endregion

        #region Catch hotkeys

        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, int fsModifiers, int vk);

        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        public const int MOD_ALT = 0x1;
        public const int MOD_CONTROL = 0x2;
        public const int MOD_SHIFT = 0x4;
        public const int MOD_WIN = 0x8;
        public const int WM_HOTKEY = 0x312;

        private void BindHotkey()
        {
            RegisterHotKey(this.Handle, 1, HotkeyModifier, (int)Hotkey);
        }
        private void UnbindHotkey()
        {
            UnregisterHotKey(this.Handle, 1);
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_HOTKEY && (int)m.WParam == 1)
                TranslateCurrentOneNotePage();
            base.WndProc(ref m);
        }

        #endregion
    }
}
