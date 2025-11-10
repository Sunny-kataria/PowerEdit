//using Microsoft.Office.Interop.Word;
//using Microsoft.Office.Tools;
//using Microsoft.Web.WebView2.Core;
//using Microsoft.Web.WebView2.WinForms;
//using System;
//using System.IO;
//using System.Threading.Tasks;
//using Office = Microsoft.Office.Core;
//using Word = Microsoft.Office.Interop.Word;
//using System.Collections.Generic;
//using System.Runtime.InteropServices;
//using System.Text.RegularExpressions;



//namespace PowerEditAddIn
//{
//    public partial class ThisAddIn
//    {
//        private CustomTaskPane _pane;
//        private DesignerPane _control;

//        private async void ThisAddIn_Startup(object sender, System.EventArgs e)
//        {
//            _control = new DesignerPane();
//            _pane = this.CustomTaskPanes.Add(_control, "PowerEdit");
//            _pane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionLeft;
//            _pane.Width = 420;
//            _pane.Visible = false;

//            // === MOVE INIT HERE ===
//            var dataFolder = Path.Combine(
//                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
//                "PowerEditAddin", "WebView2Data");
//            Directory.CreateDirectory(dataFolder);

//            var env = await CoreWebView2Environment.CreateAsync(
//                userDataFolder: dataFolder);

//            await _control.webView21.EnsureCoreWebView2Async(env);

//            string basePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ui");
//            _control.webView21.CoreWebView2.SetVirtualHostNameToFolderMapping(
//                "app.local", basePath,
//                CoreWebView2HostResourceAccessKind.Allow);

//            _control.webView21.CoreWebView2.Navigate("https://app.local/index.html");
//            _control.webView21.CoreWebView2.WebMessageReceived += (s, args) =>
//            {
//                var doc = System.Text.Json.JsonDocument.Parse(args.WebMessageAsJson);
//                var root = doc.RootElement;
//                var type = root.GetProperty("type").GetString();

//                if (type == "browse")
//                    BrowseAndOpenDocx();

//                else if (type == "runAction")
//                {
//                    var action = root.GetProperty("action").GetString();
//                    RunSelectedAction(action);
//                }
//            };

//        }

//        public void TogglePane()
//        {
//            if (_pane == null)
//            {
//                System.Windows.Forms.MessageBox.Show("Pane is NULL!");
//                return;
//            }

//            _pane.Visible = !_pane.Visible;
//        }
//        public void BrowseAndOpenDocx()
//        {
//            try
//            {
//                using (var dlg = new System.Windows.Forms.OpenFileDialog())
//                {
//                    dlg.Filter = "Word Documents|*.docx";
//                    dlg.Title = "Select a Word Document";

//                    if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
//                    {
//                        var app = this.Application;

//                        // 1) Close current active doc WITHOUT SAVE silently
//                        try
//                        {
//                            if (app.Documents.Count > 0)
//                            {
//                                app.ActiveDocument.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
//                            }
//                        }
//                        catch { }

//                        // 2) Open selected file in SAME Word instance
//                        var doc = app.Documents.Open(dlg.FileName, ReadOnly: false, Visible: true);

//                        // 3) Re-show pane because sometimes Word hides it on file open
//                        _pane.Visible = true;

//                        NotifyClient("Opened: " + Path.GetFileName(dlg.FileName));
//                    }
//                }
//            }
//            catch (Exception ex)
//            {
//                NotifyClient("Open failed: " + ex.Message);
//            }
//        }


//        public void RunSelectedAction(string key)
//        {
//            try
//            {
//                switch (key)
//                {
//                    case "punctuation":
//                        Action_PunctuationShift();   // ✅ yeh line add
//                        break;

//                    case "query":
//                        NotifyClient("Query inserted (demo).");
//                        break;

//                    case "doi":
//                        NotifyClient("DOI validation (demo).");
//                        break;

//                    case "url":
//                        NotifyClient("URL validation (demo).");
//                        break;

//                    case "preedit":
//                        NotifyClient("PreEditing (demo).");
//                        break;

//                    default:
//                        NotifyClient("Action not implemented: " + key);
//                        break;
//                }
//            }
//            catch (Exception ex)
//            {
//                NotifyClient("Action failed: " + ex.Message);
//            }
//        }



//        public void InsertTextIntoWord(string text)
//        {
//            try
//            {
//                var sel = this.Application.Selection;
//                sel.TypeText(text ?? "");
//            }
//            catch (System.Exception ex)
//            {
//                System.Windows.Forms.MessageBox.Show("Insert failed: " + ex.Message);
//            }
//        }

//        public void NotifyClient(string msg)
//        {
//            _control?.SendToClient(new { type = "notify", text = msg });
//        }

//        private void ThisAddIn_Shutdown(object sender, System.EventArgs e) { }

//        private void InternalStartup()
//        {
//            this.Startup += new System.EventHandler(ThisAddIn_Startup);
//            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
//        }
//        private void Action_PunctuationShift()
//        {
//            var app = this.Application;
//            Word.Document doc = null;

//            try
//            {
//                doc = app.ActiveDocument;
//                if (doc == null)
//                {
//                    NotifyClient("No active document.");
//                    return;
//                }

//                // (optional) UI smoother
//                bool oldScreenUpdating = app.ScreenUpdating;
//                app.ScreenUpdating = false;

//                int moveCount = 0, skippedSuper = 0, nonNumericSkipped = 0, spaceFixCount = 0;
//                var modifiedBrackets = new List<int>();

//                foreach (Word.Hyperlink link in doc.Hyperlinks)
//                {
//                    try
//                    {
//                        string sub = link.SubAddress;
//                        if (string.IsNullOrEmpty(sub) ||
//                            !sub.StartsWith("bib", StringComparison.OrdinalIgnoreCase))
//                            continue;

//                        Word.Range range = link.Range;
//                        string linkText = range.Text.Trim();

//                        // ---- FILTERS (ONLY_HYPERLINKS) ----
//                        // Must begin with a digit (optionally [ or ()
//                        if (!Regex.IsMatch(linkText, @"^\s*[\[\(]?\s*\d"))
//                        { nonNumericSkipped++; continue; }

//                        // Skip if contains any letter
//                        if (Regex.IsMatch(linkText, @"[A-Za-z]"))
//                        { nonNumericSkipped++; continue; }

//                        // Skip if looks like a year (1900–2099)
//                        if (Regex.IsMatch(linkText, @"\b(19|20)\d{2}\b"))
//                        { nonNumericSkipped++; continue; }

//                        // Skip superscript hyperlinks
//                        bool isSuper = false;
//                        for (int i = range.Start; i < range.End; i++)
//                        {
//                            Word.Range ch = doc.Range(i, i + 1);
//                            if (ch.Font.Superscript == -1) { isSuper = true; }
//                            Marshal.ReleaseComObject(ch);
//                            if (isSuper) break;
//                        }
//                        if (isSuper) { skippedSuper++; continue; }

//                        // ---- FIND OPENING BRACKET BEFORE CITATION ----
//                        int bracketPos = -1;
//                        for (int i = range.Start - 1; i >= Math.Max(0, range.Start - 200); i--)
//                        {
//                            string t = doc.Range(i, i + 1).Text;
//                            if (t == "(" || t == "[") { bracketPos = i; break; }
//                            if (t == "\r" || t == "\n") break;
//                        }
//                        if (bracketPos == -1) continue;

//                        // ---- FIND PUNCTUATION JUST BEFORE BRACKET ----
//                        int punctPos = -1;
//                        char[] marks = { '.', ',', ';', ':', '!', '?' };
//                        for (int j = bracketPos - 1; j >= Math.Max(0, bracketPos - 10); j--)
//                        {
//                            string c = doc.Range(j, j + 1).Text;
//                            if (string.IsNullOrWhiteSpace(c)) continue;
//                            if (Array.Exists(marks, p => p.ToString() == c)) { punctPos = j; break; }
//                            else break;
//                        }
//                        if (punctPos == -1) continue;

//                        string punc = doc.Range(punctPos, punctPos + 1).Text;
//                        doc.Range(punctPos, punctPos + 1).Delete();

//                        // ---- FIND CLOSING BRACKET AFTER CITATION ----
//                        int closePos = -1;
//                        for (int k = range.End; k < Math.Min(doc.Content.End, range.End + 200); k++)
//                        {
//                            string c = doc.Range(k, k + 1).Text;
//                            if (c == ")" || c == "]") { closePos = k; break; }
//                            if (c == "\r" || c == "\n") break;
//                        }
//                        if (closePos == -1)
//                        {
//                            for (int k = range.Start; k <= Math.Min(range.End + 3, doc.Content.End); k++)
//                            {
//                                string c = doc.Range(k, k + 1).Text;
//                                if (c == ")" || c == "]") { closePos = k; break; }
//                            }
//                        }
//                        if (closePos == -1) continue;

//                        // ---- MOVE PUNCTUATION OUTSIDE CLOSING BRACKET ----
//                        doc.Range(closePos + 1, closePos + 1).InsertAfter(punc);
//                        moveCount++;
//                        modifiedBrackets.Add(bracketPos);
//                    }
//                    catch
//                    {
//                        // skip any single hyperlink error
//                    }
//                }

//                // ---- SPACE FIXES around modified bracket areas ----
//                foreach (int bracketPos in modifiedBrackets)
//                {
//                    int from = Math.Max(0, bracketPos - 40);
//                    Word.Range local = doc.Range(from, bracketPos + 1);
//                    Word.Find f = local.Find;
//                    f.ClearFormatting();
//                    f.Replacement.ClearFormatting();

//                    // collapse multiple spaces before opening bracket to single space
//                    f.Text = "[ ^s^t^32^160]{2,}([\\(\\[])";
//                    f.Replacement.Text = " \\1";
//                    f.MatchWildcards = true;
//                    object replaceAll = Word.WdReplace.wdReplaceAll;
//                    f.Execute(Replace: ref replaceAll);
//                    spaceFixCount++;
//                }

//                // ===== SAFE SAVE BLOCK =====
//                try
//                {
//                    // Skip save if document is new (like "Document1") or ReadOnly
//                    if (!doc.ReadOnly && !string.IsNullOrEmpty(doc.FullName) && !doc.FullName.Contains("Document"))
//                    {
//                        var oldAlerts = app.DisplayAlerts;
//                        app.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;

//                        try { doc.Save(); } catch { }

//                        app.DisplayAlerts = oldAlerts;
//                    }
//                }
//                catch { }


//                // Per your choice "C": no detailed counts; just a short toast
//                NotifyClient("Punctuation shift done.");

//                // restore UI
//                app.ScreenUpdating = true;
//            }
//            catch (Exception ex)
//            {
//                NotifyClient("Punctuation shift failed: " + ex.Message);
//            }
//        }

//    }
//}
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using System;
using System.IO;
using System.Threading.Tasks;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace PowerEditAddIn
{
    public partial class ThisAddIn
    {
        private CustomTaskPane _pane;
        private DesignerPane _control;

        // ✅ NEW: right-side utility pane support
        private CustomTaskPane _utilityPane;
        private UtilityPane _utilityControl;

        private async void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            _control = new DesignerPane();
            _pane = this.CustomTaskPanes.Add(_control, "PowerEdit");
            _pane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionLeft;
            _pane.Width = 420;
            _pane.Visible = false;

            var dataFolder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "PowerEditAddin", "WebView2Data");
            Directory.CreateDirectory(dataFolder);

            var env = await CoreWebView2Environment.CreateAsync(userDataFolder: dataFolder);
            await _control.webView21.EnsureCoreWebView2Async(env);

            string basePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ui");
            _control.webView21.CoreWebView2.SetVirtualHostNameToFolderMapping(
                "app.local", basePath,
                CoreWebView2HostResourceAccessKind.Allow);

            _control.webView21.CoreWebView2.Navigate("https://app.local/index.html");
            _control.webView21.CoreWebView2.WebMessageReceived += (s, args) =>
            {
                var doc = System.Text.Json.JsonDocument.Parse(args.WebMessageAsJson);
                var root = doc.RootElement;
                var type = root.GetProperty("type").GetString();

                if (type == "browse")
                    BrowseAndOpenDocx();

                else if (type == "runAction")
                {
                    var action = root.GetProperty("action").GetString();
                    RunSelectedAction(action);
                }
            };
        }

        public void TogglePane()
        {
            if (_pane == null)
            {
                System.Windows.Forms.MessageBox.Show("Pane is NULL!");
                return;
            }
            _pane.Visible = !_pane.Visible;
        }

        // ✅ NEW: Called from Ribbon “Open” button → opens both panes
        public void OpenBothPanes()
        {
            try
            {
                if (_pane != null)
                    _pane.Visible = true;

                if (_utilityPane == null)
                    CreateUtilityPane();

                _utilityPane.Visible = true;
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error opening panes: " + ex.Message);
            }
        }

        // ✅ NEW: creates right-side collapsible utility pane
        private void CreateUtilityPane()
        {
            if (_utilityPane != null) return;

            _utilityControl = new UtilityPane();
            _utilityPane = this.CustomTaskPanes.Add(_utilityControl, "PowerEdit Utility");
            _utilityPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
            _utilityPane.Width = _utilityControl.DesiredWidth;
            _utilityPane.Visible = false;

            _utilityControl.ExpandedChanged += (s, expanded) =>
            {
                _utilityPane.Width = _utilityControl.DesiredWidth;
            };
        }

        public void BrowseAndOpenDocx()
        {
            try
            {
                using (var dlg = new System.Windows.Forms.OpenFileDialog())
                {
                    dlg.Filter = "Word Documents|*.docx";
                    dlg.Title = "Select a Word Document";

                    if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        var app = this.Application;

                        try
                        {
                            if (app.Documents.Count > 0)
                            {
                                app.ActiveDocument.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                            }
                        }
                        catch { }

                        var doc = app.Documents.Open(dlg.FileName, ReadOnly: false, Visible: true);
                        _pane.Visible = true;
                        NotifyClient("Opened: " + Path.GetFileName(dlg.FileName));
                    }
                }
            }
            catch (Exception ex)
            {
                NotifyClient("Open failed: " + ex.Message);
            }
        }

        public void RunSelectedAction(string key)
        {
            try
            {
                switch (key)
                {
                    case "punctuation":
                        Action_PunctuationShift();
                        break;
                    case "query":
                        NotifyClient("Query inserted (demo).");
                        break;
                    case "doi":
                        NotifyClient("DOI validation (demo).");
                        break;
                    case "url":
                        NotifyClient("URL validation (demo).");
                        break;
                    case "preedit":
                        NotifyClient("PreEditing (demo).");
                        break;
                    default:
                        NotifyClient("Action not implemented: " + key);
                        break;
                }
            }
            catch (Exception ex)
            {
                NotifyClient("Action failed: " + ex.Message);
            }
        }

        public void InsertTextIntoWord(string text)
        {
            try
            {
                var sel = this.Application.Selection;
                sel.TypeText(text ?? "");
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Insert failed: " + ex.Message);
            }
        }

        public void NotifyClient(string msg)
        {
            _control?.SendToClient(new { type = "notify", text = msg });
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e) { }

        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        private void Action_PunctuationShift()
        {
            var app = this.Application;
            Word.Document doc = null;

            try
            {
                doc = app.ActiveDocument;
                if (doc == null)
                {
                    NotifyClient("No active document.");
                    return;
                }

                bool oldScreenUpdating = app.ScreenUpdating;
                app.ScreenUpdating = false;

                int moveCount = 0, skippedSuper = 0, nonNumericSkipped = 0, spaceFixCount = 0;
                var modifiedBrackets = new List<int>();

                foreach (Word.Hyperlink link in doc.Hyperlinks)
                {
                    try
                    {
                        string sub = link.SubAddress;
                        if (string.IsNullOrEmpty(sub) || !sub.StartsWith("bib", StringComparison.OrdinalIgnoreCase))
                            continue;

                        Word.Range range = link.Range;
                        string linkText = range.Text.Trim();

                        if (!Regex.IsMatch(linkText, @"^\s*[\[\(]?\s*\d"))
                        { nonNumericSkipped++; continue; }

                        if (Regex.IsMatch(linkText, @"[A-Za-z]"))
                        { nonNumericSkipped++; continue; }

                        if (Regex.IsMatch(linkText, @"\b(19|20)\d{2}\b"))
                        { nonNumericSkipped++; continue; }

                        bool isSuper = false;
                        for (int i = range.Start; i < range.End; i++)
                        {
                            Word.Range ch = doc.Range(i, i + 1);
                            if (ch.Font.Superscript == -1) { isSuper = true; }
                            Marshal.ReleaseComObject(ch);
                            if (isSuper) break;
                        }
                        if (isSuper) { skippedSuper++; continue; }

                        int bracketPos = -1;
                        for (int i = range.Start - 1; i >= Math.Max(0, range.Start - 200); i--)
                        {
                            string t = doc.Range(i, i + 1).Text;
                            if (t == "(" || t == "[") { bracketPos = i; break; }
                            if (t == "\r" || t == "\n") break;
                        }
                        if (bracketPos == -1) continue;

                        int punctPos = -1;
                        char[] marks = { '.', ',', ';', ':', '!', '?' };
                        for (int j = bracketPos - 1; j >= Math.Max(0, bracketPos - 10); j--)
                        {
                            string c = doc.Range(j, j + 1).Text;
                            if (string.IsNullOrWhiteSpace(c)) continue;
                            if (Array.Exists(marks, p => p.ToString() == c)) { punctPos = j; break; }
                            else break;
                        }
                        if (punctPos == -1) continue;

                        string punc = doc.Range(punctPos, punctPos + 1).Text;
                        doc.Range(punctPos, punctPos + 1).Delete();

                        int closePos = -1;
                        for (int k = range.End; k < Math.Min(doc.Content.End, range.End + 200); k++)
                        {
                            string c = doc.Range(k, k + 1).Text;
                            if (c == ")" || c == "]") { closePos = k; break; }
                            if (c == "\r" || c == "\n") break;
                        }
                        if (closePos == -1)
                        {
                            for (int k = range.Start; k <= Math.Min(range.End + 3, doc.Content.End); k++)
                            {
                                string c = doc.Range(k, k + 1).Text;
                                if (c == ")" || c == "]") { closePos = k; break; }
                            }
                        }
                        if (closePos == -1) continue;

                        doc.Range(closePos + 1, closePos + 1).InsertAfter(punc);
                        moveCount++;
                        modifiedBrackets.Add(bracketPos);
                    }
                    catch { }
                }

                foreach (int bracketPos in modifiedBrackets)
                {
                    int from = Math.Max(0, bracketPos - 40);
                    Word.Range local = doc.Range(from, bracketPos + 1);
                    Word.Find f = local.Find;
                    f.ClearFormatting();
                    f.Replacement.ClearFormatting();
                    f.Text = "[ ^s^t^32^160]{2,}([\\(\\[])";
                    f.Replacement.Text = " \\1";
                    f.MatchWildcards = true;
                    object replaceAll = Word.WdReplace.wdReplaceAll;
                    f.Execute(Replace: ref replaceAll);
                    spaceFixCount++;
                }

                try
                {
                    if (!doc.ReadOnly && !string.IsNullOrEmpty(doc.FullName) && !doc.FullName.Contains("Document"))
                    {
                        var oldAlerts = app.DisplayAlerts;
                        app.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
                        try { doc.Save(); } catch { }
                        app.DisplayAlerts = oldAlerts;
                    }
                }
                catch { }

                NotifyClient("Punctuation shift done.");
                app.ScreenUpdating = true;
            }
            catch (Exception ex)
            {
                NotifyClient("Punctuation shift failed: " + ex.Message);
            }
        }
    }
}
