using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

namespace WordLatexAddin
{
    /// <summary>Ribbon XML callback surface for the “论文工具” tab.</summary>
    [ComVisible(true)]
    public sealed class Ribbon : Office.IRibbonExtensibility
    {
        private readonly ThisAddIn _addIn;
        private Office.IRibbonUI _ribbon;

        internal bool CustomUiRequested { get; private set; }
        internal bool IsLoaded { get; private set; }
        internal string RequestedRibbonId { get; private set; }

        public Ribbon(ThisAddIn addIn)
        {
            _addIn = addIn;
        }

        public string GetCustomUI(string ribbonId)
        {
            CustomUiRequested = true;
            RequestedRibbonId = ribbonId ?? string.Empty;
            Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("WordLatexAddin.Ribbon.xml");
            if (stream == null) throw new InvalidOperationException("无法读取内嵌的 Ribbon.xml。 ");
            using (stream)
            using (StreamReader reader = new StreamReader(stream))
            {
                return reader.ReadToEnd();
            }
        }

        public void OnLoad(Office.IRibbonUI ribbonUi)
        {
            _ribbon = ribbonUi;
            IsLoaded = true;
        }

        public void ConvertCurrent(Office.IRibbonControl control)
        {
            RunConversion(delegate { return _addIn.Converter.ConvertCurrent(_addIn.Application.Selection); });
        }

        public void ConvertParagraph(Office.IRibbonControl control)
        {
            RunConversion(delegate { return _addIn.Converter.ConvertParagraph(_addIn.Application.Selection); });
        }

        public void ConvertDocument(Office.IRibbonControl control)
        {
            RunConversion(delegate { return _addIn.Converter.ConvertDocument(_addIn.Application.ActiveDocument); });
        }

        public bool GetAutomaticMode(Office.IRibbonControl control)
        {
            return _addIn.PluginSettings.AutomaticConversion;
        }

        public void SetAutomaticMode(Office.IRibbonControl control, bool pressed)
        {
            _addIn.PluginSettings.AutomaticConversion = pressed;
            _addIn.PluginSettings.Save();
            _addIn.RefreshAutomaticMode();
        }

        public void NumberEquation(Office.IRibbonControl control)
        {
            RunAction(delegate { _addIn.Numbering.AddNumber(_addIn.Application.Selection); });
        }

        public void ReferenceEquation(Office.IRibbonControl control)
        {
            RunAction(delegate { _addIn.Numbering.InsertReference(_addIn.Application.Selection); });
        }

        public void OpenSettings(Office.IRibbonControl control)
        {
            if (_addIn.PluginSettings.ShowDialog())
            {
                _addIn.RefreshAutomaticMode();
                if (_ribbon != null) _ribbon.Invalidate();
            }
        }

        private void RunConversion(Func<ConversionSummary> work)
        {
            try
            {
                ConversionSummary summary = work();
                if (_addIn.PluginSettings.ShowCompletionMessages || summary.Errors.Count > 0)
                {
                    MessageBox.Show(summary.Found == 0 ? "没有找到可转换的 LaTeX 公式。" : summary.ToUserMessage(),
                        "WordLatexVSTO", MessageBoxButtons.OK,
                        summary.Errors.Count == 0 ? MessageBoxIcon.Information : MessageBoxIcon.Warning);
                }
            }
            catch (Exception exception)
            {
                ShowError(exception);
            }
        }

        private static void RunAction(Action work)
        {
            try { work(); }
            catch (Exception exception) { ShowError(exception); }
        }

        private static void ShowError(Exception exception)
        {
            MessageBox.Show(exception.Message, "WordLatexVSTO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }
}
