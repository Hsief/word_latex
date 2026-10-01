using System;
using System.Windows.Forms;
using AfterMathCore;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace WordLatexAddin
{
    /// <summary>VSTO application add-in entry point and automatic-conversion event coordinator.</summary>
    public partial class ThisAddIn
    {
        private Timer _automaticTimer;
        private AddInAutomation _automationObject;
        private Ribbon _ribbonController;

        public Settings PluginSettings { get; private set; }
        public LatexConverter Converter { get; private set; }
        public EquationNumbering Numbering { get; private set; }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            _ribbonController = new Ribbon(this);
            return _ribbonController;
        }

        internal bool RibbonCustomUiRequested
        {
            get { return _ribbonController != null && _ribbonController.CustomUiRequested; }
        }

        internal bool RibbonLoaded
        {
            get { return _ribbonController != null && _ribbonController.IsLoaded; }
        }

        internal string RequestedRibbonId
        {
            get { return _ribbonController == null ? string.Empty : _ribbonController.RequestedRibbonId; }
        }

        internal bool ActivateRibbonTab()
        {
            return _ribbonController != null && _ribbonController.ActivateMainTab();
        }

        internal string RibbonActivationError
        {
            get { return _ribbonController == null ? "Ribbon 控制器尚未创建。" : _ribbonController.ActivationError; }
        }

        /// <summary>Exposes a small COM-visible surface used by the installer smoke test.</summary>
        protected override object RequestComAddInAutomationService()
        {
            if (_automationObject == null) _automationObject = new AddInAutomation(this);
            return _automationObject;
        }

        internal ConversionSummary ConvertCurrentForAutomation()
        {
            if (Converter == null) throw new InvalidOperationException("插件尚未完成初始化。");
            return Converter.ConvertCurrent(Application.Selection);
        }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            PluginSettings = Settings.Load();
            LatexScanner scanner = new LatexScanner();
            Converter = new LatexConverter(Application, scanner, new OmmlBuilder(new OmmlGenerator()));
            Numbering = new EquationNumbering(PluginSettings);

            _automaticTimer = new Timer();
            _automaticTimer.Tick += AutomaticTimerTick;
            RefreshAutomaticMode();
            Application.WindowSelectionChange += ApplicationWindowSelectionChange;
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            if (_automaticTimer != null)
            {
                _automaticTimer.Stop();
                _automaticTimer.Tick -= AutomaticTimerTick;
                _automaticTimer.Dispose();
            }
            Application.WindowSelectionChange -= ApplicationWindowSelectionChange;
        }

        public void RefreshAutomaticMode()
        {
            if (_automaticTimer == null) return;
            _automaticTimer.Stop();
            _automaticTimer.Interval = Math.Max(150, Math.Min(2000, PluginSettings.AutoDelayMilliseconds));
        }

        private void ApplicationWindowSelectionChange(Word.Selection selection)
        {
            if (!PluginSettings.AutomaticConversion || Converter == null || Converter.IsBusy) return;
            _automaticTimer.Stop();
            _automaticTimer.Start();
        }

        private void AutomaticTimerTick(object sender, EventArgs e)
        {
            _automaticTimer.Stop();
            if (!PluginSettings.AutomaticConversion || Converter == null || Converter.IsBusy) return;
            try { Converter.TryAutoConvert(Application.Selection); }
            catch { }
        }

        #region VSTO generated code

        /// <summary>Required by the VSTO runtime.</summary>
        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }

        #endregion
    }
}
