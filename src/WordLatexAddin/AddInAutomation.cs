using System;
using System.Runtime.InteropServices;

namespace WordLatexAddin
{
    /// <summary>COM automation API for post-install verification and diagnostics.</summary>
    [ComVisible(true)]
    [Guid("01C4F2B1-F5CC-46A5-81A0-D1F01829EE66")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public sealed class AddInAutomation
    {
        private readonly ThisAddIn _addIn;

        internal AddInAutomation(ThisAddIn addIn)
        {
            _addIn = addIn ?? throw new ArgumentNullException("addIn");
        }

        /// <summary>Converts the active Word selection and returns the native-equation count created.</summary>
        public int ConvertCurrentSelection()
        {
            ConversionSummary summary = _addIn.ConvertCurrentForAutomation();
            if (summary.Converted == 0 && summary.Errors.Count > 0)
                throw new InvalidOperationException(summary.ToUserMessage());
            return summary.Converted;
        }
    }
}
