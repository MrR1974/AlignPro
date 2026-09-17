using Office = Microsoft.Office.Core;

namespace AlignPro.AddIn
{
    public partial class ThisAddIn
    {
        private AlignProController? _controller;

        /// <summary>
        /// Created on first use rather than at startup, because
        /// <see cref="CreateRibbonExtensibilityObject"/> runs before the Startup event and we would
        /// rather not depend on exactly when <c>Application</c> becomes available.
        /// </summary>
        internal AlignProController Controller =>
            _controller ?? (_controller = new AlignProController(Application));

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject() =>
            new AlignProRibbon(() => Controller);

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // PowerPoint does not reliably raise Shutdown for COM add-ins, so nothing important
            // should depend on this running.
            _controller = null;
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
