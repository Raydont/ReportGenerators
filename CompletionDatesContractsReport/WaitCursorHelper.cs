using System;
using System.Windows.Forms;

namespace CompletionDatesContractsReport
{
    public class WaitCursorHelper : IDisposable
    {
        private bool _waitCursor;

        public WaitCursorHelper(bool useWaitCursor)
        {
            _waitCursor = Application.UseWaitCursor;
            Application.UseWaitCursor = useWaitCursor;
        }

        public void Dispose()
        {
            Application.UseWaitCursor = _waitCursor;
        }
    }

}
