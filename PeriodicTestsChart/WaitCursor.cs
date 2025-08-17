using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace PeriodicTestsChartReport
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
