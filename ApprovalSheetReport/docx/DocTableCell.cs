using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReportHelpers
{
    public class DocTableCell
    {
        public string content = "";

        public int width = 1;
        public int height = 1;
        public Action<object>[] applyStyles = new Action<object>[0];
    }

    public class DocTableCell2
    {
        public string content = "";

        public int width = 1;
        public int height = 1;
        public int positionX = 0;
        public int positionY = 0;
        public List<Action<object>> applyStyles = new List<Action<object>>();
    }
}
