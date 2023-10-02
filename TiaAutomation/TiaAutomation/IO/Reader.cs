using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using TiaAutomation.IO.Contracts;
using TiaAutomation.Utilities.Messages;

namespace TiaAutomation.IO
{
    public class Reader : IReader
    {
        public string ReadLine(TextBox textBoxObject)
        {
            return textBoxObject.Text;
        }
    }
}
