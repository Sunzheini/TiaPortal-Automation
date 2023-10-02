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
    public class Writer : IWriter
    {
        public void Write(TextBlock labelObject, string message)
        {
            labelObject.Text = message;
        }

        public void WriteLine(TextBlock labelObject, string message)
        {
            labelObject.Text += message + Environment.NewLine;
        }
    }
}
