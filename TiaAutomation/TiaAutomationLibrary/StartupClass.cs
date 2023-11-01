using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Emit;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Threading.Tasks;
using TiaAutomation.Core;
using TiaAutomation.Core.Contracts;
using TiaAutomation.IO.Contracts;

namespace TiaAutomationLibrary
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class StartupClass
    {
        private IEngine engine;
        private IWriter writer;
        private IReader reader;
        private string _resultStringFromEngine = string.Empty;
        private string _inputFromTextBox = string.Empty;
        private string _initialTextForStatusLabel = "Status: OK";

        public StartupClass()
        {
            // Initialize the backend engine
            engine = new Engine();

            // retrieve the instances from the engine
            this.writer = engine.GetWriter();
            this.reader = engine.GetReader();

            //initial text for the label
            //label1.Text = _initialTextForStatusLabel;
        }

        // Start
        //public string button1_Click()
        public string StartTia()
        {
            _resultStringFromEngine = this.engine.Start();
            //writer.Write(label1, _resultStringFromEngine);
            return _resultStringFromEngine;
        }

        // Build Solution
        //public string button2_Click()
        public string Compile()
        {
            _resultStringFromEngine = this.engine.BuildSolution();
            //writer.Write(label1, _resultStringFromEngine);
            return _resultStringFromEngine;
        }

        // Exit
        //public string button9_Click()
        public string Close()
        {
            _resultStringFromEngine = this.engine.Exit();
           //writer.Write(label1, _resultStringFromEngine);
           return _resultStringFromEngine;
        }
    }
}
