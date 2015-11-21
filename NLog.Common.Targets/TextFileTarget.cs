using System;
using System.Collections.Generic;
using System.Configuration;
using NLog.Config;
using NLog.Targets;
using System.Threading;

namespace NLog.TextFile
{
	[Target("TextFile")]
	public sealed class TextFileTarget : Target
	{
		protected override void Write(LogEventInfo logEvent)
		{
            WriteToText();
        }

        private static void WriteToText()
        {
            // Compose a string that consists of three lines.
            string lines = "First line.\r\nSecond line.\r\nThird line.";

            // Write the string to a file.
            System.IO.StreamWriter file = new System.IO.StreamWriter("c:\\test.txt");
            file.WriteLine(lines);

            file.Close();
        }
    }
}