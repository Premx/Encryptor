using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace outlookaddin
{
    public class Preview
    {
        public Preview(String text)
        {
            string tempfile = Path.Combine(Path.GetTempPath(), "Encryptor_Preview.html");
            using (StreamWriter sw = new StreamWriter(tempfile))
            {
                sw.Write(text);
            }

            System.Diagnostics.Process.Start(tempfile);
        }
    }
}
