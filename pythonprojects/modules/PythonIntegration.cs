using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using IronPython.Hosting;

namespace pythonintegration
{
    class Program
    {
        static void Main(string[] args)
        {
            var docx_file = @"D:\test\tttt\sishir\1\adfm201909391_identical_Partrefs.docx";
            var engine = Python.CreateEngine();
            var searchPaths = engine.GetSearchPaths();
            searchPaths.Add(@"C:\Python27\Lib");
            

            engine.SetSearchPaths(searchPaths);
            var script = @"D:\shishir\aptproject\pythonprojects\modules\ref_sequence.py";
 
            var source = engine.CreateScriptSourceFromFile(script);
            
            var argv = new List<string>();
            argv.Add("");
            argv.Add(docx_file);
            engine.GetSysModule().SetVariable("argv", argv);
            var elO = engine.Runtime.IO;
            var errors = new MemoryStream();
            elO.SetErrorOutput(errors, Encoding.Default);
            var results = new MemoryStream();
            elO.SetOutput(results, Encoding.Default);
            var scope = engine.CreateScope();
            try
            {
                source.Execute(scope);
                string str(byte[] x) => Encoding.Default.GetString(x);
                Console.WriteLine(str(results.ToArray()));
            }
            catch(Exception ex)
            {
                Console.WriteLine("ERROR: "+ex.Message);
            }
            
           
        }
    }
}
