using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;

namespace ConsoleApp2
{



    class Program
    {
        public static void RemoveCitations(string f_name)
        {
            _Application word_app = new Application();

            word_app.Visible = false;

            object missing = Type.Missing;
            object filename = f_name;
            object confirm_conversions = false;
            object read_only = false;
            object add_to_recent_files = false;
            object format = 0;
            _Document word_doc =
                word_app.Documents.Open(ref filename,
                    ref confirm_conversions,
                    ref read_only, ref add_to_recent_files,
                    ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref format, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing);

            object index = 1;
            while (word_doc.Hyperlinks.Count > 0)
            {
                word_doc.Hyperlinks.get_Item(ref index).Delete();
            }

            object save_changes = true;
            word_doc.Close(ref save_changes, ref missing, ref missing);

            word_app.Quit(ref save_changes, ref missing, ref missing);
            Console.WriteLine("Done");

        }

        static void Main(string[] args)
        {
            string article_type = "xyz";

            string[] articletypes = { "erratum", "xyz", "pxz","xyz" };
            var results = Array.FindAll(articletypes, s => s.Equals(article_type));
            
            if(results.Length > 0)
            {
                RemoveCitations(@"D:\advs201901198_edited\advs201901198.docx");
            }
        }
    }
}
