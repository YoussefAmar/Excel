using System;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using E = Microsoft.Office.Interop.Excel;
using System.IO;

namespace Excel
{
    class Program
    {
        static void Main()
        {
            int i = 1, j = 0;

            var xlApp = new Microsoft.Office.Interop.Excel.Application();
            E.Workbook xlWbook;
            E.Worksheet xlSheet;
            object misValue = System.Reflection.Missing.Value;

            string path = System.Reflection.Assembly.GetExecutingAssembly().Location;
            path = path.Substring(0,path.LastIndexOf("\\", StringComparison.Ordinal)) + "\\";
            string pathtxt = path  + "Prix.txt";
            string pathExcel = path + "Excel.xls";

            if (xlApp == null)
            {
                Console.WriteLine("Excel n'est pas installé");
                return;
            }

            xlWbook = xlApp.Workbooks.Add(misValue);
            xlSheet = (E.Worksheet)xlWbook.Worksheets.Item[1];

            var lignes = File.ReadAllLines(pathtxt);

            Console.WriteLine();

            Console.WriteLine("Liste des éléments dans le fichiers :");
            Console.WriteLine(pathtxt);
            Console.WriteLine();

            foreach (var l in lignes)
            {
                var data = l.Split(':').ToList();

                foreach (var s in data)
                {
                    j++;

                    Console.WriteLine(s);

                    xlSheet.Cells[i, j] = s;

                    if (j == 2)
                    {
                        i++;
                        j = 0;
                    }

                }
            }

            var moyenne = new StringBuilder("=MOYENNE(");
            moyenne.Append("B" + 1);
            moyenne.Append(":");
            moyenne.Append("B" + (i-1));
            moyenne.Append(")");

            xlSheet.Cells[i, 2] = moyenne.ToString();

            try
            {
                xlWbook.SaveAs(pathExcel, E.XlFileFormat.xlWorkbookNormal, misValue, misValue, true, misValue, E.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                Console.WriteLine("Dossier de réception : " + pathExcel);
                Process.Start("explorer.exe",pathExcel);
                Process.Start("explorer.exe", pathtxt);
            }
            catch
            {
                Console.WriteLine();
                Console.WriteLine("Refus de remplacement du fichier");
            }

            xlWbook.Close(true, misValue, misValue);

            Marshal.ReleaseComObject(xlSheet);
            Marshal.ReleaseComObject(xlWbook);
            Marshal.ReleaseComObject(xlApp);

            Console.WriteLine("Appuyez sur Enter pour fermer le programme");
            Console.ReadLine();

            return;
        }
    }
}