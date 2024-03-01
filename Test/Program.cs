using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Word;
using System.IO;

namespace Test
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Вы хотите создать чек?");
            String Otvet = Console.ReadLine();
            if (Otvet == "Да") 
            {
                Application application = new Application();
                var document = application.Documents.Add();
                document.Content.Text = "TEST33";
                document.SaveAs2(Path.GetFullPath($@"..\..\..\Чеки\{DateTime.Now.ToString("yyyyMMdd_HH mm ss")}.docx"));
                document.Close();
                application.Quit();
            }
            else if( Otvet == "Нет")
            {
                Console.WriteLine("Ладно.");
            }
            Console.ReadKey();
        }
    }
}
