using Microsoft.Extensions.DependencyInjection;
using OpenXml.IRepository;
using OpenXml.Model;
using OpenXml.Repository;

namespace OpenXml
{
    public class Program
    {
        public static void Main(String[] args)
        {
            //var a  = int.Parse(Console.ReadLine());
            IFuntionCustom Fun = new FuntionCustom();
            var a = Fun.DocDuLieuExcel<TestDocDuLieuExcelmodel>("C:/LamViec/OpenXml/OpenXml/File/2.xlsx");
            Console.WriteLine(a);
        }
    }
}
