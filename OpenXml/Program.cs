using Microsoft.Extensions.DependencyInjection;
using OpenXml.IRepository;
using OpenXml.Repository;

namespace OpenXml
{
    public class Program
    {
        public static void Main(String[] args)
        {
            var a  = int.Parse(Console.ReadLine());
            IFuntionCustom Fun = new FuntionCustom();
            Fun.DocDuLieuExcel<dynamic>("C:/LamViec/OpenXml/OpenXml/File/test.xlsx");
        }
    }
}
