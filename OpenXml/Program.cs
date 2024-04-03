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
            switch (a)
            {
                case 1:
                    Fun.DocDuLieuExcel<dynamic>("");
                    break;
                case 2:
                    Fun.ThayTheThamSo(1);
                    break;
                case 3:
                    Fun.ThayTheBang(1);
                    break;
                default:
                    Console.WriteLine("Nhập 1 2 3 thôi ngu nó vừa");
                    break;
            }
        }
    }
}
