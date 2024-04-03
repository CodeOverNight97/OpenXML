using OpenXml.IRepository;

namespace OpenXmlCustom
{
    public class FuntionOpenXmlCustom
    {
        private readonly IFuntionCustom _funtionCus;
        public FuntionOpenXmlCustom(IFuntionCustom FuntionCus)
        {
            _funtionCus = FuntionCus;
        }
        public void a()
        {
            Console.WriteLine("ádafd");
        }
    }
}