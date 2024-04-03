using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXml.IRepository
{
    public interface IFuntionCustom
    {
        void ThayTheThamSo<T>(T input);
        void ThayTheBang<T>(T input);
        List<T> DocDuLieuExcel<T>(string fileURRL);
    }
}
