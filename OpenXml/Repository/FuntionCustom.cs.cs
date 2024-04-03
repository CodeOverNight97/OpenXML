using OpenXml.IRepository;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXml.Repository
{
    public class FuntionCustom : IFuntionCustom
    {
        T IFuntionCustom.DocDuLieuExcel<T>(T input)
        {
            throw new NotImplementedException();
        }

        void IFuntionCustom.ThayTheBang<T>(T input)
        {
            throw new NotImplementedException();
        }

        void IFuntionCustom.ThayTheThamSo<T>(T input)
        {
            throw new NotImplementedException();
        }
    }
}
