using OpenXml.IRepository;
using OpenXml.Service;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXml.Repository
{
    public class FuntionCustom : IFuntionCustom
    {
        public List<T> DocDuLieuExcel<T>(string fileURRL)
        {
            ServiceDocDuLieuExcel.XuLy(fileURRL);
            List<T> data = default(List<T>);
            return data;
        }
        public void ThayTheBang<T>(T input)
        {
            ServiceThayTheBang.XuLy();
        }
        public void ThayTheThamSo<T>(T input)
        {
            ServiceThayTheThamSo.XuLy();    
        }
    }
}
