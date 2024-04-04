using OpenXml.IRepository;
using OpenXml.Model;
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
        public DocDuLieuExcel<T> DocDuLieuExcel<T>(string fileURRL)
        {
            
            var data = ServiceDocDuLieuExcel.XuLy<T>(fileURRL);
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
