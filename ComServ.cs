using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace COMServ
{  
    [Guid("2F4EFFE4-5C6C-4131-A231-CFD3290B29F4")]
    [ComVisible(true)] 
    public interface IFunction
    {
        string FuncName();
        double TheFunc(double x);
        
    }
    [ProgId("ComServ")]
    [Guid("C4949B92-371D-4F0A-8DB2-F44C3084C920")]
    [ComVisible(true)]
    public class MyClass:IFunction
    {
        public string FuncName()
        {
            return "cos(x)";
        }
        public double TheFunc(double x)
        {
            return Math.Cos(x);
        }
    }
}
