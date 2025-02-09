using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SwToolkit
{
    internal class Util

    {

        public static ISldWorks ConnectToSolidWorks()
        {
            try
            {
                // 获取正在运行的SolidWorks实例
                ISldWorks swApp = (ISldWorks)Marshal.GetActiveObject("SldWorks.Application");
                if (swApp != null)
                {
                    Console.WriteLine("成功连接到SolidWorks!");
                    return swApp;
                }
                else
                {
                    Console.WriteLine("未找到正在运行的SolidWorks实例。");
                    return null;
                }
            }
            catch (COMException ex)
            {
                Console.WriteLine("连接失败: " + ex.Message);
                return null;
            }
        }
    }
}
