using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;

namespace SwToolkit
{
    public partial class Form1 : Form
    {
        private SldWorks swApp;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
  
        private void button1_Click(object sender, EventArgs e)
        {

            MessageBox.Show("按钮被点击了！");
            // 尝试获取正在运行的SolidWorks实例
            //swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
            this.swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
            if (swApp == null)
            {
                Debug.Print("Failed to connect");
                return;
                //swApp.SendMsgToUser(msg);
            }

            ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;

            ModelDocExtension swModelDocExt = (ModelDocExtension)swModel.Extension;

            Measure swMeasure = (Measure)swModelDocExt.CreateMeasure();

            swMeasure.ArcOption = 0;
            bool status = swMeasure.Calculate(null);

            if ((status))
            {
                if ((!(swMeasure.Length == -1)))
                {
                    Debug.Print("Length: " + swMeasure.Length);
                }
                if ((!(swMeasure.Area == -1)))
                {
                    Debug.Print("Area: " + swMeasure.Area);
                }
                if ((!(swMeasure.ArcLength == -1)))
                {
                    Debug.Print("Arc length: " + swMeasure.ArcLength);
                }
                if ((!(swMeasure.ChordLength == -1)))
                {
                    Debug.Print("Chord length: " + swMeasure.ChordLength);
                }
                if ((!(swMeasure.Diameter == -1)))
                {
                    Debug.Print("Diameter: " + swMeasure.Diameter);
                }
                if ((!(swMeasure.Radius == -1)))
                {
                    Debug.Print("Radius: " + swMeasure.Radius);
                }
                if ((!(swMeasure.Perimeter == -1)))
                {
                    Debug.Print("Perimeter: " + swMeasure.Perimeter);
                }
                if ((!(swMeasure.X == -1)))
                {
                    Debug.Print("X coordinate: " + swMeasure.X);
                }
                if ((!(swMeasure.Y == -1)))
                {
                    Debug.Print("Y coordinate: " + swMeasure.Y);
                }
                if ((!(swMeasure.Z == -1)))
                {
                    Debug.Print("Z coordinate: " + swMeasure.Z);
                }
                if ((!(swMeasure.DeltaX == -1)))
                {
                    Debug.Print("DeltaX: " + swMeasure.DeltaX);
                }
                if ((!(swMeasure.DeltaY == -1)))
                {
                    Debug.Print("DeltaY: " + swMeasure.DeltaY);
                }
                if ((!(swMeasure.DeltaZ == -1)))
                {
                    Debug.Print("DeltaZ: " + swMeasure.DeltaZ);
                }
                if ((!(swMeasure.Angle == -1)))
                {
                    Debug.Print("Angle: " + swMeasure.Angle);
                }
                if ((!(swMeasure.CenterDistance == -1)))
                {
                    Debug.Print("Center distance: " + swMeasure.CenterDistance);
                }
                if ((!(swMeasure.NormalDistance == -1)))
                {
                    Debug.Print("Normal distance: " + swMeasure.NormalDistance);
                }
                if ((!(swMeasure.Distance == -1)))
                {
                    Debug.Print("Distance: " + swMeasure.Distance);
                }
                if ((!(swMeasure.TotalLength == -1)))
                {
                    Debug.Print("Total length: " + swMeasure.TotalLength);
                }
                if ((!(swMeasure.TotalArea == -1)))
                {
                    Debug.Print("Total area: " + swMeasure.TotalArea);
                }
                if ((swMeasure.IsParallel))
                {
                    Debug.Print("Is parallel: " + swMeasure.IsParallel);
                }
                if ((swMeasure.IsIntersect))
                {
                    Debug.Print("Is intersect: " + swMeasure.IsIntersect);
                }
                if ((swMeasure.IsPerpendicular))
                {
                    Debug.Print("Is perpendicular: " + swMeasure.IsPerpendicular);
                }
                if ((!(swMeasure.Projection == -1)))
                {
                    Debug.Print("Projection: " + swMeasure.Projection);
                }
                if ((!(swMeasure.Normal == -1)))
                {
                    Debug.Print("Normal: " + swMeasure.Normal);
                }
                if ((!(swMeasure.SpericalCenterDistance == -1)))
                {
                    Debug.Print("Spherical center distance: " + swMeasure.SpericalCenterDistance);
                }
                if ((swMeasure.IsConcentricSpheres))
                {
                    Debug.Print("Are concentric spheres: " + swMeasure.IsConcentricSpheres);
                }
            }
            else
            {
                Debug.Print("Invalid combination of selected entities.");
            }
            swModel.ClearSelection2(true);

            // 启动SolidWorks应用程序  
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
            if (swApp == null)
            {
                Debug.Print("Failed to connect");
                return;
            }
            ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
            ModelDocExtension swModelDocExt = (ModelDocExtension)swModel.Extension;
            bool boolstatus = swModelDocExt.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 0, null, 0);

            // Open sketch
            swModel.InsertSketch2(true);

            // Calculate the values for x, y, and z
            int i;
            int j;
            double[] x = {3,2,1,0 };
            double[] y = { 1,-1,1,0 };
     
            // Initialize the routine and sketch the first point of the spline at 0,0,0
            swModel.SketchSpline(-1, 0, 0, 0);

            // Sketch four more points of the spline
            for (j = 3; j >= 0; j += -1)
            {
                swModel.SketchSpline(j, x[j], y[j], 0);
            }

            // Exit sketch
            swModel.InsertSketch2(true);
        }
    }
}
