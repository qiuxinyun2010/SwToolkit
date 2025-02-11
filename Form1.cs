using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using static System.Security.Cryptography.ECCurve;

namespace SwToolkit
{
    public partial class Form1 : Form
    {
        private SldWorks swApp;
        private ModelDoc2 swModel;
        private Process pythonProc;

        private static readonly Dictionary<swSurfaceTypes_e, string> SurfaceTypeMaps = new Dictionary<swSurfaceTypes_e, string>
        {
            { swSurfaceTypes_e.PLANE_TYPE, "Plane" },
            { swSurfaceTypes_e.CYLINDER_TYPE, "Cylinder" },
            { swSurfaceTypes_e.CONE_TYPE, "Cone" },
            { swSurfaceTypes_e.SPHERE_TYPE, "Sphere" },
            { swSurfaceTypes_e.TORUS_TYPE, "Torus" },
            { swSurfaceTypes_e.SREV_TYPE, "SpinSurface" },
            { swSurfaceTypes_e.EXTRU_TYPE, "SweepSurface" },
            { swSurfaceTypes_e.BSURF_TYPE, "BSurface" },
            { swSurfaceTypes_e.OFFSET_TYPE, "OffsetSurface" },
            { swSurfaceTypes_e.BLEND_TYPE, "BlendSurface" },
        };

        private static readonly Dictionary<swCurveTypes_e, string> CurveTypeMaps = new Dictionary<swCurveTypes_e, string>
        {
            { swCurveTypes_e.LINE_TYPE, "Line" },
            { swCurveTypes_e.CIRCLE_TYPE, "Circle" },
            { swCurveTypes_e.ELLIPSE_TYPE, "Ellipse" },
            { swCurveTypes_e.BCURVE_TYPE, "BCurve" },
            { swCurveTypes_e.SPCURVE_TYPE, "SPCurve" },
            { swCurveTypes_e.INTERSECTION_TYPE, "ICurve" },
            { swCurveTypes_e.TRIMMED_TYPE, "TRCurve" },
            { swCurveTypes_e.CONSTPARAM_TYPE, "CONSTPARAM_TYPE" },
        };

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
            // 获取进程 ID
            int processId = this.swApp.GetProcessID();

            // 打印进程 ID
            MessageBox.Show($"SolidWorks 进程 ID: {processId}");
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

        private void createSpline(object sender, EventArgs e)
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
            double[] x = { 3, 2, 1, 0 };
            double[] y = { 1, -1, 1, 0 };

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

        private void connectSw()
        {
            swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
            if (swApp == null)
            {
                Debug.Print("Failed to connect");
                return;
            }
            swModel = (ModelDoc2)swApp.ActiveDoc;
        }

        private Process ExecutePythonScriptWithPID(int pid)
        {
            try
            {
                // 命令行要执行的命令，例如：python hook.py 1234
                string command = "python hook.py";
                string arguments = pid.ToString();

                // 启动一个新的进程来执行命令
                Process process = new Process();
                process.StartInfo.FileName = "cmd.exe"; // 或 "powershell.exe" 如果你使用的是 PowerShell
                process.StartInfo.Arguments = $"/K {command} {arguments}";
                process.StartInfo.UseShellExecute = false;
                process.StartInfo.RedirectStandardInput = true;
                process.StartInfo.RedirectStandardOutput = false;
                process.StartInfo.RedirectStandardError = false;
                process.StartInfo.CreateNoWindow = false;

                process.Start();
                Thread.Sleep(1000);
                pythonProc = process;
                return process;
                //  string output = process.StandardOutput.ReadToEnd();
                //  string error = process.StandardError.ReadToEnd();
                //  process.WaitForExit();

                // 输出结果
                // MessageBox.Show($"命令行输出:\n{output}\n\n错误输出:\n{error}");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"执行命令时出错: {ex.Message}");
            }
            return null;
        }

        private void closePythonProc()
        {
            Thread.Sleep(1000);
            using (StreamWriter writer = pythonProc.StandardInput)
            {
                writer.Write((char)26);
                writer.WriteLine();
            }
            pythonProc.WaitForExit();
            pythonProc = null;
        }

        private string[] readHookOutput(string keyword)
        {
            Thread.Sleep(100);
            string[] lines = File.ReadAllLines("temp.json");
            string[] faceLines = Array.FindAll(lines, line => line.Contains(keyword));
            foreach (string line in lines)
            {
                if (line.Contains(keyword))
                {
                    Console.WriteLine(line);
                }
            }
            return faceLines;
        }

        private void exportXt(object sender, EventArgs e)
        {
            this.swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
            if (swApp == null)
            {
                Debug.Print("Failed to connect");
                return;
            }
            ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
            int processId = this.swApp.GetProcessID();
            Process proc = this.ExecutePythonScriptWithPID(processId);
            if (swModel.GetType() == (int)swDocumentTypes_e.swDocPART || swModel.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY)
            {
                ModelDocExtension swModExt = (ModelDocExtension)swModel.Extension;

                int error = 0;

                int warnings = 0;

                //设定导出坐标系

                var setRes = swModel.Extension.SetUserPreferenceString(16, 0, "CustomerCS");

                //设置导出版本
                swApp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swParasolidOutputVersion, (int)swParasolidOutputVersion_e.swParasolidOutputVersion_161);
                string currentDirectory = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                string fileName = $"export/export_{DateTime.Now:yyyyMMdd_HHmmss}.x_t";
                string exportPath = Path.Combine(currentDirectory, fileName);

                swModExt.SaveAs(exportPath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, null, ref error, ref warnings);

                proc.CloseMainWindow();
            }
            else if (swModel.GetType() == (int)swDocumentTypes_e.swDocDRAWING)
            {
                ModelDocExtension swModExt = (ModelDocExtension)swModel.Extension;

                int error = 0;

                int warnings = 0;

                //设置dxf 导出版本 R14
                swApp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swDxfVersion, 2);

                //是否显示 草图
                swModel.SetUserPreferenceToggle(196, false);

                swModExt.SaveAs(@"C:\export.dxf", (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, null, ref error, ref warnings);
                MessageBox.Show("已导出成功");
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
            if (swApp == null)
            {
                Debug.Print("Failed to connect");
                return;
            }
            ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
            int processId = this.swApp.GetProcessID();
            // Process proc = this.ExecutePythonScriptWithPID(processId);

            // 弹出消息框
            SelectionMgr swSelMgr = (SelectionMgr)swModel.SelectionManager;
            var swEnt = (Entity)swSelMgr.GetSelectedObject6(1, -1);
            if (swEnt is IFace2)
            {
                Face2 face = (Face2)swEnt;
                Surface surf = face.GetSurface();
                surf.EvaluateAtPoint(0, 0, 0);
                swSurfaceTypes_e surfType = (swSurfaceTypes_e)surf.Identity();
                SurfaceTypeMaps.TryGetValue(surfType, out string translation);
                MessageBox.Show($"选中对象是: {translation}");
            }
            else if (swEnt is IEdge)
            {
                Edge edge = (Edge)swEnt;
                Curve curve = edge.GetCurve();
                swCurveTypes_e curveType = (swCurveTypes_e)curve.Identity();
                CurveTypeMaps.TryGetValue(curveType, out string translation);
                MessageBox.Show($"选中对象是: {translation}");
            }
            else
            {
                MessageBox.Show($"选中对象是: {(swSelectType_e)(swEnt.GetType())}");
            }

            //var lines = readHookOutput("surf");

            // proc.CloseMainWindow();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
            if (swApp == null)
            {
                Debug.Print("Failed to connect");
                return;
            }
            ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
            int pid = swApp.GetProcessID();
            Process proc = ExecutePythonScriptWithPID(pid);
            SelectionMgr swSelMgr = (SelectionMgr)swModel.SelectionManager;
            var swEnt = (Entity)swSelMgr.GetSelectedObject6(1, -1);
            if (swEnt is IFace2)
            {
                Face2 face = (Face2)swEnt;
                Body2 body = face.GetBody();
                Surface surf = face.GetSurface();
                //surf.EvaluateAtPoint(0, 0, 0);
                swSurfaceTypes_e surfType = (swSurfaceTypes_e)surf.Identity();
            }
            else if (swEnt is IEdge)
            {
                Edge edge = (Edge)swEnt;
                Curve curve = edge.GetCurve();
                swCurveTypes_e curveType = (swCurveTypes_e)curve.Identity();
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            connectSw();
            ExecutePythonScriptWithPID(swApp.GetProcessID());
            swModel.ForceRebuild3(false);
            Thread.Sleep(1000);
            closePythonProc();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            connectSw();
            ExecutePythonScriptWithPID(swApp.GetProcessID());
            swModel.EditRebuild3();
            Thread.Sleep(1000);
            closePythonProc();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            connectSw();
            swApp.Visible = true;
            swModel.ViewZoomtofit2();
            ExecutePythonScriptWithPID(swApp.GetProcessID());
            bool boolstatus = true;
            int lErrors = 0;
            int lWarnings = 0;
            swModel.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref lErrors, ref lWarnings);
            // Thread.Sleep(1000);
            // closePythonProc();
        }
    }
}