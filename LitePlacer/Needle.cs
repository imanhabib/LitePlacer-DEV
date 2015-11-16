using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace LitePlacer
{
    class NeedleClass
    {
        public struct NeedlePoint
        {
            public decimal Angle;
            public decimal X;  // X offset from nominal, in mm's, at angle
            public decimal Y;
        }

		public List<NeedlePoint> CalibrationPoints = new List<NeedlePoint>();

        private Camera Cam;
        private CNC Cnc;
        private static FormMain MainForm;

        public NeedleClass(Camera MyCam, CNC MyCnc, FormMain MainF)
        {
            MainForm = MainF;
            Calibrated = false;
            Cam = MyCam;
            Cnc = MyCnc;
            CalibrationPoints.Clear();
        }

        // private bool probingMode;
        public void ProbingMode(bool set, bool JSON)
        {
            int wait= 150;
            if(set)
            {
                if(JSON)
                {
                    // set in JSON mode
                    CNC_Write("{\"zsn\",0}");
                    Thread.Sleep(wait);
                    CNC_Write("{\"zsx\",1}");
                    Thread.Sleep(wait);
                    CNC_Write("{\"zzb\",0}");
                    Thread.Sleep(wait);
                    // probingMode = true;
                }
                else
                {
                    // set in text mode
                    CNC_Write("$zsn=0");
                    Thread.Sleep(wait);
                    CNC_Write("$zsx=1");
                    Thread.Sleep(wait);
                    CNC_Write("$zzb=0");
                    Thread.Sleep(wait);
                    // probingMode = true;
                }
            }            
            else
            {
                if (JSON)
                {
                    // clear in JSON mode
                    CNC_Write("{\"zsn\",3}");
                    Thread.Sleep(wait);
                    CNC_Write("{\"zsx\",2}");
                    Thread.Sleep(wait);
                    CNC_Write("{\"zzb\",2}");
                    Thread.Sleep(wait);
                    // probingMode = false;
                }
                else
                {
                    // clear in text mode
                    CNC_Write("$zsn=3");
                    Thread.Sleep(wait);
                    CNC_Write("$zsx=2");
                    Thread.Sleep(wait);
                    CNC_Write("$zzb=2");
                    Thread.Sleep(wait);
                    // probingMode = false;
                }
            }

        }


        public bool Calibrated { get; set; }

        public bool CorrectedPosition_m(decimal angle, out decimal X, out decimal Y)
        {
            if (Properties.Settings.Default.Placement_OmitNeedleCalibration)
            {
                X = 0;
                Y = 0;
                return true;
            };

            if (!Calibrated)
            {
                DialogResult dialogResult = MainForm.ShowMessageBox(
                    "Needle not calibrated. Calibrate now?",
                    "Needle not calibrated", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.No)
                {
                    X = 0;
                    Y = 0;
                    return false;
                };
                decimal CurrX = Cnc.CurrentX;
                decimal CurrY = Cnc.CurrentY;
                decimal CurrA = Cnc.CurrentA;
                if(!MainForm.CalibrateNeedle_m())
                {
                    X = 0;
                    Y = 0;
                    return false;
                }
                if (!MainForm.CNC_XYA_m(CurrX, CurrY, CurrA))
                {
                    X = 0;
                    Y = 0;
                    return false;
                }
            };

            while (angle < 0)
            {
                angle = angle + 360;
            };
            while (angle > 360)
            {
                angle = angle - 360;
            }
            // since we are not going to check the last point (which is the cal. value for 360)
            // in the for loop,we check that now
            if (angle > 359.98m)
            {
                X = CalibrationPoints[0].X;
                Y = CalibrationPoints[0].Y;
                return true;
            };

            for (int i = 0; i < CalibrationPoints.Count; i++)
            {
                if (Math.Abs(angle - CalibrationPoints[i].Angle) < 1)
                {
                    X = CalibrationPoints[i].X;
                    Y = CalibrationPoints[i].Y;
					return true;
                }
                if ((angle > CalibrationPoints[i].Angle)
                    &&
                    (angle < CalibrationPoints[i + 1].Angle)
                    &&
                    (Math.Abs(angle - CalibrationPoints[i + 1].Angle) > 1))
                {
                    // angle is between CalibrationPoints[i] and CalibrationPoints[i+1], and is not == CalibrationPoints[i+1]
                    decimal fract = (angle - CalibrationPoints[i+1].Angle) / (CalibrationPoints[i+1].Angle - CalibrationPoints[i].Angle);
                    X = CalibrationPoints[i].X + fract * (CalibrationPoints[i + 1].X - CalibrationPoints[i].X);
                    Y = CalibrationPoints[i].Y + fract * (CalibrationPoints[i + 1].Y - CalibrationPoints[i].Y);
					return true;
                }
            }
            MainForm.ShowMessageBox(
                "Needle Calibration value read: value not found",
                "Sloppy programmer error",
                MessageBoxButtons.OK);
            X = 0;
            Y = 0;
			return false;
        }


        public bool Calibrate(decimal tolerance)
        {
            if (Properties.Settings.Default.Placement_OmitNeedleCalibration)
            {
                return true;
            };

            CalibrationPoints.Clear();   // Presumably user changed the needle, and calibration is void no matter if we succeed here
            Calibrated = false;
            if (!Cam.IsRunning())
            {
                MainForm.ShowMessageBox(
                    "Attempt to calibrate needle, camera is not running.",
                    "Camera not running",
                    MessageBoxButtons.OK);
                return false;
            }

			decimal x = 0;
            decimal y = 0;
			int res = 0; ;
            for (int i = 0; i <= 3600; i = i + 225)
            {
                NeedlePoint Point = new NeedlePoint();
                Point.Angle = i / 10.0m;
				if (!CNC_A_m(Point.Angle))
				{
					return false;
				}
				for (int tries = 0; tries < 10; tries++)
				{
    				Thread.Sleep(100);
					res = Cam.GetClosestCircle(out x, out y, (double)tolerance);
					if (res != 0)
					{
						break;
					}

					if (tries >= 9)
					{
                        MainForm.ShowMessageBox(
							"Needle calibration: Can't see Needle",
							"No Circle found",
							MessageBoxButtons.OK);
						return false;
					}
				}
                if (res == 0)
                {
                    MainForm.ShowMessageBox(
                        "Needle Calibration: Can't find needle",
                        "No Circle found",
                        MessageBoxButtons.OK);
                    return false;
                }
                //if (res > 1)
                //{
                //    MessageBox.Show(
                //        "Needle Calibration: Ambiguous regognition result",
                //        "Too macy circles in focus",
                //        MessageBoxButtons.OK);
                //    return false;
                //}

                Point.X = x * Properties.Settings.Default.UpCam_XmmPerPixel;
                Point.Y = y * Properties.Settings.Default.UpCam_YmmPerPixel;
				// MainForm.DisplayText("A: " + Point.Angle.ToString("0.000") + ", X: " + Point.X.ToString("0.000") + ", Y: " + Point.Y.ToString("0.000"));
                CalibrationPoints.Add(Point);
            }
            Calibrated = true;
            return true;
        }

        public bool Move_m(decimal X, decimal Y, decimal A)
        {
            decimal dX;
            decimal dY;
			MainForm.DisplayText("Needle.Move_m(): X= " + X.ToString() + ", Y= " + Y.ToString() + ", A= " + A.ToString());
			if (!CorrectedPosition_m(A, out dX, out dY))
			{
				return false;
			};
            decimal Xoff = Properties.Settings.Default.DownCam_NeedleOffsetX;
            decimal Yoff = Properties.Settings.Default.DownCam_NeedleOffsetY;
            return CNC_XYA(X + Xoff + dX, Y + Yoff + dY, A);
        }



        // =================================================================================
        // CNC interface functions
        // =================================================================================
        
        private bool CNC_A_m(decimal A)
        {
			return MainForm.CNC_A_m(A);
        }

        private bool CNC_XY_m(decimal X, decimal Y)
        {
			return MainForm.CNC_XY_m(X, Y);
        }

        private bool CNC_XYA(decimal X, decimal Y, decimal A)
        {
			return MainForm.CNC_XYA_m(X, Y, A);
        }

        private bool CNC_Write(string s)
        {
			return MainForm.CNC_Write_m(s);
        }
    }
}
