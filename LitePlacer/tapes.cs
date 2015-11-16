using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;

namespace LitePlacer
{
	class TapesClass
	{
        private DataGridView Grid;
        private DataGridView CustomGrid;
        private NeedleClass Needle;
		private FormMain MainForm;
		private Camera DownCamera;
		private CNC Cnc;

        public TapesClass(DataGridView gr, DataGridView custom, NeedleClass ndl, Camera cam, CNC c, FormMain MainF)
		{
            CustomGrid = custom;
            Grid = gr;
			Needle = ndl;
			DownCamera = cam;
			MainForm = MainF;
			Cnc = c;
		}

		// ========================================================================================
		// ClearAll(): Resets TapeNumber positions and pickup/place Z's.
		public void ClearAll()
		{
			for (int tape = 0; tape < Grid.Rows.Count; tape++)
			{
				Grid.Rows[tape].Cells["Next_Column"].Value = "1";
				Grid.Rows[tape].Cells["PickupZ_Column"].Value = "--";
				Grid.Rows[tape].Cells["PlaceZ_Column"].Value = "--";
				Grid.Rows[tape].Cells["NextX_Column"].Value = Grid.Rows[tape].Cells["X_Column"].Value;
				Grid.Rows[tape].Cells["NextY_Column"].Value = Grid.Rows[tape].Cells["Y_Column"].Value;
			}
		}

		// ========================================================================================
		// Reset(): Resets one tape position and pickup/place Z's.
		public void Reset(int tape)
		{
			Grid.Rows[tape].Cells["Next_Column"].Value = "1";
			Grid.Rows[tape].Cells["PickupZ_Column"].Value = "--";
			Grid.Rows[tape].Cells["PlaceZ_Column"].Value = "--";
			Grid.Rows[tape].Cells["NextX_Column"].Value = Grid.Rows[tape].Cells["X_Column"].Value;
			Grid.Rows[tape].Cells["NextY_Column"].Value = Grid.Rows[tape].Cells["Y_Column"].Value;
		}

		// ========================================================================================
		// IdValidates_m(): Checks that tape with description of "Id" exists.
		// TapeNumber is set to the corresponding row of the Grid.
		public bool IdValidates_m(string Id, out int Tape)
		{
			Tape = -1;
			foreach (DataGridViewRow Row in Grid.Rows)
			{
				Tape++;
				if (Row.Cells["IdColumn"].Value.ToString() == Id)
				{
					return true;
				}
			}
			MainForm.ShowMessageBox(
				"Did not find tape " + Id.ToString(),
				"Tape data error",
				MessageBoxButtons.OK
			);
			return false;
		}

        // ========================================================================================
        // Fast placement:
        // ========================================================================================
        // The process measures last and first hole positions for a row in job data. It keeps track about
        // the hole location in next columns, and in this process, these are measured, not approximated.
        // The part numbers are found with the GetPartLocationFromHolePosition_m() routine.

        public bool FastParametersOk { get; set; }  // if we should use fast placement in the first place
        public decimal FastXstep { get; set; }       // steps for hole positions
        public decimal FastYstep { get; set; }
        public decimal FastXpos { get; set; }       // we don't want to mess with tape definitions
        public decimal FastYpos { get; set; }

        // ========================================================================================
        // PrepareForFastPlacement_m: Called before starting fast placement

        public bool PrepareForFastPlacement_m(string TapeID, int ComponentCount)
        {
            int TapeNum;
            if (!IdValidates_m(TapeID, out TapeNum))
            {
                FastParametersOk = false;
                return false;
            }
            int first;
            if (!int.TryParse(Grid.Rows[TapeNum].Cells["Next_Column"].Value.ToString(), out first))
            {
                MainForm.ShowMessageBox(
                    "Bad data at next column",
                    "Sloppy programmer error",
                    MessageBoxButtons.OK);
                FastParametersOk = false;
                return false;
            }
            int last = first + ComponentCount-1;
            // measure holes
            decimal lastX = 0;
            decimal lastY = 0;
            decimal firstX = 0;
            decimal firstY = 0;
            if (!GetPartHole_m(TapeNum, last, out lastX, out lastY))
            {
                FastParametersOk = false;
                return false;
            }
            if (last!= first)
            {
                if (!GetPartHole_m(TapeNum, first, out firstX, out firstY))
                {
                    FastParametersOk = false;
                    return false;
                }
            }
            else
            {
                firstX = lastX;
                firstY = lastY;
            }

            FastXpos = firstX;
            FastYpos = firstY;
            if (ComponentCount>1)
            {
                FastXstep = (lastX - firstX) / (ComponentCount-1);
                FastYstep = (lastY - firstY) / (ComponentCount-1);
            }
            else
            {
                FastXstep = 0;
                FastYstep = 0;
            }

            MainForm.DisplayText("Fast parameters:");
            MainForm.DisplayText("First X: " + firstX.ToString() + ", Y: " + firstY.ToString());
            MainForm.DisplayText("Last X: " + lastX.ToString() + ", Y: " + lastY.ToString());
            MainForm.DisplayText("Step X: " + FastXstep.ToString() + ", Y: " + FastYstep.ToString());

            return true;
        }

        // ========================================================================================
        // IncrementTape_Fast(): Updates count and next hole locations for a tape
        // Like IncrementTape(), but just using the fast parameters
        public bool IncrementTape_Fast(int TapeNum)
        {
            int pos;
            if (!int.TryParse(Grid.Rows[TapeNum].Cells["Next_Column"].Value.ToString(), out pos))
            {
                MainForm.ShowMessageBox(
                    "Bad data at next column",
                    "Sloppy programmer error",
                    MessageBoxButtons.OK);
                return false;
            }
            if (Grid.Rows[TapeNum].Cells["WidthColumn"].Value.ToString() == "8/2mm" )
	        {
		        if ((pos%2)==0)      // increment hole location only on every other "next" value on 2mm parts
	            {
                    FastXpos += FastXstep*2;
                    FastYpos += FastYstep*2;
                    Grid.Rows[TapeNum].Cells["NextX_Column"].Value = FastXpos.ToString("0.000", CultureInfo.InvariantCulture);
                    Grid.Rows[TapeNum].Cells["NextY_Column"].Value = FastYpos.ToString("0.000", CultureInfo.InvariantCulture);
                }
	        }
            else
            {
                FastXpos += FastXstep;
                FastYpos += FastYstep;
                Grid.Rows[TapeNum].Cells["NextX_Column"].Value = FastXpos.ToString("0.000", CultureInfo.InvariantCulture);
                Grid.Rows[TapeNum].Cells["NextY_Column"].Value = FastYpos.ToString("0.000", CultureInfo.InvariantCulture);
            }
            
            pos += 1;
            Grid.Rows[TapeNum].Cells["Next_Column"].Value = pos.ToString();
           return true;
        }

        // ========================================================================================
        // GetTapeParameters_m(): 
        // Get from the indicated tape the dW, part center pos from hole and Pitch, distance from one part to another
        private bool GetTapeParameters_m(int Tape, out int customTapeNum, out decimal dW, out decimal fromHole, out decimal pitch)
        {
            customTapeNum= -1;      // not custom
            dW = 0;
            pitch = 0;
            fromHole = -2;
            string Width = Grid.Rows[Tape].Cells["WidthColumn"].Value.ToString();
            // TapeNumber measurements: 
            switch (Width)
            {
                case "8/2mm":
                    dW = 3.50m;
                    pitch = 2;
                    break;
                case "8/4mm":
                    dW = 3.50m;
                    pitch = 4;
                    break;

                case "12/4mm":
                    dW = 5.50m;
                    pitch = 4;
                    break;
                case "12/8mm":
                    dW = 5.50m;
                    pitch = 8m;
                    break;

                case "16/4mm":
                    dW = 7.50m;
                    pitch = 4m;
                    break;
                case "16/8mm":
                    dW = 7.50m;
                    pitch = 8;
                    break;
                case "16/12mm":
                    dW = 7.50m;
                    pitch = 12;
                    break;

                case "24/4mm":
                    dW = 11.50m;
                    pitch = 4;
                    break;
                case "24/8mm":
                    dW = 11.50m;
                    pitch = 8;
                    break;
                case "24/12mm":
                    dW = 11.50m;
                    pitch = 12;
                    break;
                case "24/16mm":
                    dW = 11.50m;
                    pitch = 16;
                    break;
                case "24/20mm":
                    dW = 11.50m;
                    pitch = 20;
                    break;

                default:
                    if (!FindCustomTapeParameters(Width, out customTapeNum, out dW, out fromHole, out pitch))
                    {
                        MainForm.ShowMessageBox(
                            "Bad data at Tape #" + Tape.ToString() + ", Width",
                            "Tape data error",
                            MessageBoxButtons.OK
                        );
                        return false;
                    }
                    break;
            }
            return true;
        }
		// ========================================================================================

        // ========================================================================================
        // GetPartHole_m(): Measures X,Y location of the hole corresponding to part number
        public bool GetPartHole_m(int TapeNum, int PartNum, out decimal resultX, out decimal resultY)
        {
            resultX = 0;
            resultY = 0;

            // Get start points
            decimal x = 0;
            decimal y = 0;
            if (!decimal.TryParse(Grid.Rows[TapeNum].Cells["X_Column"].Value.ToString(), out x))
            {
                MainForm.ShowMessageBox(
                    "Bad data at Tape " + TapeNum.ToString() + ", X",
                    "Tape data error",
                    MessageBoxButtons.OK
                );
                return false;
            }
            if (!decimal.TryParse(Grid.Rows[TapeNum].Cells["Y_Column"].Value.ToString(), out y))
            {
                MainForm.ShowMessageBox(
                    "Bad data at Tape " + TapeNum.ToString() + ", Y",
                    "Tape data error",
                    MessageBoxButtons.OK
                );
                return false;
            }

            // Get the hole location guess
            decimal dW;
            decimal pitch;
            decimal fromHole;
            int CustomTapeNum = -1;
            if (!GetTapeParameters_m(TapeNum, out CustomTapeNum, out dW, out fromHole, out pitch))
            {
                return false;
            }
            if (Math.Abs((double)pitch-2.0)<0.01) // if pitch ==2
            {
                PartNum = (PartNum +1)/ 2;
                pitch = 4;
            }
            decimal dist = (PartNum-1) * pitch; // This many mm's from start
            switch (Grid.Rows[TapeNum].Cells["OrientationColumn"].Value.ToString())
            {
                case "+Y":
                    y = y + dist;
                    break;

                case "+X":
                    x = x + dist;
                    break;

                case "-Y":
                    y = y - dist;
                    break;

                case "-X":
                    x = x - dist;
                    break;

                default:
                    MainForm.ShowMessageBox(
                        "Bad data at Tape #" + TapeNum.ToString() + ", Orientation",
                        "Tape data error",
                        MessageBoxButtons.OK
                    );
                    return false;
            }
            // X, Y now hold the first guess
            // For custom tapes, we might not need to move at all:

            if (CustomTapeNum != -1)  // tape is custom
            {
                if (!GetCustomPartHole_m(CustomTapeNum, PartNum, x, y, out resultX, out resultY))
                {
                    return false;
                }
                return true;
            }

            // Measuring standard tapes
            if (!SetCurrentTapeMeasurement_m(TapeNum))  // having the measurement setup here helps with the automatic gain lag
            {
                return false;
            }
            // Go there:
            if (!MainForm.CNC_XY_m(x, y))
            {
                return false;
            };

            // get hole exact location:
            if (!MainForm.GoToCircleLocation_m(1.8, 0.1, out x, out y))
            {
                MainForm.ShowMessageBox(
                    "Can't find tape hole",
                    "Tape error",
                    MessageBoxButtons.OK
                );
                return false;
            }
            resultX = Cnc.CurrentX + x;
            resultY = Cnc.CurrentY + y;
            return true;
        }

		// ========================================================================================
		// GetPartLocationFromHolePosition_m(): Returns the location and rotation of the part
        // Input is the exact (measured) location of the hole

        public bool GetPartLocationFromHolePosition_m(int Tape, decimal x, decimal y, out decimal partX, out decimal partY, out decimal a)
        {
            partX = 0;
            partY = 0;
            a = 0;

            decimal dW;	// Part center pos from hole, tape width direction. Varies.
            decimal dL;   // Part center pos from hole, tape lenght direction. -2mm on all standard tapes
            decimal pitch;  // Distance from one part to another

            int CustomTapeNum = -1;
            if (!GetTapeParameters_m(Tape, out CustomTapeNum, out dW, out dL, out pitch))
	        {
		        return false;
	        }
            dL = -dL; // so up is + etc.
			// TapeNumber orientation: 
			// +Y: Holeside of tape is right, part is dW(mm) to left, dL(mm) down from hole, A= 0
			// +X: Holeside of tape is down, part is dW(mm) up, dL(mm) to left from hole, A= -90
			// -Y: Holeside of tape is left, part is dW(mm) to right, dL(mm) up from hole, A= -180
			// -X: Holeside of tape is up, part is dW(mm) down, dL(mm) to right from hole, A=-270
            int pos;
			if (!int.TryParse(Grid.Rows[Tape].Cells["Next_Column"].Value.ToString(), out pos))
			{
				MainForm.ShowMessageBox(
					"Bad data at Tape " + Tape.ToString() + ", Next",
					"Tape data error",
					MessageBoxButtons.OK
				);
				return false;
			}     
            // if pitch == 2 and part# is odd, DL=2, other
            if (Math.Abs((double)pitch - 2) < 0.01)
            {
                if ((pos % 2) == 1)
                {
                    dL = 2;
                }
                else
                {
                    dL = 0;
                }
            }

			switch (Grid.Rows[Tape].Cells["OrientationColumn"].Value.ToString())
			{
				case "+Y":
					partX = x - dW;
					partY = y - dL;
					a = 0;
					break;

				case "+X":
					partX = x - dL;
					partY = y + dW;
					a = -90;
					break;

				case "-Y":
					partX = x + dW;
					partY = y + dL;
					a = -180;
					break;

				case "-X":
					partX = x - dL;
					partY = y - dW;
					a = -270;
					break;

				default:
					MainForm.ShowMessageBox(
						"Bad data at Tape #" + Tape.ToString() + ", Orientation",
						"Tape data error",
						MessageBoxButtons.OK
					);
					return false;
			}
			// rotation:
			if (Grid.Rows[Tape].Cells["RotationColumn"].Value == null)
			{
				MainForm.ShowMessageBox(
					"Bad data at tape " + Grid.Rows[Tape].Cells["IdColumn"].Value.ToString() +" rotation",
					"Assertion error",
					MessageBoxButtons.OK
				);
				return false;
			}
			switch (Grid.Rows[Tape].Cells["RotationColumn"].Value.ToString())
			{
				case "0deg.":
					break;

				case "90deg.":
					a += 90;
					break;

				case "180deg.":
					a += 180;
					break;

				case "270deg.":
					a += 270;
					break;

				default:
					MainForm.ShowMessageBox(
						"Bad data at Tape " + Grid.Rows[Tape].Cells["IdColumn"].Value.ToString() + " rotation",
						"Tape data error",
						MessageBoxButtons.OK
					);
					return false;
					// break;
			};
			while (a > 360.1m)
			{
				a -= 360;
			}
			while (a < 0)
			{
				a += 360;
			};
            return true;
        }

        // ========================================================================================
        // IncrementTape(): Updates count and next hole locations for a tape
        // The caller knows the current hole location, so we don't need to re-measure them
        public bool IncrementTape(int tape, decimal holeX, decimal holeY)
        {
            decimal dW;	// Part center pos from hole, tape width direction. Varies.
            decimal dL;   // Part center pos from hole, tape lenght direction. -2mm on all standard tapes
            decimal pitch;  // Distance from one part to another
            int CustomTapeNum = -1;
            if (!GetTapeParameters_m(tape, out CustomTapeNum, out dW, out dL, out pitch))
            {
                return false;
            }

            int pos;
            if (!int.TryParse(Grid.Rows[tape].Cells["Next_Column"].Value.ToString(), out pos))
			{
				MainForm.ShowMessageBox(
                    "Bad data at Tape " + Grid.Rows[tape].Cells["IdColumn"].Value.ToString() + ", next",
					"S´loppy programmer error",
					MessageBoxButtons.OK
				);
				return false;
			}

            // Set next hole approximate location. On 2mm part pitch, increment only at even part count.
            if (Math.Abs((double)pitch - 2) < 0.000001)
            {
                if ((pos % 2) != 0)
                {
                    pitch = 0;
                }
                else
                {
                    pitch = 4;
                }
            };
            switch (Grid.Rows[tape].Cells["OrientationColumn"].Value.ToString())
            {
                case "+Y":
                    holeY = holeY + pitch;
                    break;

                case "+X":
                    holeX = holeX + pitch;
                    break;

                case "-Y":
                    holeY = holeY - pitch;
                    break;

                case "-X":
                    holeX = holeX - pitch;
                    break;
            };
            Grid.Rows[tape].Cells["NextX_Column"].Value = holeX.ToString();
            Grid.Rows[tape].Cells["NextY_Column"].Value = holeY.ToString();
            // increment next count
            pos++;
            Grid.Rows[tape].Cells["Next_Column"].Value = pos.ToString();
            return true;
        }

		// ========================================================================================
		// GotoNextPartByMeasurement_m(): Takes needle to exact location of the part, tape and part rotation taken in to account.
		// The hole position is measured on each call using tape holes and knowledge about tape width and pitch (see EIA-481 standard).
		// Id tells the tape name. 
        // The caller needs the hole coordinates and tape number later in the process, but they are measured and returned here.
        public bool GotoNextPartByMeasurement_m(int tapeNumber, out decimal holeX, out decimal holeY)
		{
            holeX = 0;
            holeY = 0;
            int customTapeNumber;
            // If this is a custom tape that doesn't use location marks, we'll return the set position:
            if (IsCustomTape(tapeNumber, out customTapeNumber))
            {
                if (!(CustomGrid.Rows[customTapeNumber].Cells["UsesLocationMarks_Column"].Value.ToString() == "true"))
                {
                    if (!decimal.TryParse(CustomGrid.Rows[customTapeNumber].Cells["PartOffsetX_Column"].Value.ToString(), out holeX))
                    {
                        MainForm.ShowMessageBox(
                            "Bad data at custom tape " + CustomGrid.Rows[customTapeNumber].Cells["Name_Column"].Value.ToString() + ", X offset",
                            "Custom tape data error",
                            MessageBoxButtons.OK
                        );
                        return false;
                    }
                    if (!decimal.TryParse(CustomGrid.Rows[customTapeNumber].Cells["PartOffsetY_Column"].Value.ToString(), out holeY))
                    {
                        MainForm.ShowMessageBox(
                            "Bad data at custom tape " + CustomGrid.Rows[customTapeNumber].Cells["Name_Column"].Value.ToString() + ", Y offset",
                            "Custom tape data error",
                            MessageBoxButtons.OK
                        );
                        return false;
                    }
                    // Go there:
                    if (!MainForm.CNC_XY_m(holeX, holeY))
                    {
                        return false;
                    };
                }
            }

            // Normal case:
			// Go to next hole approximate location:
			if (!SetCurrentTapeMeasurement_m(tapeNumber))  // having the measurement setup here helps with the automatic gain lag
				return false;

			decimal nextX= 0;
            decimal nextY = 0;
            if (!decimal.TryParse(Grid.Rows[tapeNumber].Cells["NextX_Column"].Value.ToString(), out nextX))
			{
				MainForm.ShowMessageBox(
                    "Bad data at Tape " + Grid.Rows[tapeNumber].Cells["IdColumn"].Value.ToString() + ", Next X",
					"Tape data error",
					MessageBoxButtons.OK
				);
				return false;
			}

            if (!decimal.TryParse(Grid.Rows[tapeNumber].Cells["NextY_Column"].Value.ToString(), out nextY))
			{
				MainForm.ShowMessageBox(
                    "Bad data at Tape " + Grid.Rows[tapeNumber].Cells["IdColumn"].Value.ToString() + ", Next Y",
					"Tape data error",
					MessageBoxButtons.OK
				);
				return false;
			}
			// Go there:
            if (!MainForm.CNC_XY_m(nextX, nextY))
			{
				return false;
			};

			// Get hole exact location:
            // We want to find the hole less than 2mm from where we think it should be. (Otherwise there is a risk
			// of picking a wrong hole.)
            if (!MainForm.GoToCircleLocation_m(1.8, 0.5, out holeX, out holeY))
			{
				MainForm.ShowMessageBox(
					"Can't find tape hole",
					"Tape error",
					MessageBoxButtons.OK
				);
				return false;
			}
			// The hole locations are:
            holeX = Cnc.CurrentX + holeX;
            holeY = Cnc.CurrentY + holeY;

            // ==================================================
            // find the part location and go there:
            decimal partX = 0;
            decimal partY = 0;
            decimal A = 0;

            if (!GetPartLocationFromHolePosition_m(tapeNumber, holeX, holeY, out partX, out partY, out A))
            {
                MainForm.ShowMessageBox(
                    "Can't find tape hole",
                    "Tape error",
                    MessageBoxButtons.OK
                );
            }

			// Now, PartX, PartY, A tell the position of the part. Take needle there:
			if (!Needle.Move_m(partX, partY, A))
			{
				return false;
			}

			return true;
		}	// end GotoNextPartByMeasurement_m


		// ========================================================================================
		// SetCurrentTapeMeasurement_m(): sets the camera measurement parameters according to the tape type.
		private bool SetCurrentTapeMeasurement_m(int row)
		{
			switch (Grid.Rows[row].Cells["TypeColumn"].Value.ToString())
			{
				case "Paper (White)":
					MainForm.SetPaperTapeMeasurement();
					Thread.Sleep(200);   // for automatic camera gain to have an effect
					return true;

				case "Black Plastic":
					MainForm.SetBlackTapeMeasurement();
					Thread.Sleep(200);   // for automatic camera gain to have an effect
					return true;

				case "Clear Plastic":
					MainForm.SetClearTapeMeasurement();
					Thread.Sleep(200);   // for automatic camera gain to have an effect
					return true;

				default:
					MainForm.ShowMessageBox(
						"Bad Type data on row " + row.ToString() + ": " + Grid.Rows[row].Cells["TypeColumn"].Value.ToString(),
						"Bad Type data",
						MessageBoxButtons.OK
					);
					return false;
			}
		}

	// ========================================================================================
	// Custom Tapes (or trays, feeders etc.):

        private bool IsCustomTape(int TapeNumber, out int CustomTapeNumber)
        {
            // Tells if TapeNumber is custom tape, sets CustomTapeNumber
            string TapeName = Grid.Rows[TapeNumber].Cells["IdColumn"].Value.ToString();
            for (int i = 0; i < CustomGrid.RowCount-1; i++)
            {
                if (CustomGrid.Rows[i].Cells["Name_Column"].Value.ToString() == TapeName)
                {
                    CustomTapeNumber = i;
                    return true;
                }
            }
            CustomTapeNumber = -1;
            return false;


        }

        private bool GetCustomPartHole_m(int customTapeNum, int partNum, decimal x, decimal y, out decimal resultX, out decimal resultY)
        {
            resultX = 0;
            resultY = 0;
            return false;
        }

        // ========================================================================================
        // FindCustomTapeParameters(): We did not find the tape width from standard tapes, so the tape must be a custom tape.
        // This routine finds the tape number and the parameters from the custom tape name:
        private bool FindCustomTapeParameters(string name, out int customTapeNum, out decimal offsetY, out decimal offsetX, out decimal pitch)
        {
            offsetY = 0;
            offsetX = 0;
            pitch = 0;
            customTapeNum = -1;
            foreach (DataGridViewRow GridRow in CustomGrid.Rows)
            {
                if (GridRow.Cells["Name_Column"].Value.ToString() == name)
                {
                    // Found it!
                    customTapeNum= GridRow.Index;
                    DataGridViewRow Row = CustomGrid.Rows[customTapeNum];
                    if (!decimal.TryParse(Row.Cells["PitchColumn"].Value.ToString(), out pitch))
                    {
                        MainForm.ShowMessageBox(
                            "Bad data at custom tape " + name +", Pitch column",
                            "Data error",
                            MessageBoxButtons.OK
                        );
                        return false;
                    }

                    if (Convert.ToBoolean(Row.Cells["UsesLocationMarks_Column"].Value) == true)
                    {
                        if (!decimal.TryParse(Row.Cells["PartOffsetX_Column"].Value.ToString(), out offsetX))
                        {
                            MainForm.ShowMessageBox(
                                "Bad data at custom tape " + name + ", Part Offset X column",
                                "Data error",
                                MessageBoxButtons.OK
                            );
                            return false;
                        }
                        if (!decimal.TryParse(Row.Cells["PartOffsetY_Column"].Value.ToString(), out offsetY))
                        {
                            MainForm.ShowMessageBox(
                                "Bad data at custom tape " + name + ", Part Offset Y column",
                                "Data error",
                                MessageBoxButtons.OK
                            );
                            return false;
                        }
                        return true;
                    }
                    return true;
                }  // end "found it"
            }
            MainForm.ShowMessageBox(
                "Did not find custom tape " + name,
                "Data error",
                MessageBoxButtons.OK
            );
            return false;
        }

        // ========================================================================================
        // Custom tapes are recognized by their name, which should be placed in the Width column

        private void ResetTapeWidths(DataGridViewComboBoxCell box)
        {
            box.Items.Clear();
            box.Items.Add("8/2mm");
            box.Items.Add("8/4mm");
            box.Items.Add("12/4mm");
            box.Items.Add("12/8mm");
            box.Items.Add("16/4mm");
            box.Items.Add("16/8mm");
            box.Items.Add("16/12mm");
            box.Items.Add("24/4mm");
            box.Items.Add("24/8mm");
            box.Items.Add("24/12mm");
            box.Items.Add("24/16mm");
            box.Items.Add("24/20mm");
        }

        public void AddCustomTapesToTapes()
        {
            if (Grid.RowCount==0)
            {
                return;
            }

            // There must be a cleaner way to do this! However, I didn't find it. :-(
            for (int i = 0; i < Grid.RowCount; i++)
            {
                DataGridViewComboBoxCell box = (DataGridViewComboBoxCell)Grid.Rows[i].Cells["WidthColumn"];
                // Clear and put in standard tape types
                ResetTapeWidths(box);
                // Add custom tape names
                for (int j = 0; j < CustomGrid.RowCount-1; j++)
                {
                    box.Items.Add(CustomGrid.Rows[j].Cells["Name_Column"].Value.ToString());
                }
            }

        }
	}

}
