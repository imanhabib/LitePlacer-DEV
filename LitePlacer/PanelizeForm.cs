﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Configuration;
using System.Globalization;

namespace LitePlacer
{
    public partial class PanelizeForm : Form
    {
        public static FormMain MainForm;
        public DataGridView CadData;
        public DataGridView JobData;

        // =================================================================================
        private decimal XFirstOffset = 0;
        private decimal YFirstOffset = 0;

        private int XRepeats = 0;
        private int YRepeats = 0;

        private decimal XIncrement = 0;
        private decimal YIncrement = 0;
        // =================================================================================
        
        public PanelizeForm(FormMain MainF)
        {
            MainForm = MainF;
            InitializeComponent();
        }

        private void PanelizeForm_Load(object sender, EventArgs e)
        {
            string path = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None).FilePath;
            int i = path.LastIndexOf('\\');
            path = path.Remove(i + 1);
            MainForm.LoadDataGrid(path + "LitePlacer.PanelFids", PanelFiducials_dataGridView);
            XFirstOffset = Properties.Settings.Default.Panel_XFirstOffset;
            XFirstOffset_textBox.Text = XFirstOffset.ToString("0.00", CultureInfo.InvariantCulture);
            YFirstOffset = Properties.Settings.Default.Panel_YFirstOffset;
            YFirstOffset_textBox.Text = YFirstOffset.ToString("0.00", CultureInfo.InvariantCulture);
            XRepeats = Properties.Settings.Default.Panel_XRepeats;
            XRepeats_textBox.Text = XRepeats.ToString();
            YRepeats = Properties.Settings.Default.Panel_YRepeats;
            YRepeats_textBox.Text = YRepeats.ToString();
            XIncrement = Properties.Settings.Default.Panel_XIncrement;
            XIncrement_textBox.Text = XIncrement.ToString("0.00", CultureInfo.InvariantCulture);
            YIncrement = Properties.Settings.Default.Panel_YIncrement;
            YIncrement_textBox.Text = YIncrement.ToString("0.00", CultureInfo.InvariantCulture);
            UseBoardFids_checkBox.Checked = Properties.Settings.Default.Panel_UseBoardFids;
        }

        public bool OK = false;

        private void Cancel_button_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void OK_button_Click(object sender, EventArgs e)
        {
            if (!Panelize())
            {
                return;
            }
            OK = true;
            string path = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None).FilePath;
            int i = path.LastIndexOf('\\');
            path = path.Remove(i + 1);
            MainForm.SaveDataGrid(path + "LitePlacer.PanelFids", PanelFiducials_dataGridView);
            Properties.Settings.Default.Panel_XFirstOffset = XFirstOffset;
            Properties.Settings.Default.Panel_YFirstOffset = YFirstOffset;
            Properties.Settings.Default.Panel_XRepeats = XRepeats;
            Properties.Settings.Default.Panel_YRepeats = YRepeats;
            Properties.Settings.Default.Panel_XIncrement = XIncrement;
            Properties.Settings.Default.Panel_YIncrement = YIncrement;
            Properties.Settings.Default.Panel_UseBoardFids = UseBoardFids_checkBox.Checked;
            this.Close();
        }

        // =================================================================================
        // ValidateData: Check, that the Panelize process has all the data and the data is good
        // =================================================================================
        private bool ValidateData()
        {
            decimal val;
            int intval;

            if (decimal.TryParse(XFirstOffset_textBox.Text, out val))
            {
                XFirstOffset = val;
            }
            else
            {
                MainForm.ShowMessageBox(
                    "Invalid value in Offset to lower left board, X",
                    "Invalid value",
                    MessageBoxButtons.OK);
                return false;
            }

            if (decimal.TryParse(YFirstOffset_textBox.Text, out val))
            {
                YFirstOffset = val;
            }
            else
            {
                MainForm.ShowMessageBox(
                    "Invalid value in Offset to lower left board, Y",
                    "Invalid value",
                    MessageBoxButtons.OK);
                return false;
            }

            if (int.TryParse(XRepeats_textBox.Text, out intval))
            {
                XRepeats = intval;
                if (intval < 1)
                {
                    MainForm.ShowMessageBox(
                        "Invalid value in X repeats",
                        "Invalid value",
                        MessageBoxButtons.OK);
                    return false;
                }
            }
            else
            {
                MainForm.ShowMessageBox(
                    "Invalid value in X repeats",
                    "Invalid value",
                    MessageBoxButtons.OK);
                return false;
            }

            if (int.TryParse(YRepeats_textBox.Text, out intval))
            {
                YRepeats = intval;
                if (intval < 1)
                {
                    MainForm.ShowMessageBox(
                        "Invalid value in Y repeats",
                        "Invalid value",
                        MessageBoxButtons.OK);
                    return false;
                }
            }
            else
            {
                MainForm.ShowMessageBox(
                    "Invalid value in Y repeats",
                    "Invalid value",
                    MessageBoxButtons.OK);
                return false;
            }

            if (decimal.TryParse(XIncrement_textBox.Text, out val))
            {
                XIncrement = val;
            }
            else
            {
                MainForm.ShowMessageBox(
                    "Invalid value in X increment",
                    "Invalid value",
                    MessageBoxButtons.OK);
                return false;
            }

            if (decimal.TryParse(YIncrement_textBox.Text, out val))
            {
                YIncrement = val;
            }
            else
            {
                MainForm.ShowMessageBox(
                    "Invalid value in Y increment",
                    "Invalid value",
                    MessageBoxButtons.OK);
                return false;
            }

            if (UseBoardFids_checkBox.Checked)
            {
                // Check, that the fiducials are correctly indicated already
                bool found = false;
                foreach (DataGridViewRow Row in JobData.Rows)
                {
                    if (Row.Cells["GroupMethod"].Value.ToString() == "Fiducials")
                    {
                        found = true;
                        break;
                    }
                }
                if (!found)
                {
                MainForm.ShowMessageBox(
                    "Board fiducials not found or indicated",
                    "No fiducials",
                    MessageBoxButtons.OK);
                return false;
                }
            }
            else
            {
                // Check, that the data grid view has good data:
                // Less than three is error
                if (PanelFiducials_dataGridView.RowCount < 3)
                {
                MainForm.ShowMessageBox(
                    "Need data for at least three (four preferred) fiducials",
                    "Not enough fiducials",
                    MessageBoxButtons.OK);
                return false;
                }
                // Three is warning
                if (PanelFiducials_dataGridView.RowCount == 3)
                {
                    DialogResult dialogResult = MainForm.ShowMessageBox(
                        "Only three fiducials, results might be inaccurate. Continue?",
                        "Three fiducials",
                        MessageBoxButtons.OKCancel
                    );
                    if (dialogResult == DialogResult.Cancel)
                    {
                        return false;
                    }
                }
                // more than three is ok
                // X and Y columns need to have good data, too; others if empty, fill:
                bool OK = true;
                DataGridViewRow Row;
                int i=0;
                for (i = 0; i < PanelFiducials_dataGridView.RowCount-1; i++)
                {
                    Row = PanelFiducials_dataGridView.Rows[i];
                    if (Row.Cells["Designator_column"].Value == null)
                    {
                        Row.Cells["Designator_column"].Value = "Fid" + i.ToString();
                    }
                    if (Row.Cells["Footprint_Column"].Value == null)
                    {
                        Row.Cells["Footprint_Column"].Value = "fiducial";
                    }
                    if (Row.Cells["X_column"].Value == null)
                    {
                        OK = false;
                        break;
                    }
                    if (!decimal.TryParse(Row.Cells["X_column"].Value.ToString(), out val))
                    {
                        OK = false;
                        break;
                    }
                    if (Row.Cells["Y_column"].Value == null)
                    {
                        OK = false;
                        break;
                    }
                    if (!decimal.TryParse(Row.Cells["Y_column"].Value.ToString(), out val))
                    {
                        OK = false;
                        break;
                    }
                    if (Row.Cells["Rotation_Column"].Value == null)
                    {
                        Row.Cells["Rotation_Column"].Value = "0.0";
                    }
                }
                if (!OK)
                {
                    MainForm.ShowMessageBox(
                        "Bad data in fiducials table, line " + i.ToString(),
                        "No fiducials",
                        MessageBoxButtons.OK);
                    return false;
                }
            }
            return true;
        }

        // =================================================================================
        // Panelize: Builds multiple copies to CAD data table. 
        // =================================================================================

        // =================================================================================
        // Helper: FindExistingFiducials: 
        // Returns the row number from job data that has the fiducials, -1 if there are none
        private int FindExistingFiducials()
        {
            int FiducialIndex = -1;
            foreach (DataGridViewRow Row in JobData.Rows)
            {
                if (Row.Cells["GroupMethod"].Value.ToString() == "Fiducials")
                {
                    FiducialIndex = Row.Index;
                    break;
                }
            }
            return FiducialIndex;
        }

        // =================================================================================
        private bool Panelize()
        {
            if (!ValidateData())
            {
                return false;
            };

            // Fill CAD data with panelized values:
            // Take a copy of current CAD data grid...
            DataGridView CadData_copy = new DataGridView();
            CadData_copy.AllowUserToAddRows = false;  // this prevents an empty row in the end
            for (int i = 0; i < CadData.Columns.Count; i++)
            {
                CadData_copy.Columns.Add(CadData.Columns[i].Clone() as DataGridViewColumn);
                CadData_copy.Columns[i].HeaderText = CadData.Columns[i].HeaderText;                
            }
            CadData_copy.Name = "CadData_copy";
            MainForm.DataGridViewCopy(CadData, ref CadData_copy, false);

            // ... and clear existing
            CadData.Rows.Clear();

            // For each component in the copy, multiply it:
            string Component = "";
            decimal val;
            bool OK = true;
            foreach (DataGridViewRow Row in CadData_copy.Rows)
            {
                // Panels are (row, column):
                // ...
                // (2,1), (2,2), (2,3) ...
                // (1,1), (1,2), (1,3) ...
                for (int PanelRow = 1; PanelRow <= YRepeats; PanelRow++)
                {
                    for (int PanelColumn = 1; PanelColumn <= XRepeats; PanelColumn++)
                    {
                        CadData.Rows.Add();
                        int Last = CadData.RowCount - 1;
                        // Component:
                        Component = Row.Cells["Component"].Value.ToString() + "_" + PanelRow.ToString() + PanelColumn.ToString();
                        CadData.Rows[Last].Cells["Component"].Value = Component;
                        // Value_Footprint:
                        CadData.Rows[Last].Cells["Value_Footprint"].Value = Row.Cells["Value_Footprint"].Value;
                        // X_nominal:
                        if (!decimal.TryParse(Row.Cells["X_nominal"].Value.ToString(), out val))
                        {
                            OK = false;
                            Component = Row.Cells["Component"].Value.ToString();
                            break;
                        }
                        val = XFirstOffset + (PanelColumn - 1) * XIncrement + val;
                        CadData.Rows[Last].Cells["X_nominal"].Value = val;
                        // Y_nominal:
                        if (!decimal.TryParse(Row.Cells["Y_nominal"].Value.ToString(), out val))
                        {
                            OK = false;
                            Component = Row.Cells["Component"].Value.ToString();
                            break;
                        }
                        val = YFirstOffset + (PanelRow - 1) * YIncrement + val;
                        CadData.Rows[Last].Cells["Y_nominal"].Value = val;
                        // Rotation:
                        CadData.Rows[Last].Cells["Rotation"].Value = Row.Cells["Rotation"].Value;
                        CadData.Rows[Last].Cells["X_Machine"].Value = "Nan";   // will be set later 
                        CadData.Rows[Last].Cells["Y_Machine"].Value = "Nan";
                        CadData.Rows[Last].Cells["Rotation_machine"].Value = "Nan";
                    }
                }
            }
            if (!OK)
            {
                MainForm.DataGridViewCopy(CadData_copy, ref CadData, false);
                MainForm.ShowMessageBox(
                    "Error in " + Component + " data.",
                    "Fiducial data error",
                    MessageBoxButtons.OK);
                return false;

            }

            // Find existing fiducials, we'll need them
            int FiducialIndex = FindExistingFiducials();
            string[] OriginalFiducials = JobData.Rows[FiducialIndex].Cells["ComponentList"].Value.ToString().Split(',');

            // Build Job data:
            if (!UseBoardFids_checkBox.Checked)  // if using user defined fiducials:
            {
                // Remove existing:
              
                if (FiducialIndex>=0)
                {
                    foreach (string CurrentFiducial in OriginalFiducials)
                    {
                        // loop from bottom, so we can modify the collection we are looping through:
                        for (int i =  CadData.RowCount - 1; i >= 0; i--)
			            {
                            if (CadData.Rows[i].Cells["Component"].Value.ToString().Contains(CurrentFiducial))
                            {
                                CadData.Rows.RemoveAt(i);
                            }
			 
			            }
                    }
                }
                // Add user defined fiducials to Cad data
                for (int i = 0; i < PanelFiducials_dataGridView.RowCount-1; i++)
                {
                    CadData.Rows.Add();
                    int Last = CadData.RowCount - 1;
                    DataGridViewRow Row = PanelFiducials_dataGridView.Rows[i];
                    CadData.Rows[Last].Cells["Component"].Value = Row.Cells["Designator_column"].Value.ToString();
                    CadData.Rows[Last].Cells["Value_Footprint"].Value = "Fid mark | " + Row.Cells["Footprint_Column"].Value.ToString();
                    CadData.Rows[Last].Cells["X_nominal"].Value = Row.Cells["X_column"].Value.ToString();
                    CadData.Rows[Last].Cells["Y_nominal"].Value = Row.Cells["Y_column"].Value.ToString();
                    CadData.Rows[Last].Cells["Rotation"].Value = Row.Cells["Rotation_column"].Value.ToString();
                    CadData.Rows[Last].Cells["X_Machine"].Value = "Nan";   // will be set later 
                    CadData.Rows[Last].Cells["Y_Machine"].Value = "Nan";
                    CadData.Rows[Last].Cells["Rotation_machine"].Value = "Nan";                    
                }
                // Build Job data
                MainForm.FillJobData_GridView();
                int dummy;
                MainForm.FindFiducials_m(out dummy);  // don't care of the result, just trying to find fids

                return true;  // and we're done.
            }

            // Here, we are using the fiducials data from the individual subboards.
            // (We know they are indicated already)
            // Build the job:
            MainForm.FillJobData_GridView();
            // Our job now has multiples of each board fiducials, we only want to use four that are furthest apart
            string LowLeft = OriginalFiducials[1];
            double LowLeft_val = 10000000.0;
            string HighLeft = OriginalFiducials[1];
            double HighLeft_val = 0.0;
            string LowRight = OriginalFiducials[1];
            double LowRight_val = 0.0;
            string HighRight = OriginalFiducials[1];
            double HighRight_val = 0.0;
            double X = 0.0;
            double Y = 0.0;
            // For each fiducial in the list,
            foreach (string CurrentFiducial in OriginalFiducials)
            {
                // Find its nominal coordinates:
                // find the fiducial in CAD data.
                foreach (DataGridViewRow Row in CadData.Rows)
                {
                    if (Row.Cells["Component"].Value.ToString().Contains(CurrentFiducial))
                    {
                        // Get its nominal position (value already checked).
                        double.TryParse(Row.Cells["X_nominal"].Value.ToString(), out X);
                        double.TryParse(Row.Cells["Y_nominal"].Value.ToString(), out Y);
                    }
                    if((X+Y)<LowLeft_val)
                    {
                        LowLeft = Row.Cells["Component"].Value.ToString();
                        LowLeft_val = X + Y;
                    }
                    if ((Y-X) > HighLeft_val)
                    {
                        HighLeft = Row.Cells["Component"].Value.ToString();
                        HighLeft_val = Y - X;
                    }
                    if ((X - Y) > LowRight_val)
                    {
                        LowRight = Row.Cells["Component"].Value.ToString();
                        LowRight_val = X - Y;
                    }
                    if ((X + Y) > HighRight_val)
                    {
                        HighRight = Row.Cells["Component"].Value.ToString();
                        HighRight_val = X + Y;
                    }
                }
            }
            if ((LowLeft_val > 1000000.0) || (HighLeft_val < 0.1) || (LowRight_val < 0.1) || (HighRight_val < 0.1))
            {
                // At leaset one of them did not get updated, placement is not going to work
                // Likely, one fiducial on board, one dimensional panel
                MainForm.ShowMessageBox(
                    "Current fiducials don't spread out.",
                    "Fiducials data error",
                    MessageBoxButtons.OK);
                return false;
            }
            JobData.Rows[FiducialIndex].Cells["ComponentList"].Value = LowLeft + "," + HighLeft + "," + LowRight + "," + HighRight;
            JobData.Rows[FiducialIndex].Cells["ComponentCount"].Value = "4";
            return true;
        }

    }
}
