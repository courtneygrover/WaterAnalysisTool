using System;
using System.Drawing;
using System.IO;
using System.Collections.Generic;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using WaterAnalysisTool.Components;
using WaterAnalysisTool.Exceptions;

namespace WaterAnalysisTool.Loader
{
    class DataLoader
    {
        #region Attributes
        private SampleGroup CalibrationSamples;             // Quality Control Solutions (Instrument Blanks) -> Sample Type: QC
        private SampleGroup CalibrationStandards;           // Calibration Standard -> Sample Type: Cal
        private SampleGroup QualityControlSamples;          // Stated Values (CCV) -> Sample Type: QC
        private List<SampleGroup> CertifiedValueSamples;    // Certified Values (SoilB/TMDW/etc.) -> Sample Type: QC
        private List<SampleGroup> Samples;                  // Generic Samples -> Sample Type: Unk

        private List<String> Messages;
        private ExcelWorksheet Standardsws;
        private StreamReader Input;
        private ExcelPackage Output;
        #endregion

        #region Constructors
        public DataLoader(StreamReader inf, ExcelPackage outf, String method)
        {
            FileInfo fi = new FileInfo("CheckStandards.xlsx");
            if (!fi.Exists)
                throw new FileNotFoundException("Error: The CheckStandards.xlsx config file does not exist or could not be found and the input file could not be parsed.\t\nPlease refer to the user manual for instructions on how to use the config file.");

            ExcelPackage config = new ExcelPackage(fi);

            foreach (ExcelWorksheet ws in config.Workbook.Worksheets)
            {
                if (ws.Cells[1, 1].Value.ToString().Equals(method))
                    this.Standardsws = ws;
            }

            if (Standardsws == null)
                throw new ConfigurationErrorException("Error: Could not find a configuration sheet that matches the method: " + method);

            this.Input = inf;
            this.Output = outf;
            this.Output.Workbook.Worksheets.Add("Data");
            this.Output.Workbook.Worksheets.Add("Calibration Standards");

            this.CertifiedValueSamples = new List<SampleGroup>();
            this.Samples = new List<SampleGroup>();

            this.Messages = new List<String>();
        }
        #endregion

        #region Public Methods
        public void Load()
        {
            // Load performs the following functions:
            // 1. Write QC Sample, Certified Val Sample, and Sample data into the Data worksheet
            //  1.1 Method Header, Analysis Date, and Sample Name Descriptor in as first three rows
            //  1.2 Bolded Element Name headers (x2 one for mg/L and another for RSD)
            //  1.3 Bolded Units (x2 one for mg/L and another for RSD)
            //  1.4 Write QC 
            // 2. Write Calibration Sample data into the Calibration Standards Worksheet
            // Load expects the package to have all required worksheets

            #region Error Checking
            if (this.Output.Workbook == null)
                throw new ArgumentNullException("Workbook is null.\n");

            if (this.Output.Workbook.Worksheets.Count < 2)
                throw new ArgumentOutOfRangeException("Invalid number of worksheets present in workbook.\n");
            #endregion

            DataParser parser = new DataParser(this, Input, Standardsws);
            parser.Parse();

            var dataws = this.Output.Workbook.Worksheets[1]; // The Data worksheet should be the first worksheet, indeces start at 1.

            // Write header info
            if (Samples.Count > 0)
            {
                Sample headerSample = Samples[Samples.Count - 1].Samples[Samples[Samples.Count - 1].Samples.Count - 1]; // good God
                dataws.Cells["A1"].Value = headerSample.Method;
                dataws.Cells["A2"].Value = headerSample.RunTime.Split(' ')[0];
                dataws.Cells["A2"].Value = Output.Workbook.Properties.Title; // Assumes this was set to like the filename, change later to accept user input for title?

                // Write element header rows
                int col = 3; // Start at: row 5, column C
                foreach (Element e in headerSample.Elements)
                {
                    // Concentration headers
                    dataws.Cells[5, col].Value = e.Name;
                    dataws.Cells[5, col].Style.Font.Bold = true;

                    dataws.Cells[6, col].Value = e.Units;
                    dataws.Cells[6, col].Style.Font.Bold = true;

                    // RSD headers
                    dataws.Cells[5, col + headerSample.Elements.Count + 2].Value = e.Name;
                    dataws.Cells[5, col + headerSample.Elements.Count + 2].Style.Font.Bold = true;

                    dataws.Cells[6, col + headerSample.Elements.Count + 2].Value = "RSD";
                    dataws.Cells[6, col + headerSample.Elements.Count + 2].Style.Font.Bold = true;

                    col++;
                }

                // Freeze top 6 rows and left 2 columns
                dataws.View.FreezePanes(7, 3); // row, col: represents the first row/col that is not frozen

                // Write samples
                int row = 7; // Start at row 7, col 1

                if (CalibrationSamples.Samples.Count > 1) // Don't want to write calibration samples with no data other than the known concentrations
                    row = WriteSamples(dataws, CalibrationSamples, nameof(CalibrationSamples), row);

                if (QualityControlSamples.Samples.Count > 1) // Don't want to QC samples with no data other than the known concentrations
                    row = WriteSamples(dataws, QualityControlSamples, nameof(QualityControlSamples), row);

                foreach (SampleGroup g in CertifiedValueSamples)
                {
                    if (g.Samples.Count > 1)
                        row = WriteSamples(dataws, g, nameof(CertifiedValueSamples), row);
                }

                dataws.Cells[row, 1].Value = "Samples";
                dataws.Cells[row, 1].Style.Font.Bold = true;
                row++;
                foreach (SampleGroup g in Samples)
                {
                    if (Samples.Count > 0)
                    {
                        row = WriteSamples(dataws, g, nameof(Samples), row);
                        row--;
                    }
                }

                // Write calibration standards
                var calibws = this.Output.Workbook.Worksheets[2]; // The calibration worksheet is the second worksheet
                WriteStandards(calibws, CalibrationStandards);

                this.Output.Save();

                this.Messages.Add("Success: Formatted Excel sheet generated.");
            }

            else
                this.Messages.Add("Error: Parser found zero generic samples. Could not generate formatted Excel sheet.");

            foreach (String msg in this.Messages)
                Console.WriteLine(msg);
        } // end Load

        #region Add<Sample>
        public void AddCalibrationSampleGroup(SampleGroup sample)
        {
            if (sample == null)
                throw new ArgumentNullException("SampleGroup (Calibration Sample) is null.\n");

            this.CalibrationSamples = (SampleGroup) sample.Clone();
        }

        public void AddCalibrationStandard(SampleGroup sample)
        {
            if (sample == null)
                throw new ArgumentNullException("SampleGroup (Calibration Standard) is null.\n");

            this.CalibrationStandards = (SampleGroup) sample.Clone();
        }

        public void AddQualityControlSampleGroup(SampleGroup sample)
        {
            if (sample == null)
                throw new ArgumentNullException("SampleGroup (Quality Control) is null.\n");

            this.QualityControlSamples = (SampleGroup) sample.Clone();
        }

        public void AddCertifiedValueSampleGroup(SampleGroup sample)
        {
            if (sample == null)
                throw new ArgumentNullException("SampleGroup (Certified Value) is null.\n");

            this.CertifiedValueSamples.Add((SampleGroup) sample.Clone());
        }

        public void AddSampleGroup(SampleGroup sample)
        {
            if (sample == null)
                throw new ArgumentNullException("SampleGroup (Generic) is null.\n");

            this.Samples.Add((SampleGroup) sample.Clone());
        }
        #endregion
        #endregion

        #region Private Methods
        private int WriteSamples(ExcelWorksheet dataws, SampleGroup samples, String type, int row)
        {
            int count = 0;
            int rowStart, rowEnd, col;
            bool flag = false;
            Sample known;

            // Write sample name header
            switch (type)
            {
                case "CalibrationSamples":
                    dataws.Cells[row, 1].Value = "Quality Control Solutions";
                    break;

                case "QualityControlSamples":
                    dataws.Cells[row, 1].Value = "Stated Values";

                    known = samples.Samples[0];
                    col = 3;

                    foreach(Element e in known.Elements)
                    {
                        if (e.Average != Double.NaN) // Assumes parser set average in elements with no data to Double.Nan
                        {
                            dataws.Cells[row, col].Value = e.Average;
                            dataws.Cells[row, col].Style.Font.Bold = true;
                        }

                        col++;
                    }

                    break;

                case "CertifiedValueSamples":
                    dataws.Cells[row, 1].Value = "Certified Values";

                    known = samples.Samples[0];
                    col = 3;

                    foreach (Element e in known.Elements)
                    {
                        if (e.Average != Double.NaN) // Assumes parse set average in elements with no data to Double.NaN
                        {
                            dataws.Cells[row, col].Value = e.Average;
                            dataws.Cells[row, col].Style.Font.Bold = true;
                        }

                        col++;
                    }

                    break;

                default:
                    dataws.Cells[row, 1].Value = samples.Name.Split(' ')[0];

                    break;
            }

            dataws.Cells[row, 1].Style.Font.Bold = true;

            row++;
            rowStart = row;

            // Write sample data
            foreach (Sample s in samples.Samples)
            {
                col = 1;
                count = 0;

                if (type == "QualityControlSamples" || type == "CertifiedValueSamples") // Skip the first samples in these groups (known concentrations)
                {
                    if(s != samples.Samples[0])
                    {
                        dataws.Cells[row, col].Value = s.Name;
                        dataws.Cells[row, ++col].Value = s.RunTime.Split(' ')[1];

                        foreach (Element e in s.Elements)
                        {
                            count++;

                            if (e.Average != Double.NaN) // Won't bother with cells where data does not exist (assumes parser set average in elements with no data to Double.NaN)
                            {
                                // Write Analyte concentrations
                                dataws.Cells[row, col + 1].Value = e.Average;
                                dataws.Cells[row, col + 1].Style.Numberformat.Format = "0.000";

                                // Write RSD
                                dataws.Cells[row, col + 1 + s.Elements.Count + 2].Value = e.RSD;
                                dataws.Cells[row, col + 1 + s.Elements.Count + 2].Style.Numberformat.Format = "0.000";
                            }

                            col++;
                        }

                        row++;
                    }
                }

                else
                {
                    dataws.Cells[row, col].Value = s.Name;
                    dataws.Cells[row, ++col].Value = s.RunTime.Split(' ')[1];

                    foreach (Element e in s.Elements)
                    {
                        flag = false;
                        count++;

                        if (e.Average != Double.NaN) // Won't bother with cells where data does not exist (assumes parser set average in elements with no data to Double.Nan)
                        {
                            // Write Analyte concentrations
                            dataws.Cells[row, col + 1].Value = e.Average;
                            dataws.Cells[row, col + 1].Style.Numberformat.Format = "0.000";

                            // Write RSD
                            dataws.Cells[row, col + 1 + s.Elements.Count + 2].Value = e.RSD;
                            dataws.Cells[row, col + 1 + s.Elements.Count + 2].Style.Numberformat.Format = "0.000";

                            // Do QA/QC formatting to analyte concentrations
                            #region QA/AC Formatting
                            if (type == "Samples")
                            {

                                // REQ-S3R7, lowest in heirarchy
                                dataws.Cells[row, col + 1].Style.Font.Color.SetColor(Color.Green);

                                // REQ-S3R2, 1st in heirarchy
                                if (e.Average < this.CalibrationSamples.LOD[count - 1])
                                {
                                    dataws.Cells[row, col + 1].Style.Font.Color.SetColor(Color.Firebrick);
                                    flag = true;
                                }

                                // REQ-S3R3, 2nd in heirarchy
                                else if (e.Average < this.CalibrationSamples.LOQ[count - 1] && e.Average > this.CalibrationSamples.LOD[count - 1])
                                {
                                    dataws.Cells[row, col + 1].Style.Font.Color.SetColor(Color.Orange);
                                    flag = true;
                                }

                                // REQ-S3R4, 3rd in heirarchy
                                else if (!flag)
                                {
                                    foreach (SampleGroup g in this.CertifiedValueSamples)
                                        if (g.Average[count - 1] < e.Average + 0.5 && g.Average[count - 1] > e.Average - 0.5)
                                            if (g.Recovery[count - 1] > 110 || g.Recovery[count - 1] < 90)
                                                dataws.Cells[row, col + 1].Style.Font.Color.SetColor(Color.DodgerBlue);
                                }

                                // REQ-S3R5, 4th in heirarchy
                                else if (this.CalibrationSamples.Average[count - 1] > 0.05 * e.Average)
                                {
                                    dataws.Cells[row, col + 1].Style.Font.Color.SetColor(Color.Black);
                                    dataws.Cells[row, col + 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                    dataws.Cells[row, col + 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Firebrick);
                                    flag = true;
                                }

                                // REQ-S3R6, 5th in heirarchy
                                else if (!flag)
                                {
                                    Double highest = 0.0;
                                    int i = 0;

                                    foreach (Sample std in this.CalibrationStandards.Samples)
                                    {
                                        if (std.Elements[i].Average > highest)
                                            highest = std.Elements[i].Average;

                                        i++;
                                    }

                                    if (e.Average > highest)
                                        dataws.Cells[row, col + 1].Style.Font.Color.SetColor(Color.BlueViolet);
                                }
                            }
                            #endregion
                        }

                        col++;
                    }

                    row++;
                }
            }

            rowEnd = row - 1;

            #region Write Unique Rows
            switch (type)
            {
                case "CalibrationSamples":
                    dataws.Cells[row, 1].Value = "average";
                    dataws.Cells[row, 1].Style.Font.Bold = true;

                    for (col = 3; col <= count + 2; col++)
                    {
                        dataws.Cells[row, col].Formula = "AVERAGE(" + dataws.Cells[rowStart, col].Address + ":" + dataws.Cells[rowEnd, col].Address + ")";
                        dataws.Cells[row, col].Style.Font.Bold = true;
                        dataws.Cells[row, col].Style.Numberformat.Format = "0.000";
                    }

                    row++;
                    dataws.Cells[row, 1].Value = "LOD";
                    dataws.Cells[row, 1].Style.Font.Bold = true;

                    for (col = 3; col <= count + 2; col++)
                    {
                        dataws.Cells[row, col].Formula = "3*STDEV(" + dataws.Cells[rowStart, col].Address + ":" + dataws.Cells[rowEnd, col].Address + ")";
                        dataws.Cells[row, col].Style.Font.Bold = true;
                        dataws.Cells[row, col].Style.Numberformat.Format = "0.000";
                    }

                    row++;
                    dataws.Cells[row, 1].Value = "LOQ";
                    dataws.Cells[row, 1].Style.Font.Bold = true;

                    for (col = 3; col <= count + 2; col++)
                    {
                        dataws.Cells[row, col].Formula = "10*STDEV(" + dataws.Cells[rowStart, col].Address + ":" + dataws.Cells[rowEnd, col].Address + ")";
                        dataws.Cells[row, col].Style.Font.Bold = true;
                        dataws.Cells[row, col].Style.Numberformat.Format = "0.000";
                    }

                    break;

                case "QualityControlSamples":
                    dataws.Cells[row, 1].Value = "average";
                    dataws.Cells[row, 1].Style.Font.Bold = true;

                    for (col = 3; col <= count + 2; col++)
                    {
                        dataws.Cells[row, col].Formula = "AVERAGE(" + dataws.Cells[rowStart, col].Address + ":" + dataws.Cells[rowEnd, col].Address + ")";
                        dataws.Cells[row, col].Style.Font.Bold = true;
                        dataws.Cells[row, col].Style.Numberformat.Format = "0.000";
                    }

                    row++;
                    dataws.Cells[row, 1].Value = "% difference";
                    dataws.Cells[row, 1].Style.Font.Bold = true;

                    for (col = 3; col <= count + 2; col++)
                    {
                        dataws.Cells[row, col].Formula = "(" + dataws.Cells[rowEnd + 1, col].Address + "-" + dataws.Cells[rowStart - 1, col].Address + ")/" + dataws.Cells[rowStart - 1, col].Address + "*100";
                        dataws.Cells[row, col].Style.Font.Bold = true;
                        dataws.Cells[row, col].Style.Numberformat.Format = "0";
                    }

                    break;

                case "CertifiedValueSamples":
                    dataws.Cells[row, 1].Value = "average";
                    dataws.Cells[row, 1].Style.Font.Bold = true;

                    for (col = 3; col <= count + 2; col++)
                    {
                        dataws.Cells[row, col].Formula = "AVERAGE(" + dataws.Cells[rowStart, col].Address + ":" + dataws.Cells[rowEnd, col].Address + ")";
                        dataws.Cells[row, col].Style.Font.Bold = true;
                        dataws.Cells[row, col].Style.Numberformat.Format = "0.000";
                    }

                    row++;
                    dataws.Cells[row, 1].Value = "rsd (%)";
                    dataws.Cells[row, 1].Style.Font.Bold = true;

                    int i = 0;
                    for (col = 3; col <= count + 2; col++)
                    {
                        dataws.Cells[row, col].Formula = "STDEV(" + dataws.Cells[rowStart, col].Address + ":" + dataws.Cells[rowEnd, col].Address + ")/" + dataws.Cells[rowEnd + 1, col].Address + "*100";
                        dataws.Cells[row, col].Style.Font.Bold = true;
                        dataws.Cells[row, col].Style.Numberformat.Format = "0";

                        if (samples.RSD[i] > 10)
                            dataws.Cells[row, col].Style.Font.Color.SetColor(Color.Firebrick);

                        i++;
                    }

                    row++;
                    dataws.Cells[row, 1].Value = "recovery (%)";
                    dataws.Cells[row, 1].Style.Font.Bold = true;

                    i = 0;
                    for (col = 3; col <= count + 2; col++)
                    {
                        dataws.Cells[row, col].Formula = dataws.Cells[rowEnd + 1, col].Address + "/" + dataws.Cells[rowStart - 1, col].Address + "*100";
                        dataws.Cells[row, col].Style.Font.Bold = true;
                        dataws.Cells[row, col].Style.Numberformat.Format = "0";

                        if (samples.Recovery[i] < 90 || samples.Recovery[i] > 110)
                            dataws.Cells[row, col].Style.Font.Color.SetColor(Color.Firebrick);

                        i++;
                    }

                    break;

                default:
                    break;

            }
            #endregion

            this.Messages.Add("Success: " + type + "(" + samples.Name + ") written to Excel Worksheet.");

            return  row + 2;
        }// end WriteSamples

        private void WriteStandards(ExcelWorksheet calibws, SampleGroup standards)
        {
            int endRow = 0;
            int numSamples = 0;

            // Write element header rows
            Sample headerSample = standards.Samples[standards.Samples.Count - 1];

            int col = 3; //Start at row 2 col 3

            foreach (Element e in headerSample.Elements)
            {
                // Concentration headers
                calibws.Cells[2, col].Value = e.Name;
                calibws.Cells[2, col].Style.Font.Bold = true;

                calibws.Cells[3, col].Value = e.Units;
                calibws.Cells[3, col].Style.Font.Bold = true;

                col++;
            }

            // Write standards data
            bool found = false;
            int row = 4, elementRow = 2;
            col = 1;
            
            try
            {
                foreach (Sample s in standards.Samples)
                {
                    col = 1;

                    String str;
                    calibws.Cells[row, col].Value = s.Name;
                    calibws.Cells[row, ++col].Value = s.RunTime;

                    col = 3;
                    foreach (Element e in s.Elements)
                    {
                        found = false;
                        for (col = 3; calibws.Cells[elementRow, col].Value != null && !found; col++)
                        {
                            str = calibws.Cells[elementRow, col].Value.ToString();
                            if (e.Name.Equals(str))
                            {
                                found = true;
                                calibws.Cells[row, col].Value = e.Average;
                            }
                        }
                    }

                    row++;
                }

                numSamples = row - 4;
                endRow = row + 2;

                this.Messages.Add("Success: Calibration standards written to Excel Worksheet.");
            }
            
            catch(Exception e)
            {
               this.Messages.Add("Error: Could not write Calibration standards to Excel Worksheet. Reason: " + e.Message); 
            }

            // Calibration Curve
            // 1. Open the CheckStandards.xlsx sheet where the stock solution concentrations can be found and read them in
            //  1.1 Have to worry about not every concentration in the standards list (these will have to be 0's in the .xlsx)
            // 2. Create a graph with the measured counts per second in the standards list over their respective stock solution concentration
       
            // Find Calibration Standards section
            row = 1;
            int blankCount = 0;

            while(blankCount < 5 && blankCount >= 0)
            {
                if(Standardsws.Cells[row, 1].Value != null)
                {
                    if(!Standardsws.Cells[row, 1].Value.ToString().ToLower().Equals("calibration standards"))
                    {
                            row++;
                            blankCount = 0;
                    }
                    else
                        break;
                }

                else
                {
                    blankCount++;
                    row++;
                }
            }

            if(blankCount > 4)
                throw new ConfigurationErrorException("Could not find \"Calibration Standards\" section in CheckStandards.xlsx config file.");

            row++;

            // Find element names and amount of elements
            int elemCol = 3, elemRow = 1;

            while (Standardsws.Cells[elemRow, elemCol].Value == null)
                elemRow++;

            while(Standardsws.Cells[elemRow, elemCol].Value != null)
            {
                calibws.Cells[endRow, elemCol].Value = Standardsws.Cells[elemRow, elemCol].Value;
                calibws.Cells[endRow + 1, elemCol].Value = Standardsws.Cells[elemRow + 1, elemCol].Value;
                calibws.Cells[endRow, elemCol].Style.Font.Bold = true;
                calibws.Cells[endRow + 1, elemCol].Style.Font.Bold = true;
                elemCol++;
            }

            endRow += 2;
            int startRow = endRow;
            int numStandards = 0;
            col = 1;

            for ( ; Standardsws.Cells[row, col].Value != null; row++)
            {
                for(col = 1; col < elemCol; col++)
                    calibws.Cells[endRow, col].Value = Standardsws.Cells[row, col].Value;

                col = 1;
                endRow++;
                numStandards++;
            }

                #region Calibration Curves
                ExcelChart newGraph = null;
                ExcelRange yrange = null, xrange = null;
                ExcelChartSerie serie = null;

                found = false;

                int count = 0, graphCol = 1, graphRow = endRow + 2;

                // Search through Standard element names to match up with Sample element names, and graph them
                for (int sampleElementCol = 3; calibws.Cells[2, sampleElementCol].Value != null; sampleElementCol++)
                {
                    found = false;
                    for (int standardElementCol = 3; standardElementCol < elemCol && !found; standardElementCol++)
                    {
                        //startRow = beginning of standards section
                        if (calibws.Cells[2, sampleElementCol].Value.Equals(calibws.Cells[startRow - 2, standardElementCol].Value))
                        {
                            //you found the matching one, graph it!
                            found = true;

                            yrange = calibws.Cells[4, sampleElementCol, 3 + numSamples, sampleElementCol];
                            xrange = calibws.Cells[startRow, standardElementCol, numStandards + startRow - 1, standardElementCol];

                            newGraph = calibws.Drawings.AddChart(calibws.Cells[2, sampleElementCol].Value.ToString(), eChartType.XYScatter);
                            newGraph.Title.Text = calibws.Cells[2, sampleElementCol].Value.ToString();

                            // This is for output formatting
                            if(count < 5)
                            {
                                newGraph.SetPosition(graphRow, 0, graphCol, 0);
                                graphCol += 5;
                                count++;
                            }
                            else
                            {
                                count = 0;
                                graphCol = 1;
                                graphRow += 17;

                                newGraph.SetPosition(graphRow, 0, graphCol, 0);
                                graphCol += 5;
                                count++;
                            }
                                
                            newGraph.SetSize(300, 250);
                            newGraph.YAxis.MinValue = 0;
                            newGraph.XAxis.MinValue = 0;

                            serie = newGraph.Series.Add(yrange, xrange);
                            ExcelChartTrendline tl = serie.TrendLines.Add(eTrendLine.Linear);
                            tl.DisplayRSquaredValue = false;
                            tl.DisplayEquation = false;
                        }
                    }
                }             

                this.Messages.Add("Success: Calibration curves generated.");
                #endregion
        }// end WriteStandards
        #endregion
    }   
}
