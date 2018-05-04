using System;
using System.IO;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using OfficeOpenXml;
using WaterAnalysisTool.Components;
using WaterAnalysisTool.Exceptions;

namespace WaterAnalysisTool.Loader
{
    class DataParser
    {
        #region Attributes
        private DataLoader Loader;
        private StreamReader Input;

        private List<Sample> CalibrationSamples;
        private List<Sample> CalibrationStandards;
        private List<Sample> QualityControlSamples;
        private List<List<Sample>> CertifiedValueSamples; // When adding to this you will need to find the existing list with matching name, may not always exist
        private List<List<Sample>> Samples;

        private ExcelWorksheet Standardsws;

        private const int LEVENSHTEIN_DISTANCE = 3; // Minimum distance needed for the parser to think that sample names are similar enough to group them together
        #endregion

        #region Constructors
        public DataParser (DataLoader loader, StreamReader inf, ExcelWorksheet stdws)
        {
            this.Loader = loader;
            this.Input = inf;
            this.Standardsws = stdws;

            this.CalibrationSamples = new List<Sample>();
            this.CalibrationStandards = new List<Sample>();
            this.QualityControlSamples = new List<Sample>();
            this.CertifiedValueSamples = new List<List<Sample>>();
            this.Samples = new List<List<Sample>>();

            #region Initialize QualityControlSamples and CertifiedValueSamples from CheckStandards.xlsx config file
            int row = 1, col = 3, index = 0;
            int blankCount = 0;

            // Find all element names
            List<String> elementNames = new List<string>();
            while (Standardsws.Cells[3, col].Value != null)
            { 
                elementNames.Add(Standardsws.Cells[3, col].Value.ToString());
                col++;
            }

            col = 3;

            // Find Continuing Calibration Verification (CCV) section
            while (blankCount < 5)
            {
                if (Standardsws.Cells[row, 1].Value != null)
                {
                    if (!Standardsws.Cells[row, 1].Value.ToString().ToLower().Equals("continuing calibration verification (ccv)"))
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

            if (blankCount > 4)
                throw new ConfigurationErrorException("Could not find \"Continuing Calibration Verification (CCV)\" section in CheckStandards.xlsx config file.");

            row++;

            Sample calibstd = new Sample("", "CCV Standard", "", "QC", 0);
            while(Standardsws.Cells[row, col].Value != null)
            {
                calibstd.Elements.Add(new Element(elementNames[index], "mg/L", Double.Parse(Standardsws.Cells[row, col].Value.ToString()), 0.0, 0.0));
                col++;
                index++;
            }

            this.QualityControlSamples.Add(calibstd);

            // Find Check Standards section
            row = 1;
            index = 0;
            blankCount = 0;
            while (blankCount < 5)
            {
                if (Standardsws.Cells[row, 1].Value != null)
                {
                    if (!Standardsws.Cells[row, 1].Value.ToString().ToLower().Equals("check standards"))
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

            if (blankCount > 4)
                throw new ConfigurationErrorException("Could not find \"Check Standards\" section in CheckStandards.xlsx config file.");

            row++;

            while(Standardsws.Cells[row, 1].Value != null)
            {
                col = 3;
                List<Sample> sg = new List<Sample>();
                Sample checkstd = new Sample("", Standardsws.Cells[row, 1].Value.ToString(), "", "QC", 0);
                    
                while(Standardsws.Cells[row, col].Value != null)
                {
                    checkstd.Elements.Add(new Element(elementNames[index], "mg/L", Double.Parse(Standardsws.Cells[row, col].Value.ToString()), 0.0, 0.0));
                    col++;
                }

                row++;
                sg.Add(checkstd);
                this.CertifiedValueSamples.Add(sg);
            }
            #endregion
        }
        #endregion

        #region Public Methods
        public void Parse()
        {
            Sample s;
            this.Input.ReadLine(); // Consume first empty line;

            #region Reading file and Creating Samples
            while(this.Input.Peek() > -1)
            {
                // Parse [Sample Header] 
                String line = this.Input.ReadLine();
                if (!line.Equals("[Sample Header]"))
                    throw new FormatException("Error reading ICP-AES output file. Please make sure the input file has not been changed.");

                String method = this.Input.ReadLine();
                String sname = this.Input.ReadLine();
                this.Input.ReadLine(); // Username
                String comment = this.Input.ReadLine();
                this.Input.ReadLine(); // ID 1
                this.Input.ReadLine(); // ID 2
                this.Input.ReadLine(); // ID 3
                String runtime = this.Input.ReadLine();
                String type = this.Input.ReadLine();
                this.Input.ReadLine(); // Mode
                this.Input.ReadLine(); // CorrFactor
                String repeats = this.Input.ReadLine();

                s = new Sample(method.Split('=')[1], sname.Split('=')[1], comment.Split('=')[1], runtime.Split('=')[1], type.Split('=')[1], Int32.Parse(repeats.Split('=')[1]));
                line = this.Input.ReadLine(); // Consume empty line

                if(!line.Equals(""))
                    throw new FormatException("Error reading ICP-AES output file. Please make sure the input file has not been changed.");

                // Parse [Results]
                line = this.Input.ReadLine();
                if (!line.Equals("[Results]"))
                    throw new FormatException("Error reading ICP-AES output file. Please make sure the input file has not been changed.");

                this.Input.ReadLine(); // Consumes column header line

                while(!(line = this.Input.ReadLine()).Equals("")) // Adding elements
                {
                    String[] elementLine = line.Split(',');
                    if (elementLine.Length < 5)
                        throw new FormatException("Error reading ICP-AES element data at " + elementLine[0] + ". Please make sure the input file has not been changed.");

                    String elem = elementLine[0];
                    String units = elementLine[1];
                    String avg = elementLine[2];
                    String stddev = elementLine[3];
                    String rsd = elementLine[4];

                    // Input cleaning
                    if (units.Equals("ppm"))
                        units = "mg/L";

                    if (avg.Contains("*") || stddev.Contains("-") || rsd.Contains("-")) // Some elements have no data
                    {
                        avg = "NaN";
                        stddev = "NaN";
                        rsd = "NaN";
                    }

                    else
                    {
                        var pattern = @"[a-zA-Z]+\s+"; // Some elements have an 'F' at the start of their data...
                        avg = Regex.Replace(avg, pattern, "");
                        stddev = Regex.Replace(stddev, pattern, "");
                        rsd = Regex.Replace(rsd, pattern, "");

                        avg.Trim();
                        stddev.Trim();
                        rsd.Trim();
                    }

                    Double a;
                    Double sd;
                    Double r;

                    if (avg.Equals("NaN"))
                        a = Double.NaN;

                    else
                        a = Convert.ToDouble(avg);

                    if (stddev.Equals("NaN"))
                        sd = Double.NaN;

                    else
                        sd = Convert.ToDouble(stddev);

                    if (rsd.Equals("NaN"))
                        r = Double.NaN;

                    else
                        r = Convert.ToDouble(rsd);


                     s.Elements.Add(new Element(elem, units, a, sd, r));
                }

                // Consume [Internal Standards]
                this.Input.ReadLine(); // [Internal Standards]
                this.Input.ReadLine(); // Elem,Units,Avg,Stddev,RSD
                this.Input.ReadLine(); // Empty line

                // Add sample to correct list
                switch(s.SampleType)
                {
                    case "Cal":
                        this.CalibrationStandards.Add((Sample)s.Clone());
                        break;

                    case "QC":
                        CheckForQCSampleType(s, false);
                        break;

                    case "Unk":
                        CheckForQCSampleType(s, true);
                        break;

                    default:
                        throw new FormatException("Error: Unexpected sampe type encountered: " + s.SampleType);
                }

                s = null;
            }
            #endregion

            #region Adding Sample Lists to Loader
            this.Loader.AddCalibrationSampleGroup(new SampleGroup(this.CalibrationSamples, "Instrument Blanks", false));
            this.Loader.AddCalibrationStandard(new SampleGroup(this.CalibrationStandards, "Calibration Standards", false));
            this.Loader.AddQualityControlSampleGroup(new SampleGroup(this.QualityControlSamples, "Stated Values", true));

            foreach(List<Sample> l in this.CertifiedValueSamples)
                this.Loader.AddCertifiedValueSampleGroup(new SampleGroup(l, "Certified Values", true));

            foreach (List<Sample> l in this.Samples)
                this.Loader.AddSampleGroup(new SampleGroup(l, l[0].Name, false));
            #endregion
        }
        #endregion

        #region Private Methods
        private void CheckForQCSampleType(Sample s, bool unkown)
        {
            bool flag = false;

            if (s.Name.Contains("Instrument Blank")) // Sample is a quality control blank
                this.CalibrationSamples.Add((Sample)s.Clone());

            else if (s.Name.Contains("CCV")) // Sample is a continuing verification sample
                this.QualityControlSamples.Add((Sample)s.Clone());

            else
            {
                foreach (List<Sample> sg in this.CertifiedValueSamples) // Check if Sample is a certified value sample
                {
                    /*if(Utils.Utils.LevenshteinDistance(s.Name, sg[0].Name) <= LEVENSHTEIN_DISTANCE)
                    {
                        sg.Add((Sample)s.Clone());
                        flag = true;
                    }*/

                    if (s.Name.Contains(sg[0].Name))
                    {
                        sg.Add((Sample)s.Clone());
                        flag = true;
                    }
                }

                if (!flag && !unkown) // Sample is an unkown certified value
                    Console.WriteLine("\tWarning: Encountered a Certified Value sample whose check standards were not present in the CheckStandards.xlsx config file. Sample data will be missing from output file.");

                if(!flag) // Sample is a generic sample
                {
                    flag = false;
                    foreach (List<Sample> sg in this.Samples)
                    {
                        /*if (Utils.Utils.LevenshteinDistance(s.Name, sg[0].Name) <= LEVENSHTEIN_DISTANCE)
                        {
                            sg.Add((Sample)s.Clone());
                            flag = true;
                        }*/

                        if (s.Name.Contains(sg[0].Name.Split(' ')[0]))
                        {
                            sg.Add((Sample)s.Clone());
                            flag = true;
                        }
                    }

                    if (!flag)
                    {
                        List<Sample> newsg = new List<Sample>();
                        newsg.Add((Sample)s.Clone());
                        this.Samples.Add(newsg);
                    }
                }
            }
        }
        #endregion
    }
}
