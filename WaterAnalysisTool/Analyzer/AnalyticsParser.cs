using System;
using System.Collections.Generic;
using WaterAnalysisTool.Components;
using OfficeOpenXml;
using WaterAnalysisTool.Exceptions;

namespace WaterAnalysisTool.Analyzer
{
    class AnalyticsParser
    {
        #region Attributes
        private List<List<Element>> elements; // represents all data for one element within at least SampleGroup
        private AnalyticsLoader loader;
        private ExcelPackage dataWorkbook;
        private ExcelWorksheet dataws;
        private int resetRow;
        private int row;
        private int col;
        private List<String> elementNames;
        #endregion

        #region Constructor
        public AnalyticsParser(ExcelPackage datawb, AnalyticsLoader loader)
        {
            this.dataWorkbook = datawb;
            this.loader = loader;
            this.elements = new List<List<Element>>();
            this.dataws = datawb.Workbook.Worksheets[1];
            fillElementNames();
        }
        #endregion

        #region Public Methods
        public void Parse()
        {
            if (this.dataWorkbook.File.Length < 4 || !this.dataWorkbook.File.Exists)
                throw new ParseErrorException("Data Workbook does not exist or does not have correct worksheets.");

            this.row = 7;
            this.col = 3;
            int blankLineCount = 0;
            ExcelWorksheet dataws = this.dataWorkbook.Workbook.Worksheets[1]; //data worksheet

            /* Loop reads through file until it encounters the Samples section */
            while (blankLineCount < 2 && blankLineCount >= 0)
            {
                if (this.dataws.Cells[this.row, 1].Value != null)
                {
                    if (!this.dataws.Cells[this.row, 1].Value.ToString().ToLower().Equals("samples"))
                    {
                        this.row++;
                        blankLineCount = 0;
                    }
                    else
                        blankLineCount = -1;
                }
                else
                {
                    blankLineCount++;
                    this.row++;
                }
            }

            if (blankLineCount > 1)
            {
                Console.WriteLine("No samples found in file.");
                return;
            }

            /* We have reached the Samples.
               Next line should be the name of SampleGroup, 
               the line after that should be the 
               first sample name within the first SampleGroup.
            */
            this.row++;

            while (!isEndOfWorksheet())
            {
                this.loader.AddSampleName(this.dataws.Cells[this.row, 1].Value.ToString());
                this.row++;
                fillElementList();
                this.loader.AddElements(this.elements);
                this.elements.Clear();
                this.row++;
            }
        }
        #endregion

        #region Private Methods
        private void fillElementList()
        {
            this.resetRow = this.row;
            int colLength = 0;
            bool firstRun = true;

            for (int x = 0; this.dataws.Cells[this.row, this.col].Value != null; x++)
            {
                List<Element> analytes = new List<Element>();

                for (int y = 0; this.dataws.Cells[this.row, this.col].Value != null; y++)
                {
                    analytes.Add(new Element(this.elementNames[x], "", Double.Parse(this.dataws.Cells[this.row, this.col].Value.ToString()), this.row, this.col));
                    this.row++;

                    if (firstRun)
                        colLength++;
                }

                firstRun = false;
                this.row = this.resetRow;
                this.col++;

                // Add to the list that represents the sample list
                this.elements.Add(analytes);
            }

            this.row += colLength; // At blank space after first samplegroup
            this.col = 3;
        }

        private bool isEndOfWorksheet()
        {
            if (this.dataws.Cells[this.row, 1].Value != null)
                return false;

            return true;
        }

        private void fillElementNames()
        {
            int column = 3;
            this.elementNames = new List<String>();
            
            while(this.dataws.Cells[5, column].Value != null)
            {
                this.elementNames.Add(this.dataws.Cells[5, column].Value.ToString());
                column++;
            }
        }
        #endregion

    } // end class AnalyticsParser

} // end namespace WaterAnalysisTool.Analyzer
