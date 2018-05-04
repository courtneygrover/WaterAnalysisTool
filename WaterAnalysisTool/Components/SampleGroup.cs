using System;
using System.Collections.Generic;

namespace WaterAnalysisTool.Components
{
    class SampleGroup : ICloneable
    {
        #region Attributes
        private String name;
        private bool skipFirst;
        private List<Sample> samples; // first row contains data from Check Standards file
        private List<Double> average;
        private List<Double> lod;
        private List<Double> loq;
        private List<Double> percentDifference;
        private List<Double> rsd;
        private List<Double> recovery;
        #endregion

        #region Properties
        public String Name { get {return this.name;} }

        public List<Double> Average { get { return this.average; } }

        public List<Double> LOQ { get { return this.loq; } }

        public List<Double> LOD { get { return this.lod; } }

        public List<Double> PercentDifference { get { return this.percentDifference; } }

        public List<Double> RSD { get { return this.rsd; } }

        public List<Double> Recovery { get { return this.recovery; } }

        public List<Sample> Samples { get { return this.samples; } }
        #endregion

        #region Constructors
        public SampleGroup(List<Sample> sampleList, String name, bool skipFirst)
        {
            this.name = name;
            this.skipFirst = skipFirst;
            this.samples = new List<Sample>();

            foreach(Sample s in sampleList) // deeeeeep copy yo
                samples.Add(s);

            CalculateAverage();
            CalculateLODandLOQandRSD();
            CalculatePercentDifference();
            CalculateRecovery();
        }
        #endregion

        #region Public Methods
        public Object Clone()
        {
            SampleGroup clone = new SampleGroup(this.samples, this.name, this.skipFirst);
            return clone;
        }
        #endregion

        #region Private Methods
        private void CalculateAverage()
        {
            this.average = new List<Double>();

            int count = 0, index = 0;

            foreach (Element e in this.Samples[0].Elements) // do some initialization
                this.average.Add(0.0);

            bool firstRow = false;

            if (skipFirst)
                firstRow = true;

            foreach (Sample s in this.samples)
            {
                count++;
                index = 0;

                if (!firstRow)
                {
                    foreach (Element e in s.Elements)
                    {
                        this.average[index] += e.Average;
                        index++;
                    }
                }

                firstRow = false;
            }

            if (skipFirst)
                count--;

            for(index = 0; index < this.average.Count; index++)
                this.average[index] = this.average[index] / count;

        }//end CalculateAverage()

        //maybe change this name.....hahaha
        private void CalculateLODandLOQandRSD()
        {
            this.lod = new List<Double>();
            this.loq = new List<Double>();
            this.rsd = new List<Double>();
            
            int count = 0, index = 0;
            bool firstRow = false;

            if (skipFirst)
                firstRow = true;

            foreach (Element e in this.samples[0].Elements)
            {
                this.lod.Add(0.0);
                this.loq.Add(0.0);
                this.rsd.Add(0.0);
            }

            foreach (Sample s in this.samples) // start at row + 1
            {
                count++;
                index = 0;

                if (!firstRow)
                {
                    foreach (Element e in s.Elements)
                    {
                        this.lod[index] += Math.Pow((e.Average - this.average[index]), 2);
                        index++;
                    }
                }

                firstRow = false;
            }

            if (skipFirst)
                count--;

            for (index = 0; index < this.average.Count; index++)
            {
                double sum = this.lod[index];
                this.lod[index] = 3 * Math.Sqrt(sum / (count - 1));
                this.loq[index] = 10 * Math.Sqrt(sum / (count - 1));
                this.rsd[index] = Math.Sqrt(sum / (count - 1)) / this.average[index] * 100;
            }

        }//end CalculateLODandLOQandRSD()

        private void CalculatePercentDifference() // % difference = (mean - certified value) / certified value * 100
        {
            this.percentDifference = new List<Double>();

            for (int x = 0; x < this.average.Count; x++)
            {
                this.percentDifference.Add(0.0);

                if (this.samples[0].Elements[x].Average == -1)
                    this.percentDifference[x] = -1;
                else
                    this.percentDifference[x] = (this.average[x] - this.samples[0].Elements[x].Average) / this.samples[0].Elements[x].Average * 100;
            }

        }//end CalculatePercentDifference()

        private void CalculateRecovery() // %recovery = mean / certified value * 100
        {
            this.recovery = new List<Double>();

            for (int x = 0; x < this.average.Count; x++)
            {
                this.recovery.Add(0.0);

                if (this.samples[0].Elements[x].Average == -1)
                    this.recovery[x] = -1;
                else
                    this.recovery[x] = this.average[x] / this.samples[0].Elements[x].Average * 100;
            }
        }//CalculateRecovery()
        #endregion
        
    }
}
