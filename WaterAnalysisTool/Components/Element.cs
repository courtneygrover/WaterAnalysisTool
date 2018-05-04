using System;

namespace WaterAnalysisTool.Components
{
    class Element : ICloneable
    {
        #region Attributes
        private String name;
        private String units;
        private double avg;
        private double stddev;
        private double rsd;
        #endregion

        #region Properties
        public String Name
        {
            get { return this.name; }
        }

        public String Units
        {
            get { return this.units; }
        }

        public Double Average
        {
            get { return this.avg; }
        }

        public Double StandardDeviation
        {
            get { return this.stddev; }
        }

        public Double RSD
        {
            get { return this.rsd; }
        }
        #endregion

        #region Constructors
        public Element(String name, String units, Double avg, Double stddev, Double rsd)
        {
            this.name = name;
            this.units = units;
            this.avg = avg;
            this.stddev = stddev;
            this.rsd = rsd;
        }
        #endregion

        #region Public Methods
        public Object Clone()
        {
            Element clone = new Element(this.name, this.units, this.avg, this.stddev, this.rsd);
            return clone;
        }
        #endregion
    }
}
