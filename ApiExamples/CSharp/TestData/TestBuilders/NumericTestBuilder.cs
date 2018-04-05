using System;

namespace ApiExamples.TestData.TestBuilders
{
    public class NumericTestBuilder
    {
        private int? mValue1;
        private double mValue2;
        private int mValue3;
        private int? mValue4;
        private bool mLogical;
        private DateTime mDate;

        public NumericTestBuilder()
        {
            this.mValue1 = 1;
            this.mValue2 = 1.0;
            this.mValue3 = 1;
            this.mValue4 = 1;
            this.mLogical = false;
            this.mDate = new DateTime(2018,01,01);
        }
    }
}
