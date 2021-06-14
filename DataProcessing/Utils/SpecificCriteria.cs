using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataProcessing.Utils
{
    class SpecificCriteria
    {
        public int State { get; set; }
        public string Operand { get; set; }
        public int? Value { get; set; }

        public string GetOperandValue()
        {
            if (Operand == "Below") { return "<="; }
            if (Operand == "Above") { return ">="; }

            throw new Exception("Operand can be either 'Below' or 'Above'!");
        }
    }
}
