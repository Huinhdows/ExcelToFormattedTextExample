using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DriverLicenseProcessing
{
    class DriverRecord
    {
  
        // private variables
        private string commonTransportationCode = "XXXXXX";
        private string commonSubAccountNumber = "XXXXXX";
        private string commonPGPrefix = "XXXXXX";
        private string commonAddress = "XXXXXXXXXXXXXXXXXX";
        private string commonDispositionCode = "XXXXXX";
        private string commonFiller = "XXXXX";
        private char commonPurposeCode = 'X';
        private string commonAccountNumber = "XXXXX";
        private string commonSecurityCode = "XXXXX";
        private string commonDeptCode = "XXXXX";

        // public properties 
        public string TransportationCode { get { return commonTransportationCode; } }
        public string License { get; set; }
        public string LastNameAbbreviation { get { return LastName.Substring(0, 2); } }
        public string SubAccountNumber { get { return commonSubAccountNumber; } }
        public string PGPrefix { get { return commonPGPrefix; } }
        public string Department { get; set; }
        public string Payroll { get; set; }
        public string LastName { get; set; }
        public string Address { get { return commonAddress; } }
        public string DispositionCode { get { return commonDispositionCode; } }
        public string Filler { get { return commonFiller; } }
        public char PurposeCode { get { return commonPurposeCode; }}
        public string AccountNumber { get { return commonAccountNumber; } }
        public string SecurityCode { get { return commonSecurityCode; } }
        public string DeptCode { get { return commonDeptCode; } }

    }
}
