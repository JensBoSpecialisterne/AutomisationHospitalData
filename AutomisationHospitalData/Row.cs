using System;
using System.Text.RegularExpressions;

namespace AutomisationHospitalData
{
    internal class Row
    {
        public string år = "";
        public string kvartal = "";
        public string hospital = "";
        public string råvarekategori = "";
        public string leverandør = "";
        public string råvare = "";
        public string øko = "";
        public string variant = "";
        public string prisEnhed = "";
        public string prisTotal = "";
        public string kg = "";
        public string kilopris = "";
        public string oprindelse = "";

        public Row(string år = "", string kvartal = "", string hospital = "", string råvarekategori = "", string leverandør = "", string råvare = "", string øko = "", string variant = "", string prisEnhed = "", string prisTotal = "", string kg = "", string oprindelse = "")
        {
            if (float.Parse(prisTotal) <= 0)
                throw new ArgumentException("Price cannot be zero or lower");

            oprindelse = (Regex.Replace(oprindelse, @"[\d-]", string.Empty) + "   ").Substring(0, 3);

            this.år = år;
            this.kvartal = kvartal;
            this.hospital = hospital;
            this.råvarekategori = råvarekategori;
            this.leverandør = leverandør;
            this.råvare = råvare;
            this.øko = øko;
            this.variant = variant;
            this.prisEnhed = prisEnhed;
            this.prisTotal = prisTotal;
            this.oprindelse = oprindelse;

            if (float.Parse(kg) > 0)
            {
                this.kg = kg;
                kilopris = float.Parse(prisTotal) / float.Parse(kg) + "";
            }
            else
            {
                this.kg = "";
                kilopris = "";
            }
        }
    }
}
