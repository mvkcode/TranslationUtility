
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace UtilityProject.Schema
{
    public class TranslationData
    {
        public string Itemid { get; set; }
        public string Path { get; set; }
        public string ItemName { get; set; }
        public string Fieldid { get; set; }
        public string English { get; set; }
        public string CanadaEnglish { get; set; }
        public string CanadaFrench { get; set; }
        public string USEnglish { get; set; }
        //Australia
        public string EnglishAustralia { get; set; }
        //United States
        public string EnglishUnitedKingdom { get; set; }
        //Below 2 properties for French
        public string FRFrench { get; set; }
        // public string FrenchEnglish { get; set; }

        //Below 3 properties for Belgium
        //public string BelgiumEnglish { get; set; }
        public string BelgiumFrench { get; set; }
        public string BelgiumDutch { get; set; }

        //Below 2 properties for Germany
        public string GermanyGerman { get; set; }
        //  public string GermanyEnglish { get; set; }

        public string ItalyItalian { get; set; }

        //Below 2 properties for Japan
        public string JapanJapanese { get; set; }
        // public string JapanEnglish { get; set; }

        //Below 2 properties for Netherlands
        public string NetherlandDutch { get; set; }
        // public string NetherlandEnglish { get; set; }

        //Below 3 for  Switzerland
        public string SwitzerlandGerman { get; set; }
        public string SwitzerlandFrench { get; set; }
        // public string SwitzerlandEnglish { get; set; }

        public string MexicoSpanish { get; set; }

        public string BrazilPortugese { get; set; }
        public string PortugalPortugese { get; set; }

        public string SpainSpainish { get; set; }
        public string DenmarkDanish { get; set; }

        public string FinlandFinnish { get; set; }
        public string PolandPolish { get; set; }
        public string SwedenSwedish { get; set; }
        public string NorwayNorwegian { get; set; }
        public string ArgentinaSpanish { get; set; }
        public string KoreaKorean { get; set; }
    }

}
