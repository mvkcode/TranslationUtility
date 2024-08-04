//Copyright © 2022 ManpowerGroup. All Rights Reserved
using System.Collections.Generic;
using System.Xml.Serialization;

namespace UtilityProject
{
	[XmlRoot(ElementName = "phrase")]
	public class Phrase
	{

		[XmlElement(ElementName = "en")]
		public string En { get; set; }

		[XmlElement(ElementName = "en-CA")]
		public string EnCA { get; set; }

		[XmlElement(ElementName = "fr-CA")]
		public string FRCA { get; set; }

		[XmlElement(ElementName = "en-US")]
		public string EnUS { get; set; }
		//Added New
		[XmlElement(ElementName = "en-AU")]
		public string ENAustralia { get; set; }

		[XmlElement(ElementName = "en-GB")]
		public string ENUnitedKingdom { get; set; }
		//New for French
		[XmlElement(ElementName = "fr-FR")]
		public string FRFrench { get; set; }

		//[XmlElement(ElementName = "en-FR")]
		//public string FREnglish { get; set; }
		//End for French

		//New for Belgium
		//[XmlElement(ElementName = "en-BE")]
		//public string BelgiumEnglish { get; set; }

		[XmlElement(ElementName = "fr-BE")]
		public string BelgiumFrench { get; set; }

		[XmlElement(ElementName = "nl-BE")]
		public string BelgiumDutch { get; set; }
		//End of Belgium

		//New for Germany
		[XmlElement(ElementName = "de-DE")]
		public string GermanyGerman { get; set; }

		[XmlElement(ElementName = "it-IT")]
		public string ItalyItalian { get; set; }

		//[XmlElement(ElementName = "en-DE")]
		//public string GermanyEnglish { get; set; }
		//End of Germany

		//New for Japan
		[XmlElement(ElementName = "ja-JP")]
		public string JapanJapanese { get; set; }

		//[XmlElement(ElementName = "en-JP")]
		//public string JapanEnglish { get; set; }
		//End of Japan

		//New for Netherlands
		[XmlElement(ElementName = "nl-NL")]
		public string NetherlandDutch { get; set; }

		//[XmlElement(ElementName = "en-NL")]
		//public string NetherlandEnglish { get; set; }
		//End of Netherlands

		//New for Switzerland
		[XmlElement(ElementName = "de-CH")]
		public string SwitzerlandGerman { get; set; }

		[XmlElement(ElementName = "fr-CH")]
		public string SwitzerlandFrench { get; set; }

		//[XmlElement(ElementName = "en-CH")]
		//public string SwitzerlandEnglish { get; set; }
		//End of Switzerland

		[XmlElement(ElementName = "es-MX")]
		public string MexicoSpanish { get; set; }

		[XmlElement(ElementName = "pt-BR")]
		public string BrazilPortugese { get; set; }

		[XmlElement(ElementName = "pt-PT")]
		public string PortugalPortugese { get; set; }

		[XmlElement(ElementName = "es-ES")]
		public string SpainSpainish { get; set; }

		[XmlElement(ElementName = "da-DK")]
		public string DenmarkDanish { get; set; }

		[XmlElement(ElementName = "fi-FI")]
		public string FinlandFinnish { get; set; }

		[XmlElement(ElementName = "pl-PL")]
		public string PolandPolish { get; set; }

		[XmlElement(ElementName = "sv-SE")]
		public string SwedenSwedish { get; set; }

		[XmlElement(ElementName = "nb-NO")]
		public string NorwayNorwegian { get; set; }

		[XmlElement(ElementName = "es-AR")]
		public string ArgentinaSpanish { get; set; }

		[XmlElement(ElementName = "ko-KR")]
		public string KoreaKorean { get; set; }

		[XmlAttribute(AttributeName = "path")]
		public string Path { get; set; }

		[XmlAttribute(AttributeName = "key")]
		public string ItemName { get; set; }

		[XmlAttribute(AttributeName = "itemid")]
		public string Itemid { get; set; }

		[XmlAttribute(AttributeName = "fieldid")]
		public string Fieldid { get; set; }

		[XmlAttribute(AttributeName = "updated")]
		public string Updated { get; set; }

		[XmlText]
		public string Text { get; set; }
	}

	[XmlRoot(ElementName = "sitecore")]
	public class Sitecore1
	{

		[XmlElement(ElementName = "phrase")]
		public List<Phrase> Phrase { get; set; }
	}
}

