using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using HtmlAgilityPack;

namespace OOXMLConverter;

public class Converter
{
	private readonly XNamespace _w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

	public void CreateWordFile(string htmlFilePath, string wordFilePath)
	{
		using (WordprocessingDocument doc = WordprocessingDocument.Create(wordFilePath, WordprocessingDocumentType.Document))
		{
			MainDocumentPart mainPart = doc.AddMainDocumentPart();

			Document document = new Document();
			mainPart.Document = document;

			XElement xmlBody = CreateBody(htmlFilePath);

			Body body = new Body(xmlBody.ToString());
			document.Append(body);
		}
	}

	private XElement CreateBody(string htmlFilePath)
	{
		XElement xmlBody = new XElement(_w + "body", new XAttribute(XNamespace.Xmlns + "w", _w));
		
		HtmlDocument html = new HtmlDocument();

		Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
		html.Load(htmlFilePath, Encoding.GetEncoding(1251));

		string targetTags = "//p";
		HtmlNodeCollection tags = html.DocumentNode.SelectNodes(targetTags);

		if (tags != null)
		{
			foreach (HtmlNode tag in tags)
			{
				if (tag.Name == "p")
					xmlBody.Add(AddXmlText(tag));
			}
		}

		return xmlBody;
	}

	private XElement AddXmlText(HtmlNode tag)
	{
		string font = "";
		int size = 24;
		bool isBold = true;
		string italic = "single";
		string align = "both";
		float indent = 0;

		// выделение расположения текста
		HtmlAttribute? alignAttr = tag.Attributes.Where(x => x.Name == "align").FirstOrDefault();
		if (alignAttr != null)
		{
			if (alignAttr.Value == "center")
				align = "center";
        }

		HtmlAttribute? styleAttr = tag.Attributes.Where(x => x.Name == "style").FirstOrDefault();
		if (styleAttr != null)
		{

			if (float.TryParse(Regex.Match(styleAttr.Value, "text-indent: (.*?)cm")
				.Groups[1].Value, out indent))
			{
				indent *= 567;
            }
		}

		foreach (var elem in tag.ChildNodes)
		{
			string outerHtml = elem.OuterHtml;

            // выделение названия шрифта
            font = Regex.Match(outerHtml, "face=\"([^\"]*)\"").Groups[1].Value;
			font = Regex.Replace(font, ", ", ";");

			// выделение размера шрифта
			if (int.TryParse(Regex.Match(outerHtml, "style=\"font-size: ([0-9]*)pt\"")
				.Groups[1].Value, out size))
			{
				size *= 2;
			}

			// проверка жирный ли текст
			if (Regex.IsMatch(outerHtml, "style=\"font-weight: normal\""))
			{
				isBold = false;
			}

			// проверка курсивный ли текст
			if (Regex.IsMatch(outerHtml, "style=\"font-style: normal\"")
				&& !Regex.IsMatch(outerHtml, "<u>"))
			{
				italic = "";
			}
        }

		XElement wp = new XElement(_w + "p",
			new XElement(_w + "pPr",
				new XElement(_w + "ind", 
					new XAttribute(_w + "firstLine", indent)),
				new XElement(_w + "jc", 
					new XAttribute(_w + "val", align))),
				new XElement(_w + "r",
					new XElement(_w + "rPr",
						new XElement(_w + "rFonts",
							new XAttribute(_w + "ascii", font),
							new XAttribute(_w + "hAnsi", font),
							new XAttribute(_w + "cs", font),
							new XAttribute(_w + "eastAsia", font)),
						new XElement(_w + "b", 
							new XAttribute(_w + "val", isBold.ToString().ToLower())),
						new  XElement(_w + "sz", 
							new XAttribute(_w + "val", size)),
						new XElement(_w + "u",
							new XAttribute(_w + "val", italic))),
					new XElement(_w + "t", 
						new XAttribute(XNamespace.Xml + "space", "preserve"), tag.InnerText.Trim())));

		return wp;
	}
}
