using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace SpireHtmlBug
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Spire.Doc.License.LicenseProvider.SetLicenseFileName("license.elic.doc.xml");
            Spire.Doc.License.LicenseProvider.LoadLicense();

            var htmlContent = "Pommes Frites (Pommes frites (Kartoffeln, Rapsöl), HO Sonnenblumenöl, Speisesalz), Kalbfleisch (Kalbfleisch (Schweiz).), Gemüsemischung (Karotten, Blumenkohl, Bohnen grün, Zwiebeln, <strong>Butter</strong>, Petersilie glatt (mit Antioxidationsmittel: Milchsäure)), Panade-Mischung (Paniermehl (Hart<strong>weizen</strong>mehl, <strong>Weizen</strong>mehl, Hefe, Kochsalz jodiert.), <strong>Vollei </strong>flüssig, pasteurisiert, Bodenhaltung; Säuerungsmittel E330.), Mehlmischung (<strong>Weizen</strong>mehl, <strong>Hartweizendunst</strong>, <strong>Weizengluten</strong>, <strong>Gerstenmalzmehl</strong>)), Pflanzenöl (HO Sonnenblumenöl), Zitronen (Zitronen.)";

            var document = new Document("TestHtml.docx");
            var bookmarkNavigator = new BookmarksNavigator(document);
            bookmarkNavigator.MoveToBookmark("HTMLContent");

            if (bookmarkNavigator.CurrentBookmark == null)
            {
                return;
            }

            var tempSection = document.AddSection();
            tempSection.AddParagraph().AppendHTML(htmlContent);

            var replacementFirstItem = tempSection.Paragraphs[0].Items.FirstItem as ParagraphBase;
            var replacementLastItem = tempSection.Paragraphs[tempSection.Paragraphs.Count - 1].Items.LastItem as ParagraphBase;
            var selection = new TextBodySelection(replacementFirstItem, replacementLastItem);
            var part = new TextBodyPart(selection);

            bookmarkNavigator.ReplaceBookmarkContent(part);

            document.SaveToFile(@"C:\Temp\TestHtml.docx", FileFormat.Docx);
            document.SaveToFile(@"C:\Temp\TestHtml.pdf", new ToPdfParameterList { IsEmbeddedAllFonts = true });
        }
    }
}