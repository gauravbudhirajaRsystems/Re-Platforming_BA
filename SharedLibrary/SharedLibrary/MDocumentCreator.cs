using Litera.Document.Common;
using Litera.Document.Normalizer;
using System.Xml.Linq;

namespace SharedLibrary
{
    public static class MDocumentCreator
    {
        public static MDocument CreateMDocument(string xmlRepresentation, bool insertions = true, bool deletions = true, string archivePath = null, string customTemplatePath = null)
        {
            var extendedPropertiesGenerator = new ExtendedPropertiesInfoGenerator();
            var themeGenerator = new ThemeInfoGenerator();
            var blockGenerator = new BlockInfoGenerator();
            var styleGenerator = new StyleInfoGenerator();
            var numInfoGenerator = new AbstractNumberingInfoGenerator();
            var numInstanceGenerator = new NumberingInstanceInfoGenerator();
            var settingsGenerator = new SettingsInfoGenerator();
            var bookmarkGenerator = new BookMarkInfoGenerator();
            var headerFooterGenerator = new HeaderFooterGenerator();
            var preConfiguredTemplateGenerator = new PreconfiguredNumberingGenerator();

            var xdoc = XDocument.Parse(xmlRepresentation);

            var doc = new MDocument();

            doc.ExtendedPropertiesInfo = extendedPropertiesGenerator.GenerateExtendedPropertiesInfo(xdoc);
            doc.ThemeInfo = themeGenerator.GenerateThemeInfo(xdoc);
            doc.Settings = settingsGenerator.GenerateSettingsInfo(xdoc);
            doc.Blocks = blockGenerator.GenerateBlockInfo(xdoc, doc, insertions, deletions);
            doc.Styles = styleGenerator.GenerateStyleInfo(xdoc, doc);
            doc.NumberingInstances = numInstanceGenerator.GenerateNumberingInfo(xdoc);
            doc.AbstractNumberings = numInfoGenerator.GenerateNumberingInfo(xdoc, doc);
            doc.Bookmarks = bookmarkGenerator.GenerateBookmarkInfo(xdoc, doc);
            doc.Headers = headerFooterGenerator.GenerateHeaderInfo(xdoc, doc);
            doc.Footers = headerFooterGenerator.GenerateFooterInfo(xdoc, doc);

            doc.PreconfiguredTemplates = preConfiguredTemplateGenerator.GeneratePreconfiguredTemplates(archivePath);
            doc.CustomTemplates = GenerateCustomTemplates(customTemplatePath);

            doc.DocumentOptions.Add(DocumentOptions.CharacterStyleSuffix, " Char");	//add in the default value. this string will get localized outside of test scenarios
            doc.DocumentOptions.Add(DocumentOptions.LocalizedListSeparator, ",");	//default list separator
            doc.DocumentOptions.Add(DocumentOptions.LCID, "1033");					//default language ID
            doc.DocumentOptions.Add(DocumentOptions.FileName, "");
            doc.DocumentOptions.Add(DocumentOptions.FilePath, "");
            doc.DefaultSectPr = settingsGenerator.GenerateSectPrInfo(xdoc);

            return doc;
        }

        private static List<string> GenerateCustomTemplates(string customTemplatePath)
        {
            var validExtensions = new List<string> { ".docx", ".docm", ".doc", ".dotx", ".dotm", ".dot" };
            var retVal = new List<string>();

            if (!string.IsNullOrEmpty(customTemplatePath))
            {
                if (File.Exists(customTemplatePath))
                {
                    if (validExtensions.Contains(new FileInfo(customTemplatePath).Extension.ToLower()))
                        retVal.Add(customTemplatePath);
                }
                else if (Directory.Exists(customTemplatePath))
                {
                    foreach (var filePath in Directory.GetFiles(customTemplatePath))
                    {
                        //check for valid extension and file not to be hidden which means locked or leftover by Word
                        var fileInfo = new FileInfo(filePath);
                        if (validExtensions.Contains(fileInfo.Extension.ToLower()) && (fileInfo.Attributes & FileAttributes.Hidden) != FileAttributes.Hidden)
                            retVal.Add(filePath);
                    }
                }
            }

            return retVal;
        }
    }
}
