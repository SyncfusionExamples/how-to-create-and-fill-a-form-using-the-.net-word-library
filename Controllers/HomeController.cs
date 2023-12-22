using Form.Models;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;

namespace Form.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult CreateForm()
        {
            WordDocument document = new WordDocument();
            IWSection section = document.AddSection();
            section.PageSetup.Margins.All = 50;
            WTextBody textBody = section.Body;

            WTable table = textBody.AddTable() as WTable;
            table.ResetCells(1, 2);
            table.TableFormat.Borders.BorderType = BorderStyle.None;
            table.TableFormat.Borders.Bottom.LineWidth = 3;
            table.TableFormat.Borders.Bottom.Color = Color.Blue;

            WParagraph paragraph = table.Rows[0].Cells[0].AddParagraph() as WParagraph;
            WTextRange text = paragraph.AppendText("Project Status Report") as WTextRange;
            ApplyCharacterFormat(text, 18, true, Color.Blue);

            paragraph = table.Rows[0].Cells[1].AddParagraph() as WParagraph;
            paragraph.ParagraphFormat.BeforeSpacing = 8;
            text = paragraph.AppendText("Overall Status ") as WTextRange;
            ApplyCharacterFormat(text, 12, false, Color.Blue);

            AddDropDown(paragraph, "Status",
                new string[] { "In-Progress", "Testing", "Review", "Completed" });

            paragraph = textBody.AddParagraph() as WParagraph;
            paragraph.ParagraphFormat.AfterSpacing = 24;
            paragraph.ParagraphFormat.BeforeSpacing = 12;
            text = paragraph.AppendText("  Date: ") as WTextRange;
            ApplyCharacterFormat(text, 12, false, Color.Black);

            AddDate(paragraph, "Date");

            table = textBody.AddTable() as WTable;
            table.ResetCells(8, 2);
            table.TableFormat.Borders.BorderType = BorderStyle.None;

            paragraph = table.Rows[0].Cells[0].AddParagraph() as WParagraph;
            paragraph.ParagraphFormat.AfterSpacing = 24;
            text = paragraph.AppendText("Project Name: ") as WTextRange;
            ApplyCharacterFormat(text, 12, true, Color.Black);

            paragraph = table.Rows[0].Cells[1].AddParagraph() as WParagraph;
            AddText(paragraph, "ProjectName");

            paragraph = table.Rows[1].Cells[0].AddParagraph() as WParagraph;
            paragraph.ParagraphFormat.AfterSpacing = 24;
            text = paragraph.AppendText("Team Size: ") as WTextRange;
            ApplyCharacterFormat(text, 12, true, Color.Black);

            paragraph = table.Rows[1].Cells[1].AddParagraph() as WParagraph;
            AddText(paragraph, "TeamSize");

            paragraph = table.Rows[2].Cells[0].AddParagraph() as WParagraph;
            paragraph.ParagraphFormat.AfterSpacing = 24;
            text = paragraph.AppendText("Language: ") as WTextRange;
            ApplyCharacterFormat(text, 12, true, Color.Black);

            paragraph = table.Rows[2].Cells[1].AddParagraph() as WParagraph;
            AddCheckbox(paragraph, "C#");
            AddCheckbox(paragraph, "VB");

            paragraph = table.Rows[3].Cells[0].AddParagraph() as WParagraph;
            paragraph.ParagraphFormat.AfterSpacing = 24;
            text = paragraph.AppendText("Start Date: ") as WTextRange;
            ApplyCharacterFormat(text, 12, true, Color.Black);

            paragraph = table.Rows[3].Cells[1].AddParagraph() as WParagraph;
            AddDate(paragraph, "StartDate");

            paragraph = table.Rows[4].Cells[0].AddParagraph() as WParagraph;
            paragraph.ParagraphFormat.AfterSpacing = 24;
            text = paragraph.AppendText("Project Manager: ") as WTextRange;
            ApplyCharacterFormat(text, 12, true, Color.Black);

            paragraph = table.Rows[4].Cells[1].AddParagraph() as WParagraph;
            AddText(paragraph, "ProjectManager");

            paragraph = table.Rows[5].Cells[0].AddParagraph() as WParagraph;
            paragraph.ParagraphFormat.AfterSpacing = 24;
            text = paragraph.AppendText("Team Name: ") as WTextRange;
            ApplyCharacterFormat(text, 12, true, Color.Black);

            paragraph = table.Rows[5].Cells[1].AddParagraph() as WParagraph;
            AddText(paragraph, "TeamName");

            paragraph = table.Rows[6].Cells[0].AddParagraph() as WParagraph;
            paragraph.ParagraphFormat.AfterSpacing = 24;
            text = paragraph.AppendText("Platform: ") as WTextRange;
            ApplyCharacterFormat(text, 12, true, Color.Black);

            paragraph = table.Rows[6].Cells[1].AddParagraph() as WParagraph;
            AddDropDown(paragraph, "Platform",
                new string[] { "ASP.NET", "ASP.NET MVC", "ASP.NET Core", "Blazor" });

            paragraph = table.Rows[7].Cells[0].AddParagraph() as WParagraph;
            paragraph.ParagraphFormat.AfterSpacing = 24;
            text = paragraph.AppendText("End Date: ") as WTextRange;
            ApplyCharacterFormat(text, 12, true, Color.Black);

            paragraph = table.Rows[7].Cells[1].AddParagraph() as WParagraph;
            AddDate(paragraph, "EndDate");

            section.HeadersFooters.Footer.AddParagraph();
            BlockContentControl contact = section.HeadersFooters.Footer.AddBlockContentControl(ContentControlType.RichText) as BlockContentControl;
            contact.ContentControlProperties.Title = "ContactInformation";
            paragraph = contact.TextBody.AddParagraph() as WParagraph;
            table = contact.TextBody.AddTable() as WTable;
            table.ResetCells(2, 2);
            table.TableFormat.Borders.BorderType = BorderStyle.None;

            paragraph = table.Rows[0].Cells[0].AddParagraph() as WParagraph;
            paragraph.ParagraphFormat.AfterSpacing = 8;
            text = paragraph.AppendText("Contact Information: ") as WTextRange;
            ApplyCharacterFormat(text, 12, true, Color.Black);

            paragraph = table.Rows[1].Cells[0].AddParagraph() as WParagraph;
            text = paragraph.AppendText("Client Project Manager ") as WTextRange;
            ApplyCharacterFormat(text, 12, true, Color.Black);

            paragraph = table.Rows[1].Cells[0].AddParagraph() as WParagraph;
            text = paragraph.AppendText("Mobile: (206) 555-9857-x5467") as WTextRange;
            ApplyCharacterFormat(text, 12, false, Color.Black);

            paragraph = table.Rows[1].Cells[0].AddParagraph() as WParagraph;
            text = paragraph.AppendText("Mail id: janet@xylook.com") as WTextRange;
            ApplyCharacterFormat(text, 12, false, Color.Black);

            MemoryStream stream = new MemoryStream();
            document.Save(stream, FormatType.Docx);
            document.Close();
            return File(stream, "application/docx", "Form.docx");
        }

        public IActionResult FillForm()
        {
            FileStream fileStream = new FileStream(Path.GetFullPath("Data/Form.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            WordDocument document = new WordDocument(fileStream, FormatType.Docx);
            fileStream.Dispose();

            InlineContentControl status = document.FindItemByProperty(EntityType.InlineContentControl,
                "ContentControlProperties.Title", "Status") as InlineContentControl;
            WTextRange textRange = status.ParagraphItems[0] as WTextRange;
            textRange.Text = "In-Progress";
            ApplyCharacterFormat(textRange);

            InlineContentControl date = document.FindItemByProperty(EntityType.InlineContentControl,
                "ContentControlProperties.Title", "Date") as InlineContentControl;
            textRange = date.ParagraphItems[0] as WTextRange;
            textRange.Text = DateTime.Now.ToShortDateString();
            ApplyCharacterFormat(textRange);

            date = document.FindItemByProperty(EntityType.InlineContentControl,
                "ContentControlProperties.Title", "StartDate") as InlineContentControl;
            textRange = date.ParagraphItems[0] as WTextRange;
            textRange.Text = DateTime.Now.AddDays(-6).ToShortDateString();
            ApplyCharacterFormat(textRange);

            date = document.FindItemByProperty(EntityType.InlineContentControl,
                "ContentControlProperties.Title", "EndDate") as InlineContentControl;
            textRange = date.ParagraphItems[0] as WTextRange;
            textRange.Text = DateTime.Now.AddDays(-1).ToShortDateString();
            ApplyCharacterFormat(textRange);

            InlineContentControl inline = document.FindItemByProperty(EntityType.InlineContentControl,
                "ContentControlProperties.Title", "ProjectName") as InlineContentControl;
            inline.ContentControlProperties.LockContents = true;

            textRange = inline.ParagraphItems[0] as WTextRange;
            textRange.Text = "Website for Adventure works cycle";
            ApplyCharacterFormat(textRange);

            inline = document.FindItemByProperty(EntityType.InlineContentControl,
                "ContentControlProperties.Title", "TeamName") as InlineContentControl;
            textRange = inline.ParagraphItems[0] as WTextRange;
            textRange.Text = "Adventure works cycle";
            ApplyCharacterFormat(textRange);

            inline = document.FindItemByProperty(EntityType.InlineContentControl,
                "ContentControlProperties.Title", "TeamSize") as InlineContentControl;
            textRange = inline.ParagraphItems[0] as WTextRange;
            textRange.Text = "10";
            ApplyCharacterFormat(textRange);

            inline = document.FindItemByProperty(EntityType.InlineContentControl,
                "ContentControlProperties.Title", "ProjectManager") as InlineContentControl;
            textRange = inline.ParagraphItems[0] as WTextRange;
            textRange.Text = "Nancy Davolio";
            ApplyCharacterFormat(textRange);

            inline = document.FindItemByProperty(EntityType.InlineContentControl,
                "ContentControlProperties.Title", "C#") as InlineContentControl;
            inline.ContentControlProperties.IsChecked = true;
            inline.ContentControlProperties.LockContentControl = true;

            inline = document.FindItemByProperty(EntityType.InlineContentControl,
                "ContentControlProperties.Title", "Platform") as InlineContentControl;
            inline.ContentControlProperties.LockContentControl = true;
            inline.ContentControlProperties.LockContents = true;

            textRange = inline.ParagraphItems[0] as WTextRange;
            textRange.Text = "ASP.NET";
            ApplyCharacterFormat(textRange);

            MemoryStream stream = new MemoryStream();
            document.Save(stream, FormatType.Docx);
            document.Close();
            return File(stream, "application/docx", "FilledForm.docx");
        }

        private void ApplyCharacterFormat(WTextRange textRange)
        {
            textRange.CharacterFormat.FontName = "Century Gothic";
            textRange.CharacterFormat.FontSize = 12;
            textRange.CharacterFormat.TextColor = Color.Black;
        }
        private void AddCheckbox(WParagraph paragraph, string title)
        {
            InlineContentControl checkbox = paragraph.AppendInlineContentControl(ContentControlType.CheckBox) as InlineContentControl;
            checkbox.ContentControlProperties.Title = title;
            checkbox.ContentControlProperties.Tag = title;
            checkbox.ContentControlProperties.IsChecked = false;
            WTextRange text = paragraph.AppendText(" " + title + " ") as WTextRange;
            ApplyCharacterFormat(text, 12, false, Color.Black);
        }
        private void AddText(WParagraph paragraph, string title)
        {
            InlineContentControl text = paragraph.AppendInlineContentControl(ContentControlType.RichText) as InlineContentControl;
            text.ContentControlProperties.Title = title;
            text.ContentControlProperties.Tag = title;
            WTextRange textRange = new WTextRange(paragraph.Document);
            textRange.Text = "Click or tap to enter text";
            ApplyCharacterFormat(textRange, 12, false, Color.Gray);
            text.ParagraphItems.Add(textRange);
        }
        private void AddDate(WParagraph paragraph, string title)
        {
            InlineContentControl date = paragraph.AppendInlineContentControl(ContentControlType.Date) as InlineContentControl;
            date.ContentControlProperties.Title = title;
            date.ContentControlProperties.Tag = title;
            date.ContentControlProperties.DateDisplayFormat = "M/d/yyyy";
            WTextRange textRange = new WTextRange(paragraph.Document);
            textRange.Text = "Click or tap to enter date";
            ApplyCharacterFormat(textRange, 12, false, Color.Gray);
            date.ParagraphItems.Add(textRange);
        }

        private void AddDropDown(WParagraph paragraph, string title, string[] items)
        {
            InlineContentControl dropdown = paragraph.AppendInlineContentControl(ContentControlType.DropDownList) as InlineContentControl;
            dropdown.ContentControlProperties.Title = title;
            WTextRange textRange = new WTextRange(paragraph.Document);
            textRange.Text = "Choose an item";
            ApplyCharacterFormat(textRange, 12, false, Color.Gray);
            dropdown.ParagraphItems.Add(textRange);
            int i = 1;
            foreach(string itemName in items)
            {
                ContentControlListItem item = new ContentControlListItem();
                item.DisplayText = itemName;
                item.Value = i.ToString();
                i++;
                dropdown.ContentControlProperties.ContentControlListItems.Add(item);
            }
        }

        private void ApplyCharacterFormat(WTextRange textRange, float size, bool isBold, Color color)
        {
            textRange.CharacterFormat.FontName = "Century Gothic";
            textRange.CharacterFormat.FontSize = size;
            textRange.CharacterFormat.Bold = isBold;
            textRange.CharacterFormat.TextColor = color;
        }
        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}