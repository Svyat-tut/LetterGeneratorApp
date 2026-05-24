using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace LetterGeneratorApp.Services
{
    public class DocumentGenerationService
    {
        private const string TemplateFileName = "Templates/letter_template.docx";
        private const string OutputDirectory = "GeneratedLetters";

        public DocumentGenerationService()
        {
            // Создать директорию для сохранения писем, если её нет
            if (!Directory.Exists(OutputDirectory))
            {
                Directory.CreateDirectory(OutputDirectory);
            }
        }

        public string GenerateLetterDocument(LetterData letterData)
        {
            if (letterData == null)
                throw new ArgumentNullException(nameof(letterData));

            // Проверить существование шаблона
            if (!File.Exists(TemplateFileName))
                throw new FileNotFoundException($"Шаблон документа не найден: {TemplateFileName}");

            // Создать имя выходного файла
            string outputFileName = GenerateOutputFileName(letterData.RecipientName);
            string outputPath = Path.Combine(OutputDirectory, outputFileName);

            // Копировать шаблон в новый файл
            File.Copy(TemplateFileName, outputPath, overwrite: true);

            // Открыть документ и заменить плейсхолдеры
            using (WordprocessingDocument doc = WordprocessingDocument.Open(outputPath, isEditable: true))
            {
                // Заменить в основном теле документа
                ReplaceTextInBody(doc.MainDocumentPart!.Document.Body!, letterData);

                // Заменить в колонтитулах (если они есть)
                ReplaceTextInHeadersFooters(doc, letterData);

                // Сохранить изменения
                doc.MainDocumentPart.Document.Save();
            }

            return Path.GetFullPath(outputPath);
        }

        private void ReplaceTextInBody(Body body, LetterData letterData)
        {
            var replacements = CreateReplacements(letterData);

            // Пройтись по всем параграфам в теле документа
            foreach (var paragraph in body.Descendants<Paragraph>())
            {
                ReplacePlaceholdersInParagraph(paragraph, replacements);
            }

            // Пройтись по всем таблицам
            foreach (var table in body.Descendants<Table>())
            {
                foreach (var paragraph in table.Descendants<Paragraph>())
                {
                    ReplacePlaceholdersInParagraph(paragraph, replacements);
                }
            }
        }

        private void ReplaceTextInHeadersFooters(WordprocessingDocument doc, LetterData letterData)
        {
            var replacements = CreateReplacements(letterData);

            // Заменить в колонтитулах
            foreach (var headerPart in doc.MainDocumentPart!.HeaderParts)
            {
                foreach (var paragraph in headerPart.Header.Descendants<Paragraph>())
                {
                    ReplacePlaceholdersInParagraph(paragraph, replacements);
                }
            }

            foreach (var footerPart in doc.MainDocumentPart.FooterParts)
            {
                foreach (var paragraph in footerPart.Footer.Descendants<Paragraph>())
                {
                    ReplacePlaceholdersInParagraph(paragraph, replacements);
                }
            }
        }

        private void ReplacePlaceholdersInParagraph(Paragraph paragraph, Dictionary<string, string> replacements)
        {
            // Получить весь текст параграфа
            string fullText = GetParagraphText(paragraph);

            // Проверить, есть ли плейсхолдеры
            bool hasPlaceholders = replacements.Keys.Any(key => fullText.Contains(key));

            if (!hasPlaceholders)
                return;

            // Заменить плейсхолдеры в текстовых элементах
            foreach (var textRun in paragraph.Descendants<Text>())
            {
                string originalText = textRun.Text;
                string modifiedText = originalText;

                foreach (var replacement in replacements)
                {
                    modifiedText = modifiedText.Replace(replacement.Key, replacement.Value);
                }

                if (originalText != modifiedText)
                {
                    textRun.Text = modifiedText;
                }
            }

            // Если плейсхолдеры разделены между несколькими Run элементами, объединить их
            MergeRunsWithPlaceholders(paragraph, replacements);
        }

        private void MergeRunsWithPlaceholders(Paragraph paragraph, Dictionary<string, string> replacements)
        {
            // Получить все текстовые элементы в параграфе
            var textElements = paragraph.Descendants<Text>().ToList();

            // Объединить все текстовые элементы в одно целое
            string combinedText = string.Concat(textElements.Select(t => t.Text));

            // Проверить наличие плейсхолдеров
            bool hasPlaceholders = replacements.Keys.Any(key => combinedText.Contains(key));

            if (!hasPlaceholders)
                return;

            // Заменить плейсхолдеры в объединённом тексте
            foreach (var replacement in replacements)
            {
                combinedText = combinedText.Replace(replacement.Key, replacement.Value);
            }

            // Очистить все текстовые элементы
            foreach (var text in textElements)
            {
                text.Text = "";
            }

            // Вставить результат в первый текстовый элемент
            if (textElements.Count > 0)
            {
                textElements[0].Text = combinedText;
            }
            else
            {
                // Если нет текстовых элементов, создать новый Run
                var run = new Run();
                run.AppendChild(new Text { Text = combinedText });
                paragraph.AppendChild(run);
            }
        }

        private string GetParagraphText(Paragraph paragraph)
        {
            return string.Concat(paragraph.Descendants<Text>().Select(t => t.Text));
        }

        private Dictionary<string, string> CreateReplacements(LetterData letterData)
        {
            return new Dictionary<string, string>
            {
                { "{RECIPIENT_NAME}", letterData.RecipientName },
                { "{RECIPIENT_POST}", letterData.RecipientPost },
                { "{COMPANY_NAME}", letterData.CompanyName },
                { "{LETTER_SUBJECT}", letterData.LetterSubject },
                { "{LETTER_DATE}", letterData.LetterDate },
                { "{LETTER_BODY}", letterData.LetterBody },
                { "{SENDER_NAME}", letterData.SenderName },
                { "{SENDER_POST}", letterData.SenderPost }
            };
        }

        private string GenerateOutputFileName(string recipientName)
        {
            // Создать безопасное имя файла на основе имени адресата и времени
            string safeName = Regex.Replace(recipientName, @"[<>:""/\\|?*]", "_");
            string timestamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
            return $"Letter_{safeName}_{timestamp}.docx";
        }
    }
}
