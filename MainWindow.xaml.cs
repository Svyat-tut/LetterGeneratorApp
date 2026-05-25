using System;
using System.Windows;
using LetterGeneratorApp.Services;

namespace LetterGeneratorApp
{
    public partial class MainWindow : Window
    {
        private readonly DocumentGenerationService _documentService;

        public MainWindow()
        {
            InitializeComponent();
            
            // Убедиться, что шаблон существует и валиден
            string templatePath = System.IO.Path.Combine("Templates", "letter_template.docx");
            TemplateInitializer.EnsureTemplateExists(templatePath);
            
            _documentService = new DocumentGenerationService();
            
            // Установить сегодняшнюю дату по умолчанию
            LetterDatePicker.SelectedDate = DateTime.Now;
        }

        private void GenerateLetterButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Валидация данных
                if (!ValidateInput())
                    return;

                // Собрать данные из формы
                var letterData = new LetterData
                {
                    RecipientName = RecipientNameTextBox.Text,
                    RecipientPost = RecipientPostTextBox.Text,
                    CompanyName = CompanyNameTextBox.Text,
                    LetterSubject = LetterSubjectTextBox.Text,
                    LetterDate = LetterDatePicker.SelectedDate?.ToString("dd.MM.yyyy") ?? DateTime.Now.ToString("dd.MM.yyyy"),
                    LetterBody = LetterBodyTextBox.Text,
                    SenderName = SenderNameTextBox.Text,
                    SenderPost = SenderPostTextBox.Text
                };

                // Сгенерировать документ
                string outputPath = _documentService.GenerateLetterDocument(letterData);

                // Показать успешное сообщение
                StatusTextBlock.Text = $"✓ Письмо успешно создано: {outputPath}";
                StatusTextBlock.Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(76, 175, 80));
            }
            catch (Exception ex)
            {
                StatusTextBlock.Text = $"✗ Ошибка: {ex.Message}";
                StatusTextBlock.Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(244, 67, 54));
            }
        }

        private void ClearFormButton_Click(object sender, RoutedEventArgs e)
        {
            RecipientNameTextBox.Clear();
            RecipientPostTextBox.Clear();
            CompanyNameTextBox.Clear();
            LetterSubjectTextBox.Clear();
            LetterBodyTextBox.Clear();
            SenderNameTextBox.Clear();
            SenderPostTextBox.Clear();
            LetterDatePicker.SelectedDate = DateTime.Now;
            StatusTextBlock.Text = "";
        }

        private bool ValidateInput()
        {
            if (string.IsNullOrWhiteSpace(RecipientNameTextBox.Text))
            {
                StatusTextBlock.Text = "✗ Пожалуйста, заполните поле 'ФИО адресата'";
                StatusTextBlock.Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(244, 67, 54));
                return false;
            }

            if (string.IsNullOrWhiteSpace(LetterBodyTextBox.Text))
            {
                StatusTextBlock.Text = "✗ Пожалуйста, заполните поле 'Текст письма'";
                StatusTextBlock.Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(244, 67, 54));
                return false;
            }

            return true;
        }
    }
}
