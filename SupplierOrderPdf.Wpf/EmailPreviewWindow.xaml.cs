using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Input;
using Microsoft.Win32;
using MimeKit;

namespace SupplierOrderPdf;

/// <summary>
/// Окно предпросмотра и редактирования email сообщения перед отправкой.
/// 
/// Основные функции:
/// - Отображение и редактирование полей "Кому", "Тема", "Тело письма"
/// - Управление вложениями (добавление/удаление файлов)
/// - Поддержка Drag & Drop для добавления вложений
/// - Открытие вложений по двойному клику
/// - Валидация email адресов получателей
/// 
/// Используется как модальный диалог перед отправкой заявки на закупку,
/// позволяя пользователю проверить и при необходимости скорректировать
/// параметры отправляемого письма.
/// 
/// Поддерживаемые форматы вложений: любые файлы.
/// Drag & Drop: файлы можно перетаскивать прямо в список вложений.
/// </summary>
public partial class EmailPreviewWindow : Window
{
    private readonly ObservableCollection<string> _attachments;

    public string ToText => ToTextBox.Text.Trim();
    public string SubjectText => SubjectTextBox.Text.Trim();
    public string BodyText => BodyTextBox.Text;

    public IReadOnlyList<string> Attachments => _attachments.ToList();

    public EmailPreviewWindow(string to, string subject, string body, IEnumerable<string> attachments)
    {
        InitializeComponent();

        ToTextBox.Text = to ?? string.Empty;
        SubjectTextBox.Text = subject ?? string.Empty;
        BodyTextBox.Text = body ?? string.Empty;

        _attachments = new ObservableCollection<string>(
            (attachments ?? Enumerable.Empty<string>())
            .Where(f => !string.IsNullOrWhiteSpace(f))
        );
        AttachmentsListBox.ItemsSource = _attachments;
    }

    // ----- Кнопки -----

    private void OkButton_Click(object sender, RoutedEventArgs e)
    {
        if (!ValidateInputs())
            return;

        DialogResult = true;
        Close();
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }

    private void AddAttachmentButton_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFileDialog
        {
            Filter = "Все файлы (*.*)|*.*",
            Multiselect = true
        };
        if (dlg.ShowDialog(this) == true)
        {
            foreach (var f in dlg.FileNames)
            {
                if (!_attachments.Contains(f))
                    _attachments.Add(f);
            }
        }
    }

    private void RemoveAttachmentButton_Click(object sender, RoutedEventArgs e)
    {
        var toRemove = AttachmentsListBox.SelectedItems.Cast<string>().ToList();
        foreach (var f in toRemove)
            _attachments.Remove(f);
    }

    // ----- Drag&Drop и двойной клик по вложению -----

    private void AttachmentsListBox_DragOver(object sender, DragEventArgs e)
    {
        if (e.Data.GetDataPresent(DataFormats.FileDrop))
            e.Effects = DragDropEffects.Copy;
        else
            e.Effects = DragDropEffects.None;

        e.Handled = true;
    }

    private void AttachmentsListBox_Drop(object sender, DragEventArgs e)
    {
        if (!e.Data.GetDataPresent(DataFormats.FileDrop))
            return;

        var files = (string[])e.Data.GetData(DataFormats.FileDrop);
        foreach (var f in files)
        {
            if (File.Exists(f) && !_attachments.Contains(f))
                _attachments.Add(f);
        }
    }

    private void AttachmentsListBox_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        if (AttachmentsListBox.SelectedItem is not string path)
            return;

        if (!File.Exists(path))
            return;

        try
        {
            Process.Start(new ProcessStartInfo(path)
            {
                UseShellExecute = true
            });
        }
        catch
        {
            // тихо игнорируем ошибку
        }
    }

    // ----- Валидация -----

    private bool ValidateInputs()
    {
        if (string.IsNullOrWhiteSpace(ToTextBox.Text))
        {
            MessageBox.Show(this,
                "Укажите адрес(а) получателя.",
                "Письмо",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return false;
        }

        var invalid = SplitEmails(ToTextBox.Text)
            .Where(a => !IsValidEmail(a))
            .ToList();

        if (invalid.Count > 0)
        {
            MessageBox.Show(this,
                "Некорректный e-mail:\n" + string.Join(Environment.NewLine, invalid),
                "Письмо",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return false;
        }

        return true;
    }

    private static IEnumerable<string> SplitEmails(string emails)
    {
        if (string.IsNullOrWhiteSpace(emails))
            yield break;

        var parts = emails.Split(new[] { ';', ',', ' ' },
            StringSplitOptions.RemoveEmptyEntries);

        foreach (var p in parts)
        {
            var s = p.Trim();
            if (s.Length > 0)
                yield return s;
        }
    }

    private static bool IsValidEmail(string email)
    {
        try
        {
            return MailboxAddress.TryParse(email, out _);
        }
        catch
        {
            return false;
        }
    }
}