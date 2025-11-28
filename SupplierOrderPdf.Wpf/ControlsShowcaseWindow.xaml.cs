using Microsoft.Web.WebView2.Core;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows;

namespace SupplierOrderPdf.Wpf
{
    /// <summary>
    /// Interaction logic for ControlsShowcaseWindow.xaml
    /// </summary>
    public partial class ControlsShowcaseWindow : Window
    {
        public ControlsShowcaseWindow()
        {
            InitializeComponent();

            // Initialize sample data for DataGrid
            SampleDataGrid.ItemsSource = GetSampleData();

            // Initialize WebView2 with sample content
            Loaded += async (s, e) =>
            {
                try
                {
                    await PreviewWebView.EnsureCoreWebView2Async();
                    PreviewWebView.NavigateToString(GetSampleHtml());
                }
                catch (Exception ex)
                {
                    // Handle WebView2 initialization error
                    MessageBox.Show($"Failed to initialize WebView2: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            };
        }

        private ObservableCollection<SampleItem> GetSampleData()
        {
            return new ObservableCollection<SampleItem>
            {
                new SampleItem { Id = 1, Name = "Item 1", Status = "Active", Date = DateTime.Now.AddDays(-1) },
                new SampleItem { Id = 2, Name = "Item 2", Status = "Inactive", Date = DateTime.Now.AddDays(-2) },
                new SampleItem { Id = 3, Name = "Item 3", Status = "Pending", Date = DateTime.Now.AddDays(-3) },
                new SampleItem { Id = 4, Name = "Item 4", Status = "Active", Date = DateTime.Now.AddDays(-4) },
                new SampleItem { Id = 5, Name = "Item 5", Status = "Active", Date = DateTime.Now.AddDays(-5) }
            };
        }

        private string GetSampleHtml()
        {
            return @"<!DOCTYPE html>
<html lang='en'>
<head>
    <meta charset='UTF-8'>
    <meta name='viewport' content='width=device-width, initial-scale=1.0'>
    <title>Sample Preview</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            margin: 20px;
            background-color: #f5f5f5;
            color: #333;
        }
        h1 {
            color: #007acc;
        }
        p {
            line-height: 1.6;
        }
    </style>
</head>
<body>
    <h1>Sample Preview Content</h1>
    <p>This is a sample HTML content loaded in the WebView2 control for demonstration purposes.</p>
    <p>You can display any HTML content here, such as PDF previews, web pages, or generated documents.</p>
</body>
</html>";
        }
    }

    public class SampleItem : INotifyPropertyChanged
    {
        private int _id;
        private string _name = string.Empty;
        private string _status = string.Empty;
        private DateTime _date;

        public int Id
        {
            get => _id;
            set
            {
                _id = value;
                OnPropertyChanged(nameof(Id));
            }
        }

        public string Name
        {
            get => _name;
            set
            {
                _name = value;
                OnPropertyChanged(nameof(Name));
            }
        }

        public string Status
        {
            get => _status;
            set
            {
                _status = value;
                OnPropertyChanged(nameof(Status));
            }
        }

        public DateTime Date
        {
            get => _date;
            set
            {
                _date = value;
                OnPropertyChanged(nameof(Date));
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}