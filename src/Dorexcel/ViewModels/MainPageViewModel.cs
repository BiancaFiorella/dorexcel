using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;

namespace Dorexcel.ViewModels;
public class MainPageViewModel : INotifyPropertyChanged
{ 
    private string documentFilepath = string.Empty;
    public string DocumentFilepath
    {
        get => documentFilepath;
        set
        {
            if (documentFilepath != value)
            {
                documentFilepath = value;
                OnPropertyChanged(nameof(DocumentFilepath));
                IsDocumentLoaded = !string.IsNullOrWhiteSpace(value);
            }
        }
    }

    private bool isDocumentLoaded;
    public bool IsDocumentLoaded
    {
        get => isDocumentLoaded;
        set
        {
            if (isDocumentLoaded != value)
            {
                isDocumentLoaded = value;
                OnPropertyChanged(nameof(IsDocumentLoaded));
            }
        }
    }

    private ObservableCollection<string> sheets = [];
    public ObservableCollection<string> Sheets
    {
        get => sheets;
        set
        {
            if (sheets != value)
            {
                sheets = value;
                OnPropertyChanged(nameof(Sheets));
            }
        }
    }

    private string selectedSheet = string.Empty;
    public string SelectedSheet
    {
        get => selectedSheet;
        set
        {
            if (selectedSheet != value)
            {
                selectedSheet = value;
                OnPropertyChanged(nameof(SelectedSheet));
                IsSheetSelected = !string.IsNullOrWhiteSpace(value);
                IsHeaderRowNumberValid = false;
            }
        }
    }

    private bool isSheetSelected;
    public bool IsSheetSelected
    {
        get => isSheetSelected;
        set
        {
            if (isSheetSelected != value)
            {
                isSheetSelected = value;
                OnPropertyChanged(nameof(IsSheetSelected));
            }
        }
    }

    private int headerRowNumber = 1;
    public int HeaderRowNumber
    {
        get => headerRowNumber;
        set
        {
            if (headerRowNumber != value)
            {
                headerRowNumber = value;
                OnPropertyChanged(nameof(HeaderRowNumber));                
            }
        }
    }

    private bool isHeaderRowNumberValid = false;
    public bool IsHeaderRowNumberValid
    {
        get => isHeaderRowNumberValid;
        set
        {
            if (isHeaderRowNumberValid != value)
            {
                isHeaderRowNumberValid = value;
                OnPropertyChanged(nameof(IsHeaderRowNumberValid));
            }
        }
    }

    private ObservableCollection<string> columns = [];
    public ObservableCollection<string> Columns
    {
        get => columns;
        set
        {
            if (columns != value)
            {
                columns = value;
                OnPropertyChanged(nameof(Columns));
            }
        }
    }

    public Dictionary<string, SortedSet<string>> ColumnValues { get; set; } = [];

    private ObservableCollection<string> nonIdColumns = [];
    public ObservableCollection<string> NonIdColumns
    {
        get => nonIdColumns;
        set
        {
            if (nonIdColumns != value)
            {
                nonIdColumns = value;
                OnPropertyChanged(nameof(NonIdColumns));
            }
        }
    }

    private string selectedIdColumn = string.Empty;
    public string SelectedIdColumn
    {
        get => selectedIdColumn;
        set
        {
            if (selectedIdColumn != value)
            {
                selectedIdColumn = value;
                OnPropertyChanged(nameof(SelectedIdColumn));
                IsIdColumnSelected = !string.IsNullOrWhiteSpace(value);                
            }
        }
    }

    private bool isIdColumnSelected;
    public bool IsIdColumnSelected
    {
        get => isIdColumnSelected;
        set
        {
            if (isIdColumnSelected != value)
            {
                isIdColumnSelected = value;
                OnPropertyChanged(nameof(IsIdColumnSelected));

                if (isIdColumnSelected)
                {
                    NonIdColumns = new ObservableCollection<string>(Columns.Where(c => c != SelectedIdColumn));
                }
            }
        }
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    private void OnPropertyChanged(string propertyName)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}
