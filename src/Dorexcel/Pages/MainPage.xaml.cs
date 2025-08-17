using Dorexcel.ViewModels;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Input;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Windows.Storage.Pickers;
using WinRT.Interop;

namespace Dorexcel.Pages;
public sealed partial class MainPage : Page
{
    public MainPageViewModel ViewModel { get; } = new();

    private ExcelPackage? Excel;
    private ExcelWorksheet SelectedSheet => Excel!.Workbook.Worksheets[ViewModel.SelectedSheet];

    private string? GetCellValue(int row, int column)
    {
        var cell = SelectedSheet.Cells[row, column];
        return cell?.Value?.ToString();
    }

    private void SetCellValue(int row, int column, string value)
    {
        var cell = SelectedSheet.Cells[row, column];
        cell.Value = value;
    }

    private string GetEntryId() => IdColumnTextBox.Text.Trim();

    private int FindEntryIdRowOrNextAvailable()
    {
        var selectedIdColumnIndex = ViewModel.Columns.IndexOf(ViewModel.SelectedIdColumn) + 1;
        var row = ViewModel.HeaderRowNumber;

        string? cellContent;
        var entryId = GetEntryId();
        do
        {
            row += 1;
            cellContent = GetCellValue(row, selectedIdColumnIndex);
            if (!string.IsNullOrWhiteSpace(cellContent) &&
                cellContent.Trim().Equals(entryId, StringComparison.InvariantCultureIgnoreCase))
            {
                return row;
            }
        }
        while (!string.IsNullOrWhiteSpace(cellContent));

        return row;
    }

    private SortedSet<string> GetUniqueSortedColumnValues(int columnIndex)
    {
        var values = new SortedSet<string>(StringComparer.InvariantCultureIgnoreCase);

        var row = ViewModel.HeaderRowNumber + 1;
        string? cellContent;

        while (true)
        {
            cellContent = GetCellValue(row, columnIndex);
            if (string.IsNullOrWhiteSpace(cellContent)) break;
            values.Add(cellContent.Trim());
            row += 1;
        }

        return values;
    }

    private async Task<ContentDialogResult> ShowDialogSafeAsync(ContentDialog dialog)
    {
        try
        {
            dialog.XamlRoot = XamlRoot;
            return await dialog.ShowAsync();
        }
        catch
        {
            return ContentDialogResult.None;
        }
    }

    private void RegenerateSuggestions()
    {
        foreach (var columnValues in ViewModel.ColumnValues.Values)
        {
            columnValues.Clear();
        }

        ViewModel.ColumnValues.Clear();

        foreach (var column in ViewModel.Columns)
        {
            var uniqueSortedValues = GetUniqueSortedColumnValues(ViewModel.Columns.IndexOf(column) + 1);
            ViewModel.ColumnValues.Add(column, uniqueSortedValues);
        }        
    }


    public MainPage()
    {
        InitializeComponent();
    }

    private void LoadSpreadsheet()
    {
        try
        {
            Excel = new ExcelPackage(ViewModel.DocumentFilepath);

            ViewModel.Sheets.Clear();

            foreach (var worksheet in Excel.Workbook.Worksheets)
            {
                ViewModel.Sheets.Add(worksheet.Name);
            }
        }
        catch (Exception)
        {

        }
    }

    private async Task FindSpreadsheetFileAsync()
    {
        try
        {
            FileOpenPicker filePicker = new()
            {
                ViewMode = PickerViewMode.Thumbnail,
                SuggestedStartLocation = PickerLocationId.Desktop,
            };

            filePicker.FileTypeFilter.Add(".xlsx");

            InitializeWithWindow.Initialize(filePicker, WindowNative.GetWindowHandle(App.Window));

            var file = await filePicker.PickSingleFileAsync();

            if (file == null) return;

            if (Path.Exists(file.Path))
            {
                ViewModel.DocumentFilepath = file.Path;
                LoadSpreadsheet();
            }
        }
        catch (Exception)
        {
            // nothing for now...
        }
    }

    public async void OnOpenSpreadsheetButtonClick(object sender, RoutedEventArgs e)
    {
        var senderButton = sender as Button;

        if (senderButton == null) return;

        senderButton.IsEnabled = false;

        await FindSpreadsheetFileAsync();

        senderButton.IsEnabled = true;
    }

    public async void OnDefineHeaderRowNumberButtonClick(object sender, RoutedEventArgs e)
    {
        ViewModel.Columns.Clear();

        ViewModel.IsHeaderRowNumberValid = false;

        if (ViewModel.HeaderRowNumber < 1)
        {
            await ShowDialogSafeAsync(new()
            {
                Title = "Error",
                Content = "El numero de fila del encabezado debe ser mayor a 0.",
                CloseButtonText = "Aceptar",
            });

            return;
        }

        while (true)
        {
            var columnIndex = ViewModel.Columns.Count + 1;
            var columnName = GetCellValue(ViewModel.HeaderRowNumber, columnIndex);

            if (string.IsNullOrWhiteSpace(columnName)) break;
       
            ViewModel.Columns.Add(columnName);
        }

        RegenerateSuggestions();

        ViewModel.IsHeaderRowNumberValid = ViewModel.Columns.Any();

        if (!ViewModel.IsHeaderRowNumberValid)
        {
            await ShowDialogSafeAsync(new()
            {
                Title = "Error",
                Content = "No se encontraron encabezados en la fila definida",
                CloseButtonText = "Aceptar",
            });
        }
        else
        {
            await FocusManager.TryFocusAsync(IdColumnTextBox, FocusState.Programmatic);
        }
    }

    private async void OnSaveButtonClick(object sender, RoutedEventArgs e)
    {
        var row = FindEntryIdRowOrNextAvailable();
        var entryId = GetEntryId();

        foreach (var column in ViewModel.Columns)
        {
            var isIdColumn = column == ViewModel.SelectedIdColumn;

            if (isIdColumn)
            {
                SetCellValue(row, ViewModel.Columns.IndexOf(column) + 1, entryId);
            }
            else
            {
                var textBoxForColumn = (AutoSuggestBox)FieldsRepeater.TryGetElement(ViewModel.NonIdColumns.IndexOf(column));
                SetCellValue(row, ViewModel.Columns.IndexOf(column) + 1, textBoxForColumn!.Text.Trim());
            }
        }

        try
        {
            await Excel!.SaveAsync();

            RegenerateSuggestions();

            await ShowDialogSafeAsync(new()
            {
                Title = "Informacion",
                Content = "Registro guardado exitosamente.",
                CloseButtonText = "Aceptar",

            });
        }
        catch
        {
            await ShowDialogSafeAsync(new()
            {
                Title = "Error",
                Content = "No se pudo guardar el registro.\nSi tiene la hoja de calculo abierta, cierrela e intente de nuevo.",
                CloseButtonText = "Aceptar",
            });
        }

    }

    private async Task FindForEntryAsync()
    {
        var entryId = GetEntryId();

        if (!string.IsNullOrWhiteSpace(entryId))
        {
            var row = FindEntryIdRowOrNextAvailable();
            var cellForId = GetCellValue(row, ViewModel.Columns.IndexOf(ViewModel.SelectedIdColumn) + 1);

            if (!string.IsNullOrWhiteSpace(cellForId))
            {
                foreach (var column in ViewModel.Columns)
                {
                    var isIdColumn = column == ViewModel.SelectedIdColumn;

                    if (isIdColumn) continue;

                    var textBoxForColumn = (AutoSuggestBox)FieldsRepeater.TryGetElement(ViewModel.NonIdColumns.IndexOf(column));
                    textBoxForColumn!.Text = GetCellValue(row, ViewModel.Columns.IndexOf(column) + 1);
                }

                return;
            }

            await ShowDialogSafeAsync(new()
            {
                Title = "Informacion",
                Content = $"No se encontro un registro para {ViewModel.SelectedIdColumn} \"{entryId}\".\nPero puede crearlo llenando los campos y tocando el boton de guardar.",
                CloseButtonText = "Aceptar",
            });
        }

        foreach (var column in ViewModel.Columns)
        {
            var isIdColumn = column == ViewModel.SelectedIdColumn;

            if (isIdColumn) continue;

            var textBoxForColumn = (AutoSuggestBox)FieldsRepeater.TryGetElement(ViewModel.NonIdColumns.IndexOf(column));

            if (textBoxForColumn == null) continue;

            textBoxForColumn!.Text = string.Empty;
        }
    }

    private async void OnIdColumnQuerySubmitted(AutoSuggestBox sender, AutoSuggestBoxQuerySubmittedEventArgs args)
    {
        await FindForEntryAsync();
    }

    private async void OnSearchButtonClick(object sender, RoutedEventArgs e)
    {
        await FindForEntryAsync();
    }

    private void OnAutoSuggestBoxTextChanged(AutoSuggestBox sender, AutoSuggestBoxTextChangedEventArgs args)
    {
        if (args.Reason == AutoSuggestionBoxTextChangeReason.UserInput)
        {
            var text = sender.Text.Trim().ToLower();
            var suggestions = ViewModel.ColumnValues[sender.Header.ToString()!];
            sender.ItemsSource = suggestions.Where(s => s.Trim().ToLower().StartsWith(text, StringComparison.InvariantCultureIgnoreCase)).ToArray(); ;
        }
    }

    private async void OnAutoSuggestBoxSuggestionChosen(AutoSuggestBox sender, AutoSuggestBoxSuggestionChosenEventArgs args)
    {
        sender.Text = args.SelectedItem.ToString();

        if(sender.Header.ToString() == ViewModel.SelectedIdColumn)
        {
            await FindForEntryAsync();
        }            
    }
}

