using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using System.Collections.Immutable;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using LicenseContext = OfficeOpenXml.LicenseContext;
using Terminal.Gui;


ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

Dictionary<String, long> fileSizes = new Dictionary<String, long>();
Dictionary<String, ExtensionStats> extensionsStats = new Dictionary<String, ExtensionStats>();

bool CreateExcel(DirectoryInfo directory, int maxDepth, FileInfo saveFile)
{
    using (ExcelPackage ep = new ExcelPackage(saveFile))
    {
        ep.Workbook.Properties.Title = "Lab1";
        ep.Workbook.Properties.Author = "Anna Butowska";
        ep.Workbook.Properties.SetCustomPropertyValue("index", "180109");

        ExcelWorksheet strukturaKatalogu = ep.Workbook.Worksheets.Add("Struktura katalogu");
        ExcelWorksheet statystyki = ep.Workbook.Worksheets.Add("Statystyki");

        PrintDirectories(strukturaKatalogu, 1, 1, maxDepth, directory, maxDepth);

        foreach(var col in strukturaKatalogu.Columns)
        {
            col.AutoFit();
        }

        Print10Biggest(statystyki);

        foreach (var col in statystyki.Columns)
        {
            col.AutoFit();
        }

        AddDiagrams(statystyki);
        try 
        {
            ep.Save(); 
        }
        catch(UnauthorizedAccessException exception)
        {
            return false;
        }
        return true;
        
    }
}

void Print10Biggest(ExcelWorksheet excelWorksheet)
{
    var sorted = fileSizes.OrderByDescending(file => file.Value).ToDictionary(file => file.Key, file => file.Value);
    for(int i=1;i<=10;i++)
    {
        excelWorksheet.Cells[i, 1].Value = sorted.ToArray()[i-1].Key;
        excelWorksheet.Cells[i, 2].Value = sorted.ToArray()[i-1].Value;
    }
}


void AddDiagrams(ExcelWorksheet excelWorksheet)
{
    int row = 0;
    foreach (var extension in extensionsStats)
    {
        row++;
        excelWorksheet.Cells[row, 4].Value = extension.Key;
        excelWorksheet.Cells[row, 5].Value = extension.Value.count;
        excelWorksheet.Cells[row, 6].Value = extension.Value.size;
    }
    ExcelAddress extensionsAdd = new ExcelAddress(1, 4, row, 4);
    ExcelAddress countsAdd = new ExcelAddress(1, 5, row, 5);
    ExcelAddress sizesAdd = new ExcelAddress(1, 6, row, 6);

    var countChart = (excelWorksheet.Drawings.AddChart("PieChart_Count", eChartType.Pie3D) as ExcelPieChart);
    var sizeChart = (excelWorksheet.Drawings.AddChart("PieChart_Size", eChartType.Pie3D) as ExcelPieChart);
    countChart.Title.Text = "Wykres Ilościowy";
    sizeChart.Title.Text = "Wykres Powierzchniowy";
    countChart.SetPosition(10, 5, 0, 5);
    sizeChart.SetPosition(10, 5, 3, 5);
    countChart.SetSize(600, 300);
    sizeChart.SetSize(600, 300);
    var countsSeries = (countChart.Series.Add(countsAdd.Address, extensionsAdd.Address) as ExcelPieChartSerie);
    var sizesSeries = (sizeChart.Series.Add(sizesAdd.Address, extensionsAdd.Address) as ExcelPieChartSerie);
    countChart.DataLabel.ShowPercent = true;
    sizeChart.DataLabel.ShowPercent = true;
}


int PrintDirectories(ExcelWorksheet excelWorksheet, int col, int row, int depth, DirectoryInfo directoryInfo, int maxDepth)
{
    excelWorksheet.Row(row).OutlineLevel = maxDepth - depth + 1;
    excelWorksheet.Cells[row, col].Value = directoryInfo.FullName;
    row++;
    col++;

    foreach (var file in directoryInfo.GetFiles())
    {
        excelWorksheet.Cells[row, col].Value = file.FullName;
        excelWorksheet.Cells[row, col + 1].Value = file.Extension;
        excelWorksheet.Cells[row, col + 2].Value = file.Length;
        excelWorksheet.Cells[row, col + 3].Value = file.Attributes;

        fileSizes.Add(file.FullName, file.Length);

        ExtensionStats extensionStats;
        bool v = extensionsStats.TryGetValue(file.Extension, out extensionStats);
        if (v)
        {
            extensionStats.count++;
            extensionStats.size += file.Length;
            extensionsStats[file.Extension] = extensionStats;
        }
        else
        {
            extensionStats = new ExtensionStats();
            extensionStats.count = 1;
            extensionStats.size = file.Length;
            extensionsStats.Add(file.Extension, extensionStats);
        }

        row++;
    }

    if (depth > 0)
    {
        foreach (var directory in directoryInfo.GetDirectories())
        {
            row = PrintDirectories(excelWorksheet, col, row, depth - 1, directory, maxDepth);
        }
    }

    return row;
}


Application.Init();
var top = Application.Top;

var win = new Window("Pobierz katalog plików - naciśnij Ctrl+Q, aby zamknąć")
{
    X = 0,
    Y = 1, // Leave one row for the toplevel menu

    // By using Dim.Fill(), it will automatically resize without manual intervention
    Width = Dim.Fill(),
    Height = Dim.Fill()
};

top.Add(win);

var savePath = new Label("Ścieżka zapisu: ") { X = 3, Y = 2 };
var path = new Label("Ścieżka do przeszukania: ")
{
    X = Pos.Left(savePath),
    Y = Pos.Top(savePath) + 2
};
var depth = new Label("Głębokość przeszukiwania: ")
{
    X = Pos.Left(path),
    Y = Pos.Top(path) + 2
};
var savePathText = new TextField("")
{
    X = Pos.Right(depth),
    Y = Pos.Top(savePath),
    Width = 40
};
var pathText = new TextField("")
{
    X = Pos.Left(savePathText),
    Y = Pos.Top(path),
    Width = 40
};
var depthText = new TextField("")
{
    X = Pos.Left(savePathText),
    Y = Pos.Top(depth),
    Width = Dim.Width(pathText)
};

var textBox = new Label("")
{
    X = Pos.Left(path),
    Y = Pos.Top(depthText) + 10,
};

void StartExecution()
{
    //new FileInfo(@"D:\Studia\semestr7\JezykiProgramowaniaNaPlatformie.NET\Lab1\ZadDom\lab3.xlsx")

    if (savePathText.Text.IsEmpty || pathText.Text.IsEmpty || depthText.Text.IsEmpty)
    {
        textBox.Text = "Nie podano wszystkich parametrów";
        return;
    }

    string[] parts = savePathText.Text.ToString().Split(".");


    if (!parts[parts.Length - 1].Equals("xlsx"))
    {
        savePathText.Text += ".xlsx";
    }

    FileInfo saveFile = new FileInfo(savePathText.Text.ToString());
    if (saveFile.Exists)
    {
        textBox.Text = "Plik o takiej nazwie już istnieje!";
        return;
    }

    DirectoryInfo directory = new DirectoryInfo(pathText.Text.ToString());
    if (!directory.Exists)
    {
        textBox.Text = "Ten katalog nie istnieje!";
        return;
    }

    FileAttributes dirFileAttr = File.GetAttributes(pathText.Text.ToString());
    if (!((dirFileAttr & FileAttributes.Directory) == FileAttributes.Directory))
    {
        textBox.Text = "To plik, a nie katalog!";
        return;
    }

    bool doNotExecute = false;
    int maxDepth;
    if (!Int32.TryParse(depthText.Text.ToString(), out maxDepth))
    {
        textBox.Text = "Podana glebokosc nie jest liczba!";
        return;
    }
    if(CreateExcel(directory, maxDepth, saveFile))
        textBox.Text = "Pomyślnie utworzono plik!\n" +
            "Znajduje się on w pliku : " + savePathText.Text.ToString();
    else
        textBox.Text = "Odmowa dostępu do pliku : " + savePathText.Text.ToString();

}

var submit = new Button(50, 10, "Submit");

submit.Clicked += () => { StartExecution(); };

win.Add(
    // The ones with my favorite layout system, Computed
    savePath, path, depth, savePathText, pathText, depthText,

    // The ones laid out like an australopithecus, with Absolute positions:
    submit,
    textBox
) ;

Application.Run();
Application.Shutdown();


//// Wersja z argumentami:
//int maxDepth;
//if (!Int32.TryParse(args[1], out maxDepth))
//{
//    Console.WriteLine("Podana glebokosc nie jest liczba!");
//    return;
//}

//DirectoryInfo directory = new DirectoryInfo(args[0]);
//if (!directory.Exists)
//{
//    Console.WriteLine("Ten katalog nie istnieje!");
//    return;
//}

//CreateExcel(directory, maxDepth);


