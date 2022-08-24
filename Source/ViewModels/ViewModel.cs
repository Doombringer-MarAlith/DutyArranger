using DutyArranger.Source.Commands;
using DutyArranger.Source.Entities;
using DutyArranger.Source.Helpers;
using ExcelDataReader;
using Microsoft.Win32;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Input;

namespace DutyArranger.Source
{
    class ViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        public void OnPropertyChanged([CallerMemberName] string property = "")
        {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(property));
        }

        private Collection<Soldier> _soldiers;
        private SelectedYear _dateSelection;
        private DateTime? _selectedDate;
        private int _lastDataMonth;
        private int _lastDataYear;
        private string _division;
        private string _divisionName;
        private string _makerName;

        private bool _enableSelectingDate = true;
        private bool _enableReadingPreviousData;
        private bool _enableFormulatingData;
        private bool _enableOutputtingData;

        public DateTime? SelectedDate
        {
            get => _selectedDate;
            set
            {
                _selectedDate = value;
                OnPropertyChanged();
            }
        }

        public bool EnableSelectingDate
        {
            get => _enableSelectingDate;
            set
            {
                _enableSelectingDate = value;
                OnPropertyChanged();
            }
        }

        public bool EnableReadingPreviousData
        {
            get => _enableReadingPreviousData;
            set
            {
                _enableReadingPreviousData = value;
                OnPropertyChanged();
            }
        }

        public bool EnableFormulatingData
        {
            get => _enableFormulatingData;
            set
            {
                _enableFormulatingData = value;
                OnPropertyChanged();
            }
        }

        public bool EnableOutputtingData
        {
            get => _enableOutputtingData;
            set
            {
                _enableOutputtingData = value;
                OnPropertyChanged();
            }
        }

        public string Division
        {
            get => _division;
            set
            {
                _division = value;
                OnPropertyChanged();
            }
        }

        public string DivisionName
        {
            get => _divisionName;
            set
            {
                _divisionName = value;
                OnPropertyChanged();
            }
        }

        public string MakerName
        {
            get => _makerName;
            set
            {
                _makerName = value;
                OnPropertyChanged();
            }
        }

        private RelayCommand _chooseInputSourceCommand;
        private RelayCommand _readPreviousDataCommand;
        private RelayCommand _formulateDataCommand;
        private RelayCommand _outputDataCommand;

        public ICommand ReadPreviousDataCommand
        {
            get
            {
                if (_readPreviousDataCommand == null)
                {
                    return _readPreviousDataCommand = new RelayCommand(
                        param => CheckSavedData(),
                        param => EnableReadingPreviousData
                    );
                }
                else
                    return _readPreviousDataCommand;
            }
        }

        public ICommand FormulateDataCommand
        {
            get
            {
                if (_formulateDataCommand == null)
                {
                    return _formulateDataCommand = new RelayCommand(
                        param => FormulateData(),
                        param => EnableFormulatingData
                    );
                }
                else
                    return _formulateDataCommand;
            }
        }

        public ICommand OutputDataCommand
        {
            get
            {
                if (_outputDataCommand == null)
                {
                    return _outputDataCommand = new RelayCommand(
                        param => SaveOutputFile(),
                        param => EnableOutputtingData
                    );
                }
                else
                    return _outputDataCommand;
            }
        }

        public ICommand ChooseInputSourceCommand
        {
            get
            {
                if (_chooseInputSourceCommand == null)
                {
                    return _chooseInputSourceCommand = new RelayCommand(
                        param => ChooseInputSourceFile(),
                        param => EnableSelectingDate
                    );
                }
                else
                    return _chooseInputSourceCommand;
            }
        }

        private void ChooseInputSourceFile()
        {
            var openFileDialog = new OpenFileDialog
            {
                Multiselect = false,
                Filter = "Document files (*.xlsx;*.xls;*.xlsb)|*.xlsx;*.xls;*.xlsb"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                _soldiers = new Collection<Soldier>();
                ReadInputFile(openFileDialog.FileName);
                MessageBox.Show("Duomenų failas buvo sėkmingai nuskaitytas.");
                EnableReadingPreviousData = true;
            }
            else
            {
                EnableReadingPreviousData = false;
                EnableFormulatingData = false;
                EnableOutputtingData = false;
            }
        }

        private void ReadInputFile(string filePath)
        {
            Application.Current.Dispatcher.Invoke(() =>
            {
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                {
                    // Auto-detect format, supports:
                    //  - Binary Excel files (2.0-2003 format; *.xls)
                    //  - OpenXml Excel files (2007 format; *.xlsx, *.xlsb)
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        // 1. Use the reader methods
                        do
                        {
                            var name = "";

                            while (reader.Read())
                            {
                                List<DateTime> holidayList = new List<DateTime> { };

                                name = reader.GetString(0);

                                for (var i = 1; i < reader.FieldCount; i++)
                                {
                                    if (reader.GetFieldType(i) == null)
                                        break;

                                    try
                                    {
                                        DateTime data = (DateTime)reader.GetValue(i);
                                        holidayList.Add(data);
                                    }
                                    catch (Exception)
                                    {
                                        try
                                        {
                                            string data2 = (string)reader.GetValue(i);
                                            string dataFirstPart = string.Concat(data2.TakeWhile(x => x != '-')).Trim();
                                            int dataFirstPartMonth = int.Parse(string.Concat(dataFirstPart.TakeWhile(x => x != '/')));
                                            int dataFirstPartDay = int.Parse(string.Concat(dataFirstPart.SkipWhile(x => x != '/').Skip(1)));
                                            string dataSecondPart = string.Concat(data2.SkipWhile(x => x != '-').Skip(1)).Trim();
                                            int dataSecondPartMonth = int.Parse(string.Concat(dataSecondPart.TakeWhile(x => x != '/')));
                                            int dataSecondPartDay = int.Parse(string.Concat(dataSecondPart.SkipWhile(x => x != '/').Skip(1)));
                                            var firstDate = new DateTime(DateTime.Now.Year, dataFirstPartMonth, dataFirstPartDay);
                                            var secondDate = new DateTime(DateTime.Now.Year, dataSecondPartMonth, dataSecondPartDay);
                                            for (DateTime date = firstDate; date <= secondDate; date = date.AddDays(1))
                                            {
                                                holidayList.Add(date);
                                            }
                                        }
                                        catch (Exception)
                                        {
                                            MessageBox.Show($"Datos kuriomis {name} negali budėti nebuvo nuskaitytos!");
                                            break;
                                        }
                                    }
                                }

                                _soldiers.Add(new Soldier
                                {
                                    FirstName = name,
                                    Holidays = holidayList
                                });
                            }
                        } while (reader.NextResult());
                    }
                }
            });
        }

        private void SaveOutputFile()
        {
            SaveFileDialog openDlg = new SaveFileDialog();
            openDlg.Filter = "Document files (*.xlsx;*.xls;*.xlsb)|*.xlsx;*.xls;*.xlsb";
            openDlg.InitialDirectory = Directory.GetCurrentDirectory();

            if (openDlg.ShowDialog() == true)
            {
                try
                {
                    string path = openDlg.FileName;
                    OutputData(path);
                }

                catch (Exception)
                {
                    return;
                }
            }
        }

        private void OutputData(string outputFileNamePath)
        {
            var outputFile = $"{Directory.GetCurrentDirectory()}/Išvestis/Grafikas_{_dateSelection.Year}_{_dateSelection.SelectedMonth.Month}.xlsx";

            try
            {
                if (string.IsNullOrEmpty(DivisionName))
                {
                    MessageBox.Show("Įveskite padalinio, kuriam kuriamas grafikas, pavadinimą.");
                    return;
                }

                if (string.IsNullOrEmpty(Division) || string.IsNullOrEmpty(MakerName))
                {
                    MessageBox.Show("Įveskite padalinį ir asmens vardą bei pavardę.");
                    return;
                }

                if (!Directory.Exists(Directory.GetCurrentDirectory() + "/Išvestis"))
                {
                    Directory.CreateDirectory(Directory.GetCurrentDirectory() + "/Išvestis");
                }

                using (var fs = new FileStream(outputFile, FileMode.Create, FileAccess.Write))
                {
                    IWorkbook workbook = new XSSFWorkbook();
                    ISheet sheet1 = workbook.CreateSheet("Grafikas");

                    // Style the cell with borders all around.
                    ICellStyle style = workbook.CreateCellStyle();
                    style.BorderBottom = BorderStyle.Thin;
                    style.BottomBorderColor = HSSFColor.Black.Index;
                    style.BorderLeft = BorderStyle.Thin;
                    style.LeftBorderColor = HSSFColor.Black.Index;
                    style.BorderRight = BorderStyle.Thin;
                    style.RightBorderColor = HSSFColor.Black.Index;
                    style.BorderTop = BorderStyle.Thin;
                    style.TopBorderColor = HSSFColor.Black.Index;
                    style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                    style.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;

                    ICellStyle dayNumberStyle = workbook.CreateCellStyle();
                    dayNumberStyle.BorderBottom = BorderStyle.Thick;
                    dayNumberStyle.BottomBorderColor = HSSFColor.Black.Index;
                    dayNumberStyle.BorderLeft = BorderStyle.Thin;
                    dayNumberStyle.LeftBorderColor = HSSFColor.Black.Index;
                    dayNumberStyle.BorderRight = BorderStyle.Thin;
                    dayNumberStyle.RightBorderColor = HSSFColor.Black.Index;
                    dayNumberStyle.BorderTop = BorderStyle.Thin;
                    dayNumberStyle.TopBorderColor = HSSFColor.Black.Index;
                    dayNumberStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                    dayNumberStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;

                    ICellStyle weekendStyle = workbook.CreateCellStyle();
                    weekendStyle.BorderBottom = BorderStyle.Thin;
                    weekendStyle.BottomBorderColor = HSSFColor.Black.Index;
                    weekendStyle.BorderLeft = BorderStyle.Thin;
                    weekendStyle.LeftBorderColor = HSSFColor.Black.Index;
                    weekendStyle.BorderRight = BorderStyle.Thin;
                    weekendStyle.RightBorderColor = HSSFColor.Black.Index;
                    weekendStyle.BorderTop = BorderStyle.Thin;
                    weekendStyle.TopBorderColor = HSSFColor.Black.Index;
                    weekendStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                    weekendStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
                    weekendStyle.FillForegroundColor = HSSFColor.Grey25Percent.Index;
                    weekendStyle.FillPattern = FillPattern.SolidForeground;

                    ICellStyle weekendNumberStyle = workbook.CreateCellStyle();
                    weekendNumberStyle.BorderBottom = BorderStyle.Thick;
                    weekendNumberStyle.BottomBorderColor = HSSFColor.Black.Index;
                    weekendNumberStyle.BorderLeft = BorderStyle.Thin;
                    weekendNumberStyle.LeftBorderColor = HSSFColor.Black.Index;
                    weekendNumberStyle.BorderRight = BorderStyle.Thin;
                    weekendNumberStyle.RightBorderColor = HSSFColor.Black.Index;
                    weekendNumberStyle.BorderTop = BorderStyle.Thin;
                    weekendNumberStyle.TopBorderColor = HSSFColor.Black.Index;
                    weekendNumberStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                    weekendNumberStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
                    weekendNumberStyle.FillForegroundColor = HSSFColor.Grey25Percent.Index;
                    weekendNumberStyle.FillPattern = FillPattern.SolidForeground;

                    ICellStyle holidayStyle = workbook.CreateCellStyle();
                    holidayStyle.BorderBottom = BorderStyle.Thin;
                    holidayStyle.BottomBorderColor = HSSFColor.Black.Index;
                    holidayStyle.BorderLeft = BorderStyle.Thin;
                    holidayStyle.LeftBorderColor = HSSFColor.Black.Index;
                    holidayStyle.BorderRight = BorderStyle.Thin;
                    holidayStyle.RightBorderColor = HSSFColor.Black.Index;
                    holidayStyle.BorderTop = BorderStyle.Thin;
                    holidayStyle.TopBorderColor = HSSFColor.Black.Index;
                    holidayStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                    holidayStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
                    holidayStyle.FillForegroundColor = HSSFColor.Grey50Percent.Index;
                    holidayStyle.FillPattern = FillPattern.SolidForeground;

                    ICellStyle nameStyle = workbook.CreateCellStyle();
                    nameStyle.BorderBottom = BorderStyle.Thin;
                    nameStyle.BottomBorderColor = HSSFColor.Black.Index;
                    nameStyle.BorderLeft = BorderStyle.Thin;
                    nameStyle.LeftBorderColor = HSSFColor.Black.Index;
                    nameStyle.BorderRight = BorderStyle.Thin;
                    nameStyle.RightBorderColor = HSSFColor.Black.Index;
                    nameStyle.BorderTop = BorderStyle.Thin;
                    nameStyle.TopBorderColor = HSSFColor.Black.Index;
                    nameStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
                    nameStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;

                    IFont confirmationFont = workbook.CreateFont();
                    confirmationFont.FontHeightInPoints = 9;
                    //confirmationFont.FontName = "Courier New";
                    //confirmationFont.IsItalic = true;
                    //confirmationFont.IsStrikeout = true;

                    IFont boldFont = workbook.CreateFont();
                    boldFont.IsBold = true;

                    ICellStyle emptyBorderCellStyle = workbook.CreateCellStyle();
                    emptyBorderCellStyle.TopBorderColor = HSSFColor.Black.Index;
                    emptyBorderCellStyle.LeftBorderColor = HSSFColor.Black.Index;
                    emptyBorderCellStyle.RightBorderColor = HSSFColor.Black.Index;
                    emptyBorderCellStyle.BottomBorderColor = HSSFColor.Black.Index;

                    ICellStyle fullBorderCellStyle = workbook.CreateCellStyle();
                    fullBorderCellStyle.BorderTop = BorderStyle.Thick;
                    fullBorderCellStyle.TopBorderColor = HSSFColor.Black.Index;
                    fullBorderCellStyle.BorderLeft = BorderStyle.Thick;
                    fullBorderCellStyle.LeftBorderColor = HSSFColor.Black.Index;
                    fullBorderCellStyle.BorderRight = BorderStyle.Thick;
                    fullBorderCellStyle.RightBorderColor = HSSFColor.Black.Index;
                    fullBorderCellStyle.BorderBottom = BorderStyle.Thick;
                    fullBorderCellStyle.BottomBorderColor = HSSFColor.Black.Index;
                    fullBorderCellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                    fullBorderCellStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;

                    ICellStyle confimationStyle = workbook.CreateCellStyle();
                    confimationStyle.SetFont(confirmationFont);

                    ICellStyle legendStyle = workbook.CreateCellStyle();
                    legendStyle.SetFont(boldFont);

                    // Row 1
                    IRow row = sheet1.CreateRow(0);

                    // Row 2
                    row = sheet1.CreateRow(1);
                    row.CreateCell(24).SetCellValue($"PATVIRTINTA\nLietuvos kariuomenės\nLietuvos didžiojo etmono\nJono Karolio Chodkevičiaus\npėstininkų brigados \"Žemaitija\"\nLietuvos didžiojo kunigaikščio\nButigeidžio dragūnų bataliono\nvado {_dateSelection.Year} m.\nįsakymu Nr. V-");
                    row.GetCell(24).CellStyle = confimationStyle;

                    // Row 3-9
                    row = sheet1.CreateRow(2);
                    row.CreateCell(24);
                    row.CreateCell(25);
                    row.CreateCell(26);
                    row.CreateCell(27);
                    row = sheet1.CreateRow(3);
                    row.CreateCell(24);
                    row.CreateCell(25);
                    row.CreateCell(26);
                    row.CreateCell(27);
                    row = sheet1.CreateRow(4);
                    row.CreateCell(24);
                    row.CreateCell(25);
                    row.CreateCell(26);
                    row.CreateCell(27);
                    row = sheet1.CreateRow(5);
                    row.CreateCell(24);
                    row.CreateCell(25);
                    row.CreateCell(26);
                    row.CreateCell(27);
                    row = sheet1.CreateRow(6);
                    row.CreateCell(24);
                    row.CreateCell(25);
                    row.CreateCell(26);
                    row.CreateCell(27);
                    row = sheet1.CreateRow(7);
                    row.CreateCell(24);
                    row.CreateCell(25);
                    row.CreateCell(26);
                    row.CreateCell(27);
                    row = sheet1.CreateRow(8);
                    row.CreateCell(24);
                    row.CreateCell(25);
                    row.CreateCell(26);
                    row.CreateCell(27);

                    // Row 10
                    row = sheet1.CreateRow(9);

                    // Row 11
                    row = sheet1.CreateRow(10);
                    row.CreateCell(1).SetCellValue($"{_dateSelection.Year} m. {Utilities.TranslatedMonths[_dateSelection.SelectedMonth.Month]} mėn. {DivisionName} dienos tarnybos grafikas");
                    row.GetCell(1).CellStyle = workbook.CreateCellStyle();

                    // Row 12
                    row = sheet1.CreateRow(11);
                    row.CreateCell(1).SetCellValue("Eil. Nr.");
                    row.CreateCell(2).SetCellValue("Karinis laipsnis, vardas, pavardė");
                    row.CreateCell(3).SetCellValue(_dateSelection.Year + " m. " + Utilities.TranslatedMonths[_dateSelection.SelectedMonth.Month] + " mėn.");
                    row.GetCell(1).CellStyle = fullBorderCellStyle;
                    row.GetCell(2).CellStyle = fullBorderCellStyle;
                    row.GetCell(3).CellStyle = fullBorderCellStyle;
                    for (int count = 4; count <= _dateSelection.SelectedMonth.SelectedDays.Count + 2; count++)
                    {
                        row.CreateCell(count).CellStyle = style;
                    }

                    // Row 13
                    row = sheet1.CreateRow(12);
                    row.CreateCell(1);
                    row.CreateCell(2);
                    row.GetCell(1).CellStyle = style;
                    row.GetCell(2).CellStyle = style;

                    var dayCount = 1;

                    foreach (var day in _dateSelection.SelectedMonth.SelectedDays)
                    {
                        row.CreateCell(2 + dayCount).SetCellValue(day.Day);
                        row.GetCell(2 + dayCount).CellStyle = dayNumberStyle;

                        if (day.Date.DayOfWeek == DayOfWeek.Saturday || day.Date.DayOfWeek == DayOfWeek.Sunday)
                        {
                            row.GetCell(2 + dayCount).CellStyle = weekendNumberStyle;
                            row.GetCell(2 + dayCount).CellStyle = weekendNumberStyle;
                        }

                        dayCount++;
                    }

                    var rowCount = 13;
                    var columnCount = 3;
                    var soldierCount = 1;

                    // Row 14-x
                    foreach (var soldier in _soldiers)
                    {
                        row = sheet1.CreateRow(rowCount);
                        row.CreateCell(1).SetCellValue(soldierCount);
                        row.GetCell(1).CellStyle = fullBorderCellStyle;

                        row.CreateCell(2).SetCellValue(soldier.FirstName);
                        row.GetCell(2).CellStyle = fullBorderCellStyle;
                        row.GetCell(2).CellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;

                        columnCount = 3;

                        foreach (var selectedDay in _dateSelection.SelectedMonth.SelectedDays)
                        {
                            if (soldier.Holidays.Any(x => x.Day == selectedDay.Day && x.Month == _dateSelection.SelectedMonth.Month))
                            {
                                row.CreateCell(columnCount);
                            }
                            else if (soldier.DaysOnGuard.Any(x => x.Day == selectedDay.Day && x.Month == _dateSelection.SelectedMonth.Month))
                            {
                                row.CreateCell(columnCount).SetCellValue("B");
                            }
                            else if (soldier.DaysOnReserve.Any(x => x.Day == selectedDay.Day && x.Month == _dateSelection.SelectedMonth.Month))
                            {
                                row.CreateCell(columnCount).SetCellValue("R");
                            }
                            else
                            {
                                row.CreateCell(columnCount);
                            }

                            // Paint weekends in light grey, holidays in darker grey
                            if (soldier.Holidays.Any(x => x.Day == selectedDay.Day && x.Month == _dateSelection.SelectedMonth.Month))
                            {
                                row.GetCell(columnCount).CellStyle = holidayStyle;
                            }
                            else if (selectedDay.Date.DayOfWeek == DayOfWeek.Saturday || selectedDay.Date.DayOfWeek == DayOfWeek.Sunday)
                            {
                                row.GetCell(columnCount).CellStyle = weekendStyle;
                            }
                            else
                                row.GetCell(columnCount).CellStyle = style;

                            columnCount++;
                        }

                        rowCount++;
                        soldierCount++;
                    }

                    var rowNumForReference = sheet1.PhysicalNumberOfRows;
                    row = sheet1.CreateRow(rowNumForReference);
                    row.CreateCell(3);

                    // Last Row
                    row = sheet1.CreateRow(rowNumForReference + 1);
                    row.CreateCell(2).SetCellValue("B - kuopos budėtojas");
                    row.GetCell(2).CellStyle = legendStyle;
                    row.CreateCell(4).SetCellValue("R - kuopos budėtojo rezervas");
                    row.GetCell(4).CellStyle = legendStyle;

                    // Last Row (+1-2) (empty)
                    row = sheet1.CreateRow(rowNumForReference + 2);
                    row = sheet1.CreateRow(rowNumForReference + 3);

                    // Last Row (+3)
                    row = sheet1.CreateRow(rowNumForReference + 4);
                    row.CreateCell(2).SetCellValue(Division);
                    row.CreateCell(25).SetCellValue(MakerName);

                    var cra = new NPOI.SS.Util.CellRangeAddress(11, 11, 3, 2 + _dateSelection.SelectedMonth.SelectedDays.Count);
                    sheet1.AddMergedRegion(cra);
                    sheet1.GetRow(11).GetCell(3).CellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                    sheet1.GetRow(11).GetCell(3).CellStyle.SetFont(boldFont);
                    RegionUtil.SetBorderBottom((int)BorderStyle.Thick, cra, sheet1);
                    RegionUtil.SetBorderLeft((int)BorderStyle.Thick, cra, sheet1);
                    RegionUtil.SetBorderRight((int)BorderStyle.Thick, cra, sheet1); 
                    RegionUtil.SetBorderTop((int)BorderStyle.Thick, cra, sheet1);

                    cra = new NPOI.SS.Util.CellRangeAddress(rowNumForReference, rowNumForReference, 3, 2 + _dateSelection.SelectedMonth.SelectedDays.Count);
                    sheet1.AddMergedRegion(cra);
                    RegionUtil.SetBorderTop((int)BorderStyle.Thick, cra, sheet1);

                    cra = new NPOI.SS.Util.CellRangeAddress(11, 11 + _soldiers.Count + 1, 2 + _dateSelection.SelectedMonth.SelectedDays.Count + 1, 2 + _dateSelection.SelectedMonth.SelectedDays.Count + 1);
                    sheet1.AddMergedRegion(cra);
                    RegionUtil.SetBorderLeft((int)BorderStyle.Thick, cra, sheet1);

                    cra = new NPOI.SS.Util.CellRangeAddress(11, 12, 1, 1);
                    sheet1.AddMergedRegion(cra);
                    RegionUtil.SetBorderBottom((int)BorderStyle.Thick, cra, sheet1);
                    RegionUtil.SetBorderLeft((int)BorderStyle.Thick, cra, sheet1);
                    RegionUtil.SetBorderRight((int)BorderStyle.Thick, cra, sheet1);
                    RegionUtil.SetBorderTop((int)BorderStyle.Thick, cra, sheet1);
                    cra = new NPOI.SS.Util.CellRangeAddress(11, 12, 2, 2);
                    sheet1.AddMergedRegion(cra);
                    RegionUtil.SetBorderBottom((int)BorderStyle.Thick, cra, sheet1);
                    RegionUtil.SetBorderLeft((int)BorderStyle.Thick, cra, sheet1);
                    RegionUtil.SetBorderRight((int)BorderStyle.Thick, cra, sheet1);
                    RegionUtil.SetBorderTop((int)BorderStyle.Thick, cra, sheet1);
                    cra = new NPOI.SS.Util.CellRangeAddress(1, 8, 24, 29);
                    sheet1.AddMergedRegion(cra);

                    cra = new NPOI.SS.Util.CellRangeAddress(10, 10, 1, 1 + _dateSelection.SelectedMonth.SelectedDays.Count + 1);
                    sheet1.AddMergedRegion(cra);
                    sheet1.GetRow(10).GetCell(1).CellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                    sheet1.GetRow(10).GetCell(1).CellStyle.SetFont(boldFont);

                    cra = new NPOI.SS.Util.CellRangeAddress(rowNumForReference + 1, rowNumForReference + 1, 4, 12);
                    sheet1.AddMergedRegion(cra);

                    cra = new NPOI.SS.Util.CellRangeAddress(rowNumForReference + 4, rowNumForReference + 4, 25, 29);
                    sheet1.AddMergedRegion(cra);

                    sheet1.SetColumnWidth(2, 8192);
                    foreach (var col in sheet1.GetRow(13).Cells)
                    {
                        if (col.ColumnIndex < 3)
                            continue;

                        sheet1.SetColumnWidth(col.ColumnIndex, 1024);
                    }

                    workbook.Write(fs);
                    EnableFormulatingData = false;
                    EnableOutputtingData = false;
                    EnableReadingPreviousData = false;
                }

                File.Copy(outputFile, outputFileNamePath);
                MessageBox.Show($"Rezultatų failas {outputFileNamePath} sėkmingai sukurtas.");
            }
            catch (Exception)
            {
                MessageBox.Show("Klaida. Nepavyko sukurti rezultatų failo.");
            }
        }

        private void CheckSavedData()
        {
            if (SelectedDate == null)
            {
                MessageBox.Show("Pasirinkite mėnesį!");
                return;
            }

            string recentDataFile = string.Empty;
            string temp = string.Empty;
            int highestYear = 0;
            int highestMonth = 0;

            try
            {
                // Iterate once to determine most recent year
                foreach (var file in Directory.GetFiles(Directory.GetCurrentDirectory() + "/Išvestis"))
                {
                    if (file.Contains(".xlsx"))
                    {
                        if (file.Contains("Grafikas_"))
                        {
                            temp = Path.GetFileNameWithoutExtension(file).Substring(9);
                            var parsedYear = int.Parse(temp.Substring(0, 4));
                            highestYear = highestYear < parsedYear ? parsedYear : highestYear;
                        }
                    }
                }

                // No year found means there are no files suitable to check
                if (highestYear == 0)
                {
                    MessageBox.Show("Praėjusių mėnesių duomenys nebuvo nuskaityti.");
                    EnableFormulatingData = true;
                    return;
                }

                // Iterate second time to determine most recent month of that year
                foreach (var file in Directory.GetFiles(Directory.GetCurrentDirectory() + "/Išvestis"))
                {
                    if (file.Contains(".xlsx"))
                    {
                        if (file.Contains("Grafikas_"))
                        {
                            if (file.Contains(highestYear.ToString()))
                            {
                                temp = Path.GetFileNameWithoutExtension(file).Substring(9);
                                var parsedMonth = int.Parse(temp.Substring(5));
                                highestMonth = highestMonth < parsedMonth ? parsedMonth : highestMonth;
                            }
                        }
                    }
                }

                if (highestMonth == 0)
                {
                    MessageBox.Show("Praėjusių mėnesių duomenys nebuvo nuskaityti.");
                    EnableFormulatingData = true;
                    return;
                }

                _lastDataMonth = highestMonth;
                _lastDataYear = highestYear;

                recentDataFile = Directory.GetFiles(Directory.GetCurrentDirectory() + "/Išvestis").FirstOrDefault(file => file.Contains(highestYear + "_" + highestMonth));

                using (var fs = new FileStream(recentDataFile, FileMode.Open, FileAccess.Read))
                {
                    IWorkbook workbook = WorkbookFactory.Create(fs);
                    ISheet sheet = workbook.GetSheet("Grafikas");
                    for (int row = 13; row <= sheet.LastRowNum; row++)
                    {
                        if (sheet.GetRow(row) != null) //null is when the row only contains empty cells 
                        {
                            if (sheet.GetRow(row).GetCell(2) == null)
                                continue;

                            // Check if the same soldiers
                            var soldier = _soldiers.FirstOrDefault(x => x.FirstName == sheet.GetRow(row).GetCell(2).StringCellValue);
                            if (soldier != null)
                            {
                                foreach (var cell in sheet.GetRow(row).Cells.OrderBy(y => y.ColumnIndex))
                                {
                                    // Skip name and row number
                                    if (cell.ColumnIndex < 3)
                                        continue;

                                    if (cell.StringCellValue == "B" || cell.StringCellValue == "R")
                                    {
                                        soldier.LastDayOnDutyFromPreviousMonth = cell.ColumnIndex - 2;
                                        Console.WriteLine($"name: {soldier.FirstName} last day on duty: {cell.ColumnIndex - 2}");
                                    }
                                }
                            }
                        }
                    }
                }

                MessageBox.Show("Praėjusių mėnesių duomenys buvo sėkmingai nuskaityti.");
                EnableFormulatingData = true;
            }
            catch (Exception)
            {
                MessageBox.Show("Nepavyko nuskaityti praėjusių mėnesių duomenų.");
                EnableFormulatingData = true;
            }
        }

        private void FormulateData()
        {
            try
            {
                if (SelectedDate == null)
                {
                    MessageBox.Show("Pasirinkite mėnesį!");
                    return;
                }

                Soldier guard;
                Soldier reserve;

                var selectedDate = SelectedDate;
                int selectedYear = (int)(selectedDate?.Year);
                int selectedMonth = (int)(selectedDate?.Month);

                var currentDay = 1;

                _dateSelection = new SelectedYear
                {
                    Year = selectedYear
                };

                _dateSelection.SelectedMonth.Month = selectedMonth;

                var soldiersShuffled = _soldiers.ToList();
                Utilities.Shuffle(soldiersShuffled);

                // Each day needs one person on guard and one person on reserve
                while (currentDay <= DateTime.DaysInMonth(selectedYear, selectedMonth))
                {
                    if (Utilities.Roll() >= 50000)
                    {
                        // Try to assign a guard
                        guard = AssignGuard(currentDay, selectedMonth, soldiersShuffled, false);
                        if (guard == null)
                        {
                            MessageBox.Show($"Klaida. Neįmanoma priskirti budinčio kario {currentDay}-ai dienai!");
                            return;
                        }

                        // Try to assign a reserve
                        reserve = AssignReserve(currentDay, selectedMonth, soldiersShuffled, false);
                        if (reserve == null)
                        {
                            MessageBox.Show($"Klaida. Neįmanoma priskirti budinčio kario {currentDay}-ai dienai!");
                            return;
                        }
                    }
                    else
                    {
                        // Try to assign a reserve
                        reserve = AssignReserve(currentDay, selectedMonth, soldiersShuffled, false);
                        if (reserve == null)
                        {
                            MessageBox.Show($"Klaida. Neįmanoma priskirti budinčio kario {currentDay}-ai dienai!");
                            return;
                        }

                        // Try to assign a guard
                        guard = AssignGuard(currentDay, selectedMonth, soldiersShuffled, false);
                        if (guard == null)
                        {
                            MessageBox.Show($"Klaida. Neįmanoma priskirti budinčio kario {currentDay}-ai dienai!");
                            return;
                        }
                    }

                    _dateSelection.SelectedMonth.SelectedDays.Add(new SelectedDay
                    {
                        Date = new DateTime(selectedYear, selectedMonth, currentDay),
                        Day = currentDay,
                        Guard = guard,
                        Reservee = reserve
                    });

                    currentDay++;
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Klaida. Grafikas nebuvo sudarytas.");
                EnableOutputtingData = false;
            }

            MessageBox.Show("Grafikas buvo sėkmingai sudarytas.");
            EnableOutputtingData = true;
        }

        private bool CouldBeAssigned(Soldier soldier, int startDay, int lastDay, bool reserve)
        {
            var testContainer = new List<Soldier>();
            testContainer.Add(soldier);

            for (int i = startDay; startDay <= lastDay; startDay++)
            {
                if (reserve)
                {
                    if (AssignReserve(i, _dateSelection.SelectedMonth.Month, testContainer, true) != null)
                        return true;
                }
                else
                {
                    if (AssignGuard(i, _dateSelection.SelectedMonth.Month, testContainer, true) != null)
                        return true;
                }
            }

            return false;
        }

        private Soldier AssignGuard(int dayIndex, int monthIndex, List<Soldier> soldiers, bool checkOnly)
        {
            if (_soldiers.Count <= 0)
                return null;

            foreach (var row in soldiers)
            {
                if (row == null)
                    continue;

                // Check if available on that day
                if (row.Holidays.Any(record => record.Day == dayIndex && record.Month == monthIndex))
                    continue;

                // Check if already on reserve on that day
                if (row.DaysOnReserve.Any(record => record.Day == dayIndex && record.Month == monthIndex))
                    continue;

                // Check if already guarding on that day
                if (row.DaysOnGuard.Any(record => record.Day == dayIndex && record.Month == monthIndex))
                    continue;

                var lastDayOnGuard = row.DaysOnGuard.Count > 0 ? row.DaysOnGuard.Max(x => x.Day) : 0;
                var lastDayOnReserve = row.DaysOnReserve.Count > 0 ? row.DaysOnReserve.Max(x => x.Day) : 0;
                var lastDayOnDuty = lastDayOnGuard > lastDayOnReserve ? lastDayOnGuard : lastDayOnReserve;

                // Check if already has been on any duty whithin last 4 days
                if (lastDayOnDuty + 4 >= dayIndex && lastDayOnDuty != 0)
                    continue;

                // Check if there is data for the previous month
                if ((_lastDataYear == _dateSelection.Year && _lastDataMonth + 1 == _dateSelection.SelectedMonth.Month)
                    || (_lastDataYear == _dateSelection.Year - 1 && _lastDataMonth == 12)) // Year jump case, check if last month was december
                {
                    // Skip if 4 days haven't passed from the previous month
                    if (DateTime.DaysInMonth(_lastDataYear, _lastDataMonth) - row.LastDayOnDutyFromPreviousMonth + dayIndex < 4)
                        continue;
                }

                if (checkOnly)
                    return row;

                // Lower priority for those who had already guarded/reserved this month
                var timesOnDutyThisMonth = row.TimesOnDutyThisMonth(monthIndex);
                if (row.HasAlreadyBeenOnDutyThisMonth(monthIndex))
                    if (_soldiers.Any(candidate => candidate.TimesOnDutyThisMonth(monthIndex) < timesOnDutyThisMonth && CouldBeAssigned(candidate, dayIndex, DateTime.DaysInMonth(_dateSelection.Year, monthIndex), false)))
                        continue;

                // Put on guard
                var dateTime = new DateTime(_dateSelection.Year, monthIndex, dayIndex);
                row.DaysOnGuard.Add(dateTime);
                return row;
            }

            return null;
        }

        private Soldier AssignReserve(int dayIndex, int monthIndex, List<Soldier> soldiers, bool checkOnly)
        {
            if (_soldiers.Count <= 0)
                return null;

            foreach (var row in soldiers)
            {
                if (row == null)
                    continue;

                // Check if available on that day
                if (row.Holidays.Any(record => record.Day == dayIndex && record.Month == monthIndex))
                    continue;

                // Check if already guarding on that day
                if (row.DaysOnGuard.Any(record => record.Day == dayIndex && record.Month == monthIndex))
                    continue;

                // Check if already on reserve on that day
                if (row.DaysOnReserve.Any(record => record.Day == dayIndex && record.Month == monthIndex))
                    continue;

                var lastDayOnGuard = row.DaysOnGuard.Count > 0 ? row.DaysOnGuard.Max(x => x.Day) : 0;
                var lastDayOnReserve = row.DaysOnReserve.Count > 0 ? row.DaysOnReserve.Max(x => x.Day) : 0;
                var lastDayOnDuty = lastDayOnGuard > lastDayOnReserve ? lastDayOnGuard : lastDayOnReserve;

                // Check if already has been on any duty whithin last 4 days
                if (lastDayOnDuty + 4 >= dayIndex && lastDayOnDuty != 0)
                    continue;

                // Check if there is data for the previous month
                if ((_lastDataYear == _dateSelection.Year && _lastDataMonth + 1 == _dateSelection.SelectedMonth.Month)
                    || (_lastDataYear == _dateSelection.Year - 1 && _lastDataMonth == 12)) // Year jump case, check if last month was december
                {
                    // Skip if 4 days haven't passed from the previous month
                    if (DateTime.DaysInMonth(_lastDataYear, _lastDataMonth) - row.LastDayOnDutyFromPreviousMonth + dayIndex < 4)
                        continue;
                }

                if (checkOnly)
                    return row;

                // Lower priority for those who had already guarded/reserved this month
                var timesOnDutyThisMonth = row.TimesOnDutyThisMonth(monthIndex);
                if (row.HasAlreadyBeenOnDutyThisMonth(monthIndex))
                    if (_soldiers.Any(candidate => candidate.TimesOnDutyThisMonth(monthIndex) < timesOnDutyThisMonth && CouldBeAssigned(candidate, dayIndex, DateTime.DaysInMonth(_dateSelection.Year, monthIndex), true)))
                        continue;

                // Put on reserve
                var dateTime = new DateTime(_dateSelection.Year, monthIndex, dayIndex);
                row.DaysOnReserve.Add(dateTime);
                return row;
            }

            return null;
        }
    }
}
