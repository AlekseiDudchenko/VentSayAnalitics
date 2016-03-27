using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace CreditApp
{
    class ExcelClass
    {
        // устанавливаем адрес файла
        //private string filename = Environment.CurrentDirectory + "\\17.xlsx";
        //private string filename = "C:\\12.xlsx";
        //private string filename = "D:\\Program\\База.xlsx";
        private string filename = "D:\\Program\\18.xlsx";
     
        /// <summary>
        /// Единичи измерения материалов
        /// </summary>
        public string[] Units = new string[200];

        /// <summary>
        /// Список наименований материалов
        /// </summary>
        public string[] MaterialsNames = new string[200];

        /// <summary>
        /// Количество материалов в файле
        /// </summary>
        public int NamberMaterials;

        public List<string> Providers = new List<string>(); 
 
        /// <summary>
        /// Содержит адрес и имя файла
        /// </summary>
        public string Filename
        {
            get { return filename; }
        }

        public List<string> GetProviders()
        {
            Application excelApp = new Application();
            Workbook workbook = excelApp.Workbooks.Open(Filename);
            Worksheet providerWorksheet = (Worksheet)workbook.Sheets["Поставщики"];
            Range providerRange = providerWorksheet.UsedRange;

            for (int i = 2; i <= providerRange.Rows.Count; i++)
            {
                if (providerWorksheet.Cells[i, 2] != null)
                    Providers.Add(providerWorksheet.Cells[i, 2].Value.ToString());
            }
            // закрываем Excel
            workbook.Close(true, Missing.Value, Missing.Value);
            excelApp.Quit(); 

            return Providers;
        }


        public void GetMaterials()
        {
            // открываем документ и лист для считывания данных для comboBox
            Application excelApp = new Application();
            Workbook workbook = excelApp.Workbooks.Open(Filename);
            Worksheet MatWorksheet = (Worksheet)workbook.Sheets["Mat"];
            Range MatRange = MatWorksheet.UsedRange;

            // проход по всем строкам листа Mat
            for (int i = 3; i <= MatRange.Rows.Count; i++)
            {
                // заполняем comboBox значениями
                MaterialsNames[i - 3] = MatWorksheet.Cells[i, 2].Value.ToString();
                try
                {
                    //TODO: делать проверку в конструкторе класса  
                    // запоминаем единици измерения
                    Units[i - 3] = MatWorksheet.Cells[i, 3].Value.ToString();
                }
                catch (Exception)
                {

                    MessageBox.Show(
                        "Ошибка форматов. Не удалов единици измерения конвертировать в строку. Возможно она состоит только из цифр.");
                }
            }
            
            // количество материалов
            NamberMaterials = MatRange.Rows.Count;

            // закрываем Excel
            workbook.Close(true, Missing.Value, Missing.Value);
            excelApp.Quit();       
        }
    }
}
