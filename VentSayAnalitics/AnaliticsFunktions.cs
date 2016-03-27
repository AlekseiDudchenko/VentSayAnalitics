using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CreditApp;
using Microsoft.Office.Interop.Excel;

namespace VentSayAnalitics
{
    class AnaliticsFunktions
    {
        public ObservableCollection<AnaliticsMaterial> AnaliticsMaterialsDebitCollection;

        ExcelClass excel = new ExcelClass();
        Application excelApp = new Application();
        private Workbook workbook;

        public void OpenFile()
        {          
            workbook = excelApp.Workbooks.Open(excel.Filename);
        }

        public void CloseFile()
        {
            excelApp.Quit();
        }





        public ObservableCollection<AnaliticsMaterial> GetMaterialsDebitCollection(int materialNumber)
        {
            materialNumber += 4;

            ExcelClass excel = new ExcelClass();

            AnaliticsMaterialsDebitCollection = new ObservableCollection<AnaliticsMaterial>();
            
            
            //Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            //Workbook workbook = excelApp.Workbooks.Open(excel.Filename);

            Worksheet materialDebitWorksheet = (Worksheet)workbook.Sheets["Приход материалов"];
            Range debitRange = materialDebitWorksheet.UsedRange;
            Worksheet analiticaWorksheet = workbook.Sheets["Аналитика"];
            Worksheet priceDebitWorksheet = (Worksheet)workbook.Sheets["Цена прихода"];
            Worksheet creditWorksheet = workbook.Sheets["Расход материалов"];
            Range creditRange = creditWorksheet.UsedRange;
            Worksheet costsWorksheet = workbook.Sheets["Стоимость прихода"];
            
            int count = 1;

            for (int i = 3; i < materialDebitWorksheet.UsedRange.Rows.Count; i++)
            {
                //для первого материала
                AnaliticsMaterial newAnaliticsMaterial = new AnaliticsMaterial();

                newAnaliticsMaterial.Debit = Convert.ToDouble(materialDebitWorksheet.Cells[i, materialNumber].Value);
                newAnaliticsMaterial.DebitString = Convert.ToString(materialDebitWorksheet.Cells[i, materialNumber].Value);

                
                if (newAnaliticsMaterial.DebitString != null)
                {
                    newAnaliticsMaterial.Cost = Convert.ToDouble(costsWorksheet.Cells[i, materialNumber].Value);
                    newAnaliticsMaterial.CostString = Convert.ToString(costsWorksheet.Cells[i, materialNumber].Value);

                    newAnaliticsMaterial.Date = Convert.ToString(materialDebitWorksheet.Cells[i, 2].Value);

                    newAnaliticsMaterial.Balance = Convert.ToDouble(analiticaWorksheet.Cells[5, materialNumber].Value);

                    newAnaliticsMaterial.Price = Convert.ToDouble(priceDebitWorksheet.Cells[i, materialNumber].Value);
                    newAnaliticsMaterial.PriceString =
                        Convert.ToString(priceDebitWorksheet.Cells[i, materialNumber].Value);

                    newAnaliticsMaterial.DocumentNomber = Convert.ToString(materialDebitWorksheet.Cells[i, 3].Value);

                    newAnaliticsMaterial.Provider = Convert.ToString(priceDebitWorksheet.Cells[i, 1].Value);

                    newAnaliticsMaterial.Count = count;
                    count += 1;

                    AnaliticsMaterialsDebitCollection.Add(newAnaliticsMaterial);
                }
            }

           
            
            //excelApp.Quit();
            
            return AnaliticsMaterialsDebitCollection;
        }

        public ObservableCollection<AnaliticsMaterial> AnaliticsMaterialsCreditCollection;

        public ObservableCollection<AnaliticsMaterial> GetMaterialsCreditCollection(int materialNumber)
        {
            materialNumber += 4;

            ExcelClass excel = new ExcelClass();

            AnaliticsMaterialsCreditCollection = new ObservableCollection<AnaliticsMaterial>();
            
            //Application excelApp = new Application();
            //Workbook workbook = excelApp.Workbooks.Open(excel.Filename);

            Worksheet materialDebitWorksheet = (Worksheet)workbook.Sheets["Приход материалов"];
            Range debitRange = materialDebitWorksheet.UsedRange;
            Worksheet analiticaWorksheet = workbook.Sheets["Аналитика"];
            Worksheet priceDebitWorksheet = (Worksheet)workbook.Sheets["Цена прихода"];
            Worksheet creditWorksheet = workbook.Sheets["Расход материалов"];
            Range creditRange = creditWorksheet.UsedRange;
            Worksheet costsWorksheet = workbook.Sheets["Стоимость прихода"];
            
            int count = 1;

            for (int i = 3; i < creditWorksheet.UsedRange.Rows.Count; i++)
            {
                //для первого материала
                AnaliticsMaterial newAnaliticsMaterial = new AnaliticsMaterial();

                newAnaliticsMaterial.Credit = Convert.ToDouble(creditWorksheet.Cells[i, materialNumber].Value);
                newAnaliticsMaterial.CreditString = Convert.ToString(creditWorksheet.Cells[i, materialNumber].Value);

                if (newAnaliticsMaterial.CreditString != null)
                {
                    //newAnaliticsMaterial.Cost = Convert.ToDouble(costsWorksheet.Cells[i, materialNumber].Value);
                    //newAnaliticsMaterial.CostString = Convert.ToString(costsWorksheet.Cells[i, materialNumber].Value);

                    newAnaliticsMaterial.Date = Convert.ToString(creditWorksheet.Cells[i, 2].Value);

                    //newAnaliticsMaterial.Balance = Convert.ToDouble(analiticaWorksheet.Cells[5, materialNumber].Value);

                    //newAnaliticsMaterial.Price = Convert.ToDouble(priceDebitWorksheet.Cells[i, materialNumber].Value);
                    //newAnaliticsMaterial.PriceString =Convert.ToString(priceDebitWorksheet.Cells[i, materialNumber].Value);

                    newAnaliticsMaterial.DocumentNomber = Convert.ToString(creditWorksheet.Cells[i, 3].Value);

                    //newAnaliticsMaterial.Provider = Convert.ToString(priceDebitWorksheet.Cells[i, 1].Value);

                    newAnaliticsMaterial.Count = count;
                    count += 1;

                    AnaliticsMaterialsCreditCollection.Add(newAnaliticsMaterial);
                }
            }


            //excelApp.Quit();

            return AnaliticsMaterialsCreditCollection;
        }

    }
}
