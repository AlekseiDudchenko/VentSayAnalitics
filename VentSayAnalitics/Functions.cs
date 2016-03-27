using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Media;
using Microsoft.Office.Interop.Excel;
using Label = System.Windows.Controls.Label;
using TextBox = System.Windows.Controls.TextBox;


namespace CreditApp
{
    /// <summary>
    /// Хранит набор функций используемых в этом приложении
    /// </summary>
    class Functions
    {
        /// <summary>
        /// Проверяет введены ли данные верно. 
        /// Значения текстбоксов должны конвертироваться в double и отличаться от 0, 
        /// Элемент в comboBox должен быть не пустым и должен быть выбран
        /// Возвращает true если данные введены верно
        /// </summary>
        /// <param name="textBox1"></param>
        /// <param name="textBox2"></param>
        /// <param name="comboBox"></param>
        /// <returns></returns>
        public static bool ProverkaDannih(TextBox textBox1, TextBox textBox2, ComboBox comboBox)
        {
            bool convertToDouble;
            bool result;
            try
            {
                Convert.ToDouble(textBox1.Text);
                Convert.ToDouble(textBox2.Text);
                convertToDouble = true;
            }
            catch (FormatException)
            {
                convertToDouble = false;
            }

            // Если не пустые строчи и конвертируется в цифры
            if (convertToDouble & comboBox.SelectedIndex != -1)
            {
                if (Convert.ToInt32(textBox1.Text) != 0 &
                    Convert.ToInt32(textBox2.Text) != 0 &
                    comboBox.SelectedItem.ToString() != "")
                {
                    result = true;
                }
                else // нельзя записать количество 0 или по цене 0 или не выбрав материал
                {
                    result = false;                    
                }                
            }
            else  // нельзя записывать если не конвертируется количество и цена в числа
            {
                result = false;
            }

            return result;
        }

        /// <summary>
        /// Проверяет введены ли данные верно. 
        /// Значения текстбоксов должны конвертироваться в double и отличаться от 0, 
        /// Элемент в comboBox должен быть не пустым и должен быть выбран
        /// Возвращает true если данные введены верно
        /// </summary>
        /// <param name="textBox1"></param>
        /// <param name="textBox2"></param>
        /// <param name="comboBox"></param>
        /// <returns></returns>
        public static bool ProverkaDannih(TextBox textBox1, TextBox textBox2, TextBox textBox3)
        {          
            bool result = textBox1.Text != "" & textBox2.Text != "" & textBox3.Text != "";
            return result;
        }

        
        /*
        ///ПОПЫТКА ПОЛУЧИТЬ СПИСОК МАТЕРИАЛОВ ИЗ ОТДЕЛЬНОГО ФАЙЛА. 
        /// ПРОБЛЕМА ВОЗНИКЛА ПРИ СТИРАНИИ ЗНАЧЕНИЯ ИСПОЛЬЗОВАННОЙ СТРОКИ. 
        /// НЕЛЬЗЯ ИЗМЕНИТЬ ЗНАЧЕНИЕ ComboBox напрямую, если задано IteamSourse 
        */
        
        /// <summary>
        /// Возвращает коллекцию строк с наимованиями материалов
        /// </summary>
        /// <returns></returns>
        public static Collection<string> MaterialNamesCollection()
        {
            Collection<string> materialNameCollection = new Collection<string>();            

            string[] rawMaterialsStrings = File.ReadAllLines("Материалы.txt");

            for (int i = 0; i < rawMaterialsStrings.Length; i++)
            {
                // получаем номер позиции первого символа #  и второго
                int firstPosition = rawMaterialsStrings[i].IndexOf("#") + 1;
                int secondPosition = rawMaterialsStrings[i].LastIndexOf("#");
                // "вырезаем" имя файла из строки
                string newMaterialName = rawMaterialsStrings[i].Substring(firstPosition,
                    secondPosition - firstPosition);

                // добавляем имя файла в коллекцию
                materialNameCollection.Add(newMaterialName);
            }
            return materialNameCollection;
        }
          

        /// <summary>
        /// Отрезает символы "Руб" от Label.Content
        /// </summary>
        /// <param name="nameLabel"></param>
        /// <returns></returns>
        public static string CutStringRub(Label nameLabel)
        {
            string stringWithoutRub = String.Empty;
            string localSumm = Convert.ToString(nameLabel.Content);
 
            for (int i = 0; i < localSumm.IndexOf("руб"); i++)
            {
                stringWithoutRub += localSumm[i];
            }
            return stringWithoutRub;
        }

        /// <summary>
        /// Возвращет строку содержащую имя столбца в excele по номеру столбца
        /// </summary>
        /// <param name="column"> Номер столбца</param>
        /// <returns></returns>
        public static string Letter(int column)
        {
            string letter = "";
            const int ALPHABET_COUNT = 26;
            const int COD_A = 65;

            if (column <= ALPHABET_COUNT)
            {
                letter = "" + Convert.ToChar(COD_A + column - 1);
            }
            else
            {
                for (int i = 1; i < 200; i++)
                {
                    if (i * ALPHABET_COUNT + 1 <= column & column < (i + 1) * ALPHABET_COUNT + 1)
                    {
                        letter = "" + Convert.ToChar(COD_A + i - 1) + Convert.ToChar(COD_A + (column - i * ALPHABET_COUNT) - 1);
                    }
                }
            }
            return letter;
        }


        /// <summary>
        /// Cохранение Прихода материала в Excel
        /// </summary>
        /// <param name="excel"></param>
        /// <param name="arrayDebitMaterial"></param>
        /// <param name="myDataGrid"></param>
        public static void SaveDataDebit(string fileName, ObservableCollection<DebitMaterial> myCollection)
        {
            //TODO Убедиться в однозначности связывания Collections и DataGrid
            // TODO Возвращать bool с результатом сохранения

            // РАБОТА С ФАЙЛОМ
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook = excelApp.Workbooks.Open(fileName);
            Worksheet materialDebitWorksheet = (Worksheet) workbook.Sheets["Приход материалов"];
            Range debitRange = materialDebitWorksheet.UsedRange;
            Worksheet analiticaWorksheet = workbook.Sheets["Аналитика"];
            Worksheet priceDebitWorksheet = (Worksheet) workbook.Sheets["Цена прихода"];
            Worksheet creditWorksheet = workbook.Sheets["Расход материалов"];
            Range creditRange = creditWorksheet.UsedRange;
            Worksheet costsWorksheet = workbook.Sheets["Стоимость прихода"];

            //Получаем номер последней заполненной строки
            int lastrow = materialDebitWorksheet.UsedRange.Rows.Count;

            // формируем формулы сумм и записываем
            for (int column = 4; column <= debitRange.Columns.Count; column++)
            {
                #region Формирование и перезапись формул

                // стираем старые формулы 
                materialDebitWorksheet.Cells[lastrow, column] = "";
                priceDebitWorksheet.Cells[lastrow, column] = "";
                costsWorksheet.Cells[lastrow, column] = "";

                // формируем новую формулу
                string formula = "=СУММ(" + Letter(column) + "3:" + Letter(column) + lastrow + ")";

                // записываем новую формулу суммы по столбцам в соответствующие ячейки 
                // ...для Прихода материалов
                materialDebitWorksheet.Cells[lastrow + 1, column].FormulaLocal = formula;
                // ... для Цены прихода
                priceDebitWorksheet.Cells[lastrow + 1, column].FormulaLocal = formula;
                // ... для Стоимости 
                costsWorksheet.Cells[lastrow + 1, column].FormulaLocal = formula;

                // формируем формулу для Аналитики
                string analitica = "='Приход материалов'!" + Letter(column) + (lastrow + 1) + "-'Расход материалов'!" +
                                   Letter(column) + creditRange.Rows.Count;
                // записываем новую формулу в аналитику
                Range analiticaRange = (Range) analiticaWorksheet.Cells[5, column];
                analiticaRange.FormulaLocal = analitica;

                #endregion
            }
            
            // заполняем ячейки файла данными из формы о ПРИХОДЕ материала, ЦЕНЕ и СТОИМОСТИ
            // порядковый номер
            materialDebitWorksheet.Cells[lastrow, 1] = lastrow - 2;
            priceDebitWorksheet.Cells[lastrow, 1] = myCollection[0].Provider;
            costsWorksheet.Cells[lastrow, 1] = lastrow - 2;
            // дата 
            materialDebitWorksheet.Cells[lastrow, 2] = myCollection[0].Data;
            priceDebitWorksheet.Cells[lastrow, 2] = myCollection[0].Data;
            costsWorksheet.Cells[lastrow, 2] = myCollection[0].Data;
            // номер документа
            materialDebitWorksheet.Cells[lastrow, 3] = myCollection[0].DocumentNumber;
            priceDebitWorksheet.Cells[lastrow, 3] = myCollection[0].DocumentNumber;
            costsWorksheet.Cells[lastrow, 3] = myCollection[0].DocumentNumber;
            // количество материала цена и стоимость
            for (int i = 0; i < myCollection.Count; i++)
            {
                materialDebitWorksheet.Cells[lastrow, myCollection[i].MaterialIndex + 4] =
                    myCollection[i].Debit;
                priceDebitWorksheet.Cells[lastrow, myCollection[i].MaterialIndex + 4] =
                    myCollection[i].Price;
                costsWorksheet.Cells[lastrow, myCollection[i].MaterialIndex + 4].FormulaLocal =
                    myCollection[i].Debit * myCollection[i].Price;
            }

            // закрываем Excel
            workbook.Close(true, Missing.Value, Missing.Value);
            excelApp.Quit();
        }


 

        /// <summary>
        /// Сохраняет данные о раходе в Файл
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="myCollection"></param>
        public static void SaveDataCredit(string fileName, ObservableCollection<CreditMaterial> myCollection)
        {
            //TODO Убедиться в однозначности связывания Collections и DataGrid
            // TODO Возвращать bool с результатом сохранения
            
            Application excelApp = new Application();
            Workbook workbook = excelApp.Workbooks.Open(fileName);

            Worksheet creditWorksheet = workbook.Sheets["Расход материалов"];
            Range myRange = creditWorksheet.UsedRange;

            //Получаем номер последней заполненной строки
            int lastrow = creditWorksheet.UsedRange.Rows.Count;

            // формируем формулы сумм и записываем
            for (int column = 4; column <= myRange.Columns.Count; column++)
            {
                #region Формирование и перезапись формул

                // стираем старые формулы 
                creditWorksheet.Cells[lastrow, column] = String.Empty;

                // формируем новую формулу
                string formula = "=СУММ(" + Letter(column) + "3:" + Letter(column) + lastrow + ")";

                // записываем новую формулу суммы по столбцам для Прихода материалов
                creditWorksheet.Cells[lastrow + 1, column].FormulaLocal = formula;

                {
                    Worksheet materialDebitWorksheet = (Worksheet) workbook.Sheets["Приход материалов"];
                    Range debitRange = materialDebitWorksheet.UsedRange;

                    // формируем формулу для Аналитики
                    string analitica = "='Приход материалов'!" + Letter(column) + (debitRange.Rows.Count) +
                                       "-'Расход материалов'!" +
                                       Letter(column) + (myRange.Rows.Count + 1);

                    // записываем новую формулу в аналитику
                    Worksheet analiticaWorksheet = workbook.Sheets["Аналитика"];
                    Range analiticaRange = (Range)analiticaWorksheet.Cells[5, column];

                    analiticaRange.FormulaLocal = analitica;
                }

                #endregion
            }

            // заполняем ячейки файла данными из коллекции
            // порядковый номер
            creditWorksheet.Cells[lastrow, 1] = lastrow - 2;
            // дата 
            creditWorksheet.Cells[lastrow, 2] = myCollection[0].Data;
            // номер документа
            creditWorksheet.Cells[lastrow, 3] = myCollection[0].DocumentNumber;
            // количество материала 
            for (int i = 0; i < myCollection.Count; i++)
            {
                creditWorksheet.Cells[lastrow, myCollection[i].MaterialIndex + 4] = myCollection[i].Credit;
            }

            // закрываем Excel
            workbook.Close(true, Missing.Value, Missing.Value);
            excelApp.Quit();
        }

        


        /// <summary>
        /// Сравнивает значения Label.Content и подкрашивает label1 в зависимости от результата
        /// </summary>
        /// <param name="label1"></param>
        /// <param name="label2"></param>
        public static void ColorSumm(System.Windows.Controls.Label label1, System.Windows.Controls.Label label2)
        {
            // подкращиваем введенную сумму
            Brush newBrush = Brushes.Yellow;

            if (label1.Content.ToString() == label2.Content.ToString())
                newBrush = Brushes.LawnGreen;
            if (Convert.ToInt32(label1.Content) > Convert.ToInt32(label2.Content))
                newBrush = Brushes.Red;

            label1.Background = newBrush;
        }
    }
}
