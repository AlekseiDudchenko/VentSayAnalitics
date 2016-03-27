using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using CreditApp;
using Microsoft.Office.Interop.Excel;
using Window = System.Windows.Window;

namespace VentSayAnalitics
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        ExcelClass excel = new ExcelClass();

        AnaliticsFunktions newAnaliticsFunktions = new AnaliticsFunktions();

        public MainWindow()
        {
                    
            InitializeComponent();

            excel.GetMaterials();

            // зполняем listBox названиями материалов
            for (int i = 0; i < excel.MaterialsNames.Length; i++)
            {
                if (excel.MaterialsNames[i] != null)
                {
                    ListBox.Items.Add(excel.MaterialsNames[i]);
                }
            }

            newAnaliticsFunktions.OpenFile();

        }

        ObservableCollection<AnaliticsMaterial> analiticsMaterialsCollection = new ObservableCollection<AnaliticsMaterial>();

        private void ListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            
            

            
            ObservableCollection<AnaliticsMaterial> myAnaliticsMaterialsDebitCollection = newAnaliticsFunktions.GetMaterialsDebitCollection(ListBox.SelectedIndex);
            ObservableCollection<AnaliticsMaterial> myAnaliticsMaterialsCreditCollection = newAnaliticsFunktions.GetMaterialsCreditCollection(ListBox.SelectedIndex);
            //newAnaliticsFunktions.CloseFile();

            DebitDataGrid.ItemsSource = myAnaliticsMaterialsDebitCollection;
            CreditDataGrid.ItemsSource = myAnaliticsMaterialsCreditCollection;
            //CreditDataGrid.Items.Refresh();



            //NameMaterialLabel.Content = myAnaliticsMaterialsCollection[0].MaterialName;
            NameMaterialLabel.Content = excel.MaterialsNames[ListBox.SelectedIndex];
            if (myAnaliticsMaterialsDebitCollection.Count != 0)
                BalanceLabel.Content = myAnaliticsMaterialsDebitCollection[0].Balance;
            else           
                BalanceLabel.Content = "неизвестно";
 
           
            UnitLabel.Content = excel.Units[ListBox.SelectedIndex];





        }

        private void Button_Click_1(object sender, System.Windows.RoutedEventArgs e)
        {
            newAnaliticsFunktions.CloseFile();
            this.Close();
        }
    }
}
