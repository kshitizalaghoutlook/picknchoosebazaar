using Google;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace PickNChooseBazaar
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static readonly string ApplicationName = "PickNCgooseBazaar";
        static readonly string SpreadsheetId = "10AD63XlQi9x4IvkiNNI4oWaozWe8Pg1BzOVy58cgXLo";
        static readonly string sheet = "Products";
        private const double HstRateOntario = 0.13;

        static SheetsService service;

        public static SheetsService GetSheetsService()
        {
            UserCredential credential;

            using (var stream = new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.FromStream(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
            }

            // Create Google Sheets API service.
            service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            return service;
        }

        public static List<ProductDataModel> ReadEntries()
        {
            var data = new List<ProductDataModel>();
            try
            {
              
                var range = $"{sheet}!A2:J";
                SpreadsheetsResource.ValuesResource.GetRequest request =
                        service.Spreadsheets.Values.Get(SpreadsheetId, range);

                ValueRange response = request.Execute();
                IList<IList<object>> values = response.Values;
                if (values != null && values.Count > 0)
                {
                    foreach (var row in values)
                    {
                        if (row.Count > 0)
                        {
                            ProductDataModel p = new ProductDataModel();

                            // Print columns A and B, which correspond to indices 0 and 1.
                            data.Add(new ProductDataModel
                            {
                                ProductId = Convert.ToInt32(row[0]),
                                Name = row[1].ToString(),
                                Description = row[2].ToString(),
                                CostPrice = string.IsNullOrEmpty(row[3].ToString()) ? 0 : Convert.ToInt32(row[3]),
                                SellingPrice = string.IsNullOrEmpty(row[4].ToString()) ? 0 : Convert.ToInt32(row[4]),
                                Category = row[5].ToString(),
                                Quantity = string.IsNullOrEmpty(row[6].ToString()) ? 0 : Convert.ToInt32(row[6]),
                                Manufacturer = row[7].ToString(),
                                Brand = row[8].ToString(),
                                PalletId = string.IsNullOrEmpty(row[9].ToString()) ? 0 : Convert.ToInt32(row[9])

                            });
                            Console.WriteLine($"{row[0]}, {row[1]}");
                        }
                    }
                }
                else
                {
                    Console.WriteLine("No data found.");
                }
                return data;
            }
            catch (GoogleApiException e)
            {
                Console.WriteLine($"Error: {e.Message}");
            }
            return data;
        }
        public static  void LoadData()
        {
            var range = $"{sheet}!A1:B5";
            //var service = new MainWindow();
            var data =  ReadEntries();
           // service.MyDataGrid.ItemsSource = data;
        }


        public  MainWindow()
        {
            InitializeComponent();

            comboBoxCategory.Items.Add("Furniture");
            comboBoxCategory.Items.Add("Appliance");
            comboBoxCategory.Items.Add("Car");
            comboBoxCategory.Items.Add("Other");

            comboBoxStore.Items.Add("Liquidation.com");
            comboBoxStore.Items.Add("Amazon");
            comboBoxStore.Items.Add("Bid416.com");
            comboBoxStore.Items.Add("Other");

            GetSheetsService();
           
            MyDataGrid.ItemsSource = ReadEntries(); 
            //LoadData();
        }
        private void Add_Click(object sender, RoutedEventArgs e)
        {
            double costPrice ;
            int quantity ;
            int palletId;
            

            if (string.IsNullOrWhiteSpace(txtName.Text))
            {
                MessageBox.Show("Name is required.");
                return;
            }
            if (!Double.TryParse(txtPrice.Text, out costPrice))
            {
                MessageBox.Show("Cost Price should be a number");
                return;
            }
           
            if (!int.TryParse(txtQuantity.Text, out quantity))
            {
                MessageBox.Show("Quantity should be a number");
                return;
            }
            if (!int.TryParse(txtPalletId.Text, out quantity))
            {
                MessageBox.Show("Pallet Id should be a number");
                return;
            }
            var Range = "A:A";
            // Read the last ID from the sheet
            var rangeRequest = service.Spreadsheets.Values.Get(SpreadsheetId, Range);
            var response = rangeRequest.Execute();
            var values = response.Values;
            int lastId = 0;
            if (values != null && values.Count > 0)
            {
                foreach (var row in values)
                {
                    if (row.Count > 0 && int.TryParse(row[0].ToString(), out int id))
                    {
                        lastId = id;
                    }
                }
            }

            // Increment the ID for the new row
            int newId = lastId + 1;


            // Logic for processing Form 1 data
            // ...
            string name = txtName.Text;
            string descripton = txtDescription.Text;
            costPrice =string.IsNullOrEmpty(txtPrice.Text) ? 0 : Convert.ToInt32(txtPrice.Text);
          
            string category = comboBoxCategory.Text;
             quantity = Convert.ToInt32(txtQuantity.Text);
            string manufacturer = txtManufacturer.Text;
            string store = comboBoxStore.Text;
            palletId = Convert.ToInt32(txtPalletId.Text);
            List<object> data = new List<object> { newId, name, descripton, costPrice, "", category, quantity, manufacturer, store, palletId };
            var range = $"{sheet}!A2:I2";  // Update your sheet name and range
            var valueRange = new ValueRange();
            valueRange.Values = new List<IList<object>> { data };

            UpdateGoogleSheet(service, SpreadsheetId, range, valueRange);
            MessageBox.Show("Product added!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
            ClearFormFields();
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            // Logic for processing Form 2 data
            // ...
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            // Logic for processing Form 2 data
            // ...

            ClearFormFields();
        }

        private void UpdateGoogleSheet(SheetsService service, string spreadsheetId, string range, ValueRange valueRange)
        {
            var appendRequest = service.Spreadsheets.Values.Append(valueRange, spreadsheetId, range);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            var appendResponse = appendRequest.Execute();

            // Optionally handle the response, e.g., to check if the operation was successful
        }

        private void ClearFormFields()
        {
            // Clear all TextBoxes
            foreach (var control in AddProduct.Children)
            {
                if (control is TextBox textBox)
                {
                    textBox.Clear();
                }
                // Add similar blocks for other control types like ComboBox, CheckBox, etc.
                // For example, if you have ComboBoxes, you can reset them to the default index
                else if (control is ComboBox comboBox)
                {
                    comboBox.SelectedIndex = -1; // This sets the ComboBox to have no selection
                }
            }
        }

        private void RefreshButton_Click(object sender, RoutedEventArgs e)
        {
            MyDataGrid.ItemsSource = null; ; // Clear existing items
            MyDataGrid.ItemsSource = ReadEntries(); // Reload or refresh data
        }
    }

}

    public class ProductDataModel
    {
        public int? ProductId { get; set; }
        public string? Name { get; set; }
        public string? Description { get; set; }
        public int? CostPrice { get; set; }
        public int? SellingPrice { get; set; }
        public string? Category { get; set; }
        public int? Quantity { get; set; }
        public string? Manufacturer { get; set; }
        public string? Brand { get; set; }
        public int? PalletId { get; set; }
        // Add more properties as per your Google Sheets columns
    }


   
