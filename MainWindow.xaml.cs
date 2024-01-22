using Google;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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

        public static void ReadEntries()
        {
            try
            {
                var range = $"{sheet}!A1:B5";
                SpreadsheetsResource.ValuesResource.GetRequest request =
                        service.Spreadsheets.Values.Get(SpreadsheetId, range);

                ValueRange response = request.Execute();
                IList<IList<object>> values = response.Values;
                if (values != null && values.Count > 0)
                {
                    foreach (var row in values)
                    {
                        // Print columns A and B, which correspond to indices 0 and 1.
                        Console.WriteLine($"{row[0]}, {row[1]}");
                    }
                }
                else
                {
                    Console.WriteLine("No data found.");
                }
            }
            catch (GoogleApiException e)
            {
                Console.WriteLine($"Error: {e.Message}");
            }
        }


        public MainWindow()
        {
            InitializeComponent();

            GetSheetsService();
            ReadEntries();
        }
        private void Add_Click(object sender, RoutedEventArgs e)
        {
            // Logic for processing Form 1 data
            // ...
            string name = txtName.Text;
            string descripton = txtDescription.Text;
            double costPrice = Convert.ToDouble(txtPrice.Text);
            string category = comboBoxCategory.SelectedValue.ToString();
            int quantity = Convert.ToInt32(txtQuantity.Text);

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
        }

    }
}
