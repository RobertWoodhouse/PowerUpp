﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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

namespace PowerUpp
{
    /// <summary>
    /// Interaction logic for TableView.xaml
    /// </summary>
    public partial class TableView : Page
    {
        public TableView()
        {
            //SelectionController.StartExcelAppAsync(); //TEST

            InitializeComponent();
            lblHeaderEx.Content = SelectionView.ExerciseTitle.Replace("_", " ");
            
            TableController tableData = new TableController();
            tableData.OpenExcelWorksheetAsync();
            this.dgTable.DataContext = tableData; // Load data from spreadsheet into exercises table
            this.dgExTable.DataContext = tableData; // Load data from spreadsheet into specicic exercises
        }

        private void btnPrevious_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Content = new SelectionView();
        }

        private void btnNext_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Content = new ChartView();
        }
    }
}
