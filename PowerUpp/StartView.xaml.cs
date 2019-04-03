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
    /// Interaction logic for StartView.xaml
    /// </summary>
    public partial class StartView : Page
    {
        public StartView()
        {
            InitializeComponent();
        }

        private void btnLoadXl_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Content = new SelectionView();
            SelectionController.loadFile = true;
        }

        private void btnNewXl_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Content = new SelectionView();
            SelectionController.loadFile = false;
        }
    }
}
