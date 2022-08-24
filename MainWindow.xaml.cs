using ExcelDataReader;
using DutyArranger.Source;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Input;
using DutyArranger.Source.Entities;
using System.Collections.ObjectModel;

namespace DutyArranger
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            DataContext = new ViewModel();
        }

        private void questionBox1_MouseEnter(object sender, MouseEventArgs e)
        {
            questionBox1Tooltip.Visibility = Visibility.Visible;
        }

        private void questionBox2_MouseEnter(object sender, MouseEventArgs e)
        {
            questionBox2Tooltip.Visibility = Visibility.Visible;
        }

        private void questionBox1_MouseLeave(object sender, MouseEventArgs e)
        {
            questionBox1Tooltip.Visibility = Visibility.Hidden;
        }

        private void questionBox2_MouseLeave(object sender, MouseEventArgs e)
        {
            questionBox2Tooltip.Visibility = Visibility.Hidden;
        }

        private void questionBox3_MouseEnter(object sender, MouseEventArgs e)
        {
            questionBox3Tooltip.Visibility = Visibility.Visible;
        }

        private void questionBox3_MouseLeave(object sender, MouseEventArgs e)
        {
            questionBox3Tooltip.Visibility = Visibility.Hidden;
        }
    }
}
