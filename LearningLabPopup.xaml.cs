using System;
using System.Collections.Generic;
using System.ComponentModel;
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

namespace WPF_Language_Lab_project
{
    public partial class LearningLabActivity : Window
    {
        public event Action<bool> ClosedByUser;

        public LearningLabActivity()
        {
            InitializeComponent();
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            ClosedByUser?.Invoke(true);
            base.OnClosing(e);
        }

        private void translateButton_Click(object sender, RoutedEventArgs e)
        {
            // Get the user's input from the TextBox
            string userInput = inputTextBox.Text;

            // Call your translation API here (replace with actual API call)
            string translatedText = "Translated text";

            // Update the Label with the translated text
            outputLabel.Content = translatedText;
        }

        // Dummy translation function (replace with actual API call)
        private string Translate(string input)
        {
            return input + " (translated)";
        }
    }

}
