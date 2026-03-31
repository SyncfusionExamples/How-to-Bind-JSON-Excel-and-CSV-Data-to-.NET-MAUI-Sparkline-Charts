using BindDataFromDiffSource.ViewModels;

namespace BindDataFromDiffSource
{
    public partial class MainPage : ContentPage
    {
        public MainPage()
        {
            InitializeComponent();
            
            // Set the ViewModel as BindingContext
            BindingContext = new MainViewModel();
        }
    }
}
