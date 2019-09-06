using System;
using System.Windows;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Controls;
using Xceed.Wpf.Toolkit;
using System.Windows.Media;

namespace CustomReport2020
{
    class UIMethods
    {
        public static void AddGroup(Grid groups)
        {
            Random rnd = new Random();

            RowDefinition gridRow = new RowDefinition();
            groups.RowDefinitions.Add(gridRow);

            Grid group = new Grid();

            group.Name = "GroupUI" + groups.Children.Count;
            group.Margin = new Thickness(0, 0, 0, 5);

            ColumnDefinition gridCol1 = new ColumnDefinition();
            gridCol1.Width = new GridLength(1, GridUnitType.Star);

            ColumnDefinition gridCol2 = new ColumnDefinition();
            gridCol2.Width = new GridLength(2, GridUnitType.Star);

            group.ColumnDefinitions.Add(gridCol1);
            group.ColumnDefinitions.Add(gridCol2);

            System.Windows.Controls.TextBox userTextInput = new System.Windows.Controls.TextBox();
            userTextInput.Name = "UserInput" + groups.Children.Count;
            userTextInput.HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch;
            userTextInput.Text = "Group Name " + groups.Children.Count;
            Grid.SetColumn(userTextInput, 0);

            ColorPicker userColorInput = new ColorPicker();
            userColorInput.Name = "UserColor" + groups.Children.Count;
            userColorInput.HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch;
            userColorInput.DisplayColorTooltip = true;
            userColorInput.AvailableColorsSortingMode = ColorSortingMode.HueSaturationBrightness;
            Color randomColor = Color.FromRgb((byte)rnd.Next(256), (byte)rnd.Next(256), (byte)rnd.Next(256));
            userColorInput.SelectedColor = randomColor;
            Grid.SetColumn(userColorInput, 1);

            group.Children.Add(userTextInput);
            group.Children.Add(userColorInput);

            groups.Children.Add(group);
            Grid.SetRow(group, groups.RowDefinitions.Count - 1);
        }

        public static void DeleteGroup(Grid groups)
        {
            // A List To Store UI Elements To Remove From The Controller Window
            List<UIElement> elementsToRemove = new List<UIElement>();

            if (groups.RowDefinitions.Count > 1)
            {
                ((Grid)groups.Children[groups.Children.Count - 1]).Children.Clear();

                groups.Children.RemoveAt(groups.Children.Count - 1);
                groups.RowDefinitions.RemoveAt(groups.RowDefinitions.Count - 1);

            }
            else
            {
                return;
            }
        }
    }
}
