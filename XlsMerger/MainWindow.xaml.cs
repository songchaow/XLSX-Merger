using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace XlsMerger
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Clear_Items(object sender, RoutedEventArgs e)
        {
            inputfiles.Items.Clear();
        }

        private void Remove_Items(object sender, RoutedEventArgs e)
        {
            var selected = inputfiles.SelectedItems;
            while(inputfiles.SelectedItems.Count>0)
                inputfiles.Items.Remove(inputfiles.SelectedItems[0]);
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog fileDialog = new Microsoft.Win32.OpenFileDialog
            {
                DefaultExt = ".xlsx",
                Filter = "电子表格文件|*.xlsx;*.xls",
                Multiselect = true
            };
            Nullable<bool> result = fileDialog.ShowDialog();

            if (result == true)
            {
                // prepare for the same menu
                System.Windows.Controls.ContextMenu menu = new System.Windows.Controls.ContextMenu();
                System.Windows.Controls.MenuItem clear = new System.Windows.Controls.MenuItem
                {
                    Header = "清空所有项",
                    IsCheckable = true
                };
                clear.Click += Clear_Items;
                System.Windows.Controls.MenuItem remove = new System.Windows.Controls.MenuItem
                {
                    Header = "移除所选项",
                    IsCheckable = true
                };
                remove.Click += Remove_Items;
                menu.Items.Add(clear);
                menu.Items.Add(remove);
                // Open document
                string[] filenames = fileDialog.FileNames;
                foreach (string filename in filenames)
                {
                    TextBlock item = new TextBlock
                    {
                        Text = filename,

                        ContextMenu = menu
                    };
                    inputfiles.Items.Add(item);

                }
                
            }
        }

        private void Output_Button_Click(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            DialogResult result = dialog.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
                outputpath.Text = dialog.SelectedPath;

        }

        private readonly string[] commaSeparators = { ",", "，" };
        private readonly string[] dashSeparators = { "-", "—" };
        private void showParseInfo(string messageBoxText= "所填的数据格式不正确。",string caption="嘿！")
        {
            MessageBoxButton button = MessageBoxButton.OK;
            MessageBoxImage icon = MessageBoxImage.Information;
            System.Windows.MessageBox.Show(messageBoxText, caption, button, icon);
        }
        private List<int> parseMultiNumbers(string input)
        {
            string[] raw_rows = input.Split(commaSeparators, StringSplitOptions.RemoveEmptyEntries);
            List<int> rows = new List<int>();
            foreach(var group in raw_rows)
            {
                string[] parsed_rows = group.Split(dashSeparators, StringSplitOptions.RemoveEmptyEntries);
                if(parsed_rows.Length>2)
                {
                    showParseInfo();
                    return null;
                }
                if (parsed_rows.Length == 1)
                {
                    if (int.TryParse(parsed_rows[0], out int item))
                    {
                        if (rows.IndexOf(item) < 0)
                            rows.Add(item);
                    }
                    else
                    {
                        showParseInfo(); return null;
                    }
                }
                else if (int.TryParse(parsed_rows[0], out int start) &&
                    int.TryParse(parsed_rows[1], out int end))
                {
                    for (int i = start; i <= end; i++)
                        if (rows.IndexOf(i) < 0)
                            rows.Add(i);
                }
                else
                {
                    showParseInfo();
                    return null;
                }
            }
            return rows;
        }

        private bool Check_Value()
        {
            // first if they are filled with values.
            if (publicrow_input.Text.Length == 0 || mergerow_input.Text.Length == 0)
            {
                showParseInfo("还有空没有填呢。。");
                return false;
            }
            List<int> public_rows = parseMultiNumbers(publicrow_input.Text);
            if (public_rows == null)
                return false;
            List<int> main_rows = parseMultiNumbers(mergerow_input.Text);
            if (main_rows == null)
                return false;
            /*if (int.TryParse(mergerow_input.Text, out int main_row) == false)
            {
                showParseInfo("合并行数填写格式不正确。");
                return false;
            }*/
            if (int.TryParse(sheetseq.Text, out int sheet_seq) == false)
            {
                showParseInfo("工作表序号填写不正确。");
                return false;
            }
            List<string> input_files = new List<string>();
            foreach (var item in inputfiles.Items)
                input_files.Add((item as TextBlock).Text);
            Merge mergeobj = new Merge();
            mergeobj.Merge_xlsx(input_files, outputpath.Text, filename.Text,public_rows, main_rows, sheet_seq,(bool)ignoreEmptyBox.IsChecked);

            return true;

        }

        private void Merge_Button_Click(object sender, RoutedEventArgs e)
        {
            if(Check_Value())
                showParseInfo("合并成功！");
        }
    }
}
