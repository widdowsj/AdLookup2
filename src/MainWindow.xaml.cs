using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.DirectoryServices;
using System.Linq;
using System.Security.Principal;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace AdLookup2
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private SortedList<string, string> translationMatrix;
        private Stack<string> backStack = new Stack<String>();
        private string currentDn;
        private readonly Brush _headingBrush = new LinearGradientBrush(Colors.WhiteSmoke, Colors.LightGray, 0);
        private readonly Brush _alternateRowBrush = new LinearGradientBrush(Color.FromRgb(0xe2, 0xe2, 0xe2), Colors.Transparent, 0);

        public MainWindow()
        {
            InitializeComponent();
        }

        private void SearchCommand_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = true;
        }

        private void SearchCommand_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            this.Cursor = Cursors.Wait;
            try
            {
                if (fieldComboBox.Text == "")
                    fieldComboBox.Text = "cn";
                DirectoryEntry rootDSE = new DirectoryEntry(String.Format("LDAP://{0}/rootDSE", domainTextBox.Text));
                DirectoryEntry root = new DirectoryEntry("LDAP://" + PropString(rootDSE.Properties["defaultNamingContext"]));
                DirectorySearcher ds = new DirectorySearcher(root);
                ds.Filter = String.Format("({0})", TranslateToAttribute(fieldComboBox.Text) + "=" + searchTextBox.Text);
                ds.PropertiesToLoad.Add("ADSPath");
                ds.PropertiesToLoad.Add("cn");
                ds.PropertiesToLoad.Add("thumbnailPhoto");
                ds.PropertiesToLoad.Add("jpegPhoto");
                ds.PropertiesToLoad.Add("GivenName");
                ds.PropertiesToLoad.Add("sn");
                SearchResultCollection users = ds.FindAll();

                searchListView.Items.Clear();

                foreach (SearchResult user in users)
                {
                    BitmapImage photoBitmap = null;
                    if (user.Properties["jpegPhoto"].Count > 0)
                    {
                        byte[] photoBytes = user.Properties["jpegPhoto"][0] as byte[];
                        System.IO.MemoryStream ms = new System.IO.MemoryStream(photoBytes);
                        try
                        {
                            photoBitmap = new BitmapImage();
                            photoBitmap.BeginInit();
                            photoBitmap.StreamSource = ms;
                            photoBitmap.DecodePixelWidth = 50;
                            photoBitmap.EndInit();
                        }
                        catch
                        { }
                    }
                    if (user.Properties["thumbnailPhoto"].Count > 0 && photoBitmap == null)
                    {
                        byte[] photoBytes = user.Properties["thumbnailPhoto"][0] as byte[];
                        System.IO.MemoryStream ms = new System.IO.MemoryStream(photoBytes);
                        try
                        {
                            photoBitmap = new BitmapImage();
                            photoBitmap.BeginInit();
                            photoBitmap.StreamSource = ms;
                            photoBitmap.DecodePixelWidth = 50;
                            photoBitmap.EndInit();
                        }
                        catch
                        { }
                    }
                    User u = new User()
                    {
                        cn = PropString(user.Properties["cn"], " "),
                        name = PropString(user.Properties["sn"], " ") + ", " + PropString(user.Properties["GivenName"], " "),
                        adsPath = PropString(user.Properties["ADSPath"]),
                        photo = photoBitmap
                    };
                    searchListView.Items.Add(u);
                }
                if (users.Count == 1)
                    searchListView.SelectedItem = searchListView.Items[0];
            }
            catch (Exception ex)
            {
                CurrentDnLabel.Text = "Error";
                propertiesRichTextBox.Document = new FlowDocument(new Paragraph(new Run(ex.ToString())));
            }
            this.Cursor = Cursors.Arrow;
        }

        private string PropString(System.Collections.ICollection props, string separator = "\r\n")
        {
            StringBuilder sb = new StringBuilder();

            if (props != null)
            {
                foreach (dynamic value in props)
                {
                    if (sb.Length > 0)
                        sb.Append(separator);
                    Object o = value;
                    if (o.ToString() == "System.__ComObject")
                    {
                        try
                        {
                            // Try an ADSI Large Integer
                            long dateValue = (value.HighPart * 100000000) + value.LowPart;
                            DateTime dt = DateTime.FromFileTime(dateValue);
                            // If Year(dt) = 1601 Then
                            //      sWork = dt.ToString("HH:mm")
                            // Else
                            sb.Append(dt.ToString("dd-MMM-yyyy HH:mm"));
                            // End If
                        }
                        catch
                        {
                            sb.Append(o.ToString());
                        }
                    }
                    else
                        sb.Append(o.ToString());
                }
            }
            return sb.ToString();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            searchTextBox.Text = "";
            propertiesRichTextBox.Document = new FlowDocument();
            LoadTranslationMatrix();
            fieldComboBox.Text = "sAMAccountName";
            fieldComboBox.ItemsSource = translationMatrix.Values;
            domainTextBox.Text = Environment.UserDomainName;
            searchTextBox.Focus();
        }


        private void LoadTranslationMatrix()
        {
            translationMatrix = new SortedList<string, string>();

            translationMatrix.Add("distinguishedname", "Distinguished Name");
            translationMatrix.Add("employeeid", "Emp ID");
            translationMatrix.Add("cn", "NT User ID");
            translationMatrix.Add("mail", "Email Address");
            translationMatrix.Add("personaltitle", "Personal Title");
            translationMatrix.Add("givenname", "Given Name");
            translationMatrix.Add("middlename", "Middle Name");
            translationMatrix.Add("sn", "Surname");
            translationMatrix.Add("displayname", "Display Name");
            translationMatrix.Add("streetaddress", "Street Address");
            translationMatrix.Add("l", "City");
            translationMatrix.Add("st", "State");
            translationMatrix.Add("postalcode", "Post Code");
            translationMatrix.Add("physicaldeliveryofficename", "Mail Address");
            translationMatrix.Add("telephonenumber", "Phone Number");
            translationMatrix.Add("mobile", "Mobile Phone");
            translationMatrix.Add("facsimiletelephonenumber", "Fax Number");
            translationMatrix.Add("title", "Job Title");
            translationMatrix.Add("memberof", "Groups");
            translationMatrix.Add("profilepath", "Profile Path");
            translationMatrix.Add("info", "Alert");
            translationMatrix.Add("sAMAccountName", "SAM Account Name");
        }

        private string TranslateFromAttribute(string name)
        {
            if (translationMatrix.ContainsKey(name))
                return translationMatrix[name];
            else
                return name;
        }

        private string TranslateToAttribute(string name)
        {
            string val = name;

            if (translationMatrix.ContainsValue(name))
                foreach (var key in translationMatrix.Keys)
                    if (translationMatrix[key] == name)
                        val = key;
            return val;
        }

        private void searchListView_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (e.AddedItems.Count > 0)
            {
                User u = (User)e.AddedItems[0];

                Title = "AD Lookup - " + u.cn;

                DisplayEntry(u.adsPath, true);
            }
        }

        private void DisplayEntry(string dn, bool addHistory)
        {
            DirectoryEntry de = new DirectoryEntry(dn);

            if (addHistory && !string.IsNullOrWhiteSpace(currentDn))
            {
                backStack.Push(currentDn);
                CommandManager.InvalidateRequerySuggested();
            }
            currentDn = dn;

            CurrentDnLabel.Text = PropString(de.Properties["cn"]);

            Table table = new Table();
            table.Columns.Add(new TableColumn {Width = new GridLength(300)});
            table.Columns.Add(new TableColumn {Width = new GridLength(1, GridUnitType.Auto)});

            table.RowGroups.Add(new TableRowGroup());
            FlowDocument doc = new FlowDocument(table);

            SortedSet<string> propNames = new SortedSet<string>();
            foreach (string name in de.Properties.PropertyNames)
                propNames.Add(TranslateFromAttribute(name));

            propertiesGrid.RowDefinitions.Clear();
            propertiesGrid.Children.Clear();
            double maxCellWidth = 0;
            bool odd = false;
            foreach (var name in propNames)
            {
                string val = name;
                if (name != TranslateToAttribute(name))
                    val += " (" + TranslateToAttribute(name) + ")";
                val += ":";
                TableRow tr = new TableRow();
                var run = new Bold(new Run(val));
                var block = new TextBlock(run);
                block.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity));
                var cellWidth = block.DesiredSize.Width;
                TableCell tc = new TableCell(new Paragraph(run));
                tc.Background = _headingBrush;
                if (cellWidth > maxCellWidth)
                    maxCellWidth = cellWidth; 
                tr.Cells.Add(tc);
                ICollection props;
                try
                {
                    props = de.Properties[TranslateToAttribute(name)];
                }
                catch (Exception e)
                {
                    props = new List<string> { e.ToString() };
                }
                var content = new TableCell(PropParagraph(val, props));
                if (odd)
                    content.Background = _alternateRowBrush;
                odd = !odd;
                tr.Cells.Add(content);
                table.RowGroups[0].Rows.Add(tr);

                AddProp(val, props);
            }
            table.Columns[0].Width = new GridLength(maxCellWidth + 10);
            propertiesRichTextBox.IsDocumentEnabled = true;
            propertiesRichTextBox.Document = doc;
        }

        private void AddProp(string name, System.Collections.ICollection props)
        {
            propertiesGrid.RowDefinitions.Add(new RowDefinition());
            TextBlock tb = new TextBlock();
            tb.SetValue(Grid.RowProperty, propertiesGrid.RowDefinitions.Count - 1);
            tb.SetValue(Grid.ColumnProperty, 0);
            tb.Text = name;
            tb.FontWeight = FontWeights.Bold;
            tb.Margin = new Thickness(5, 1, 5, 1);
            propertiesGrid.Children.Add(tb);

            StackPanel sp = new StackPanel();
            sp.SetValue(Grid.RowProperty, propertiesGrid.RowDefinitions.Count - 1);
            sp.SetValue(Grid.ColumnProperty, 1);
            propertiesGrid.Children.Add(sp);

            if (props != null)
            {
                foreach (dynamic value in props)
                {
                    string valueString = ((Object)value).ToString();
                    Image img = null;
                    Hyperlink h = null;
                    if (name == "userAccountControl:")
                        valueString = GetUserAccountControl(value);
                    else if (valueString == "System.__ComObject")
                        try
                        {
                            // Try an ADSI Large Integer
                            Object o = value;
                            Object low = o.GetType().InvokeMember("LowPart", System.Reflection.BindingFlags.GetProperty, null, o, null);
                            Object high = o.GetType().InvokeMember("HighPart", System.Reflection.BindingFlags.GetProperty, null, o, null);

                            //long dateValue = (value.HighPart << 32) + value.LowPart;
                            long dateValue = ((long)((int)high) << 32) + (long)((int)low);
                            DateTime dt = DateTime.FromFileTime(dateValue);
                            if (dt.ToString("dd-MMM-yyyy HH:mm") != "01-Jan-1601 11:00")
                                if (dt.Year == 1601)
                                    valueString = dt.ToString("HH:mm");
                                else
                                    valueString = dt.ToString("dd-MMM-yyyy HH:mm");
                        }
                        catch { }
                    else if (valueString == "System.Byte[]")
                    {
                        byte[] bytes = value as byte[];
                        if (bytes.Length == 16)
                        {
                            Guid guid = new Guid(bytes);
                            valueString = guid.ToString("B");
                        }
                        else if (bytes.Length == 28)
                        {

                            try
                            {
                                SecurityIdentifier sid = new SecurityIdentifier(bytes, 0);
                                valueString = sid.ToString();
                            }
                            catch
                            {   
                            }
                        }
                        else
                        {
                            System.IO.MemoryStream ms = new System.IO.MemoryStream(bytes);
                            try
                            {
                                BitmapImage photoBitmap = new BitmapImage();
                                photoBitmap.BeginInit();
                                photoBitmap.StreamSource = ms;
                                photoBitmap.EndInit();
                                img = new Image();
                                img.Source = photoBitmap;
                                img.Stretch = Stretch.None;
                            }
                            catch
                            {
                                img = null;
                            }
                        }
                    }
                    else if (valueString.ToString().StartsWith("CN=") || valueString.ToString().ToLower().StartsWith("http://"))
                    {
                        //string display = Regex.Match(valueString + ",", "CN=(.*?),").Groups[1].Captures[0].Value;
                        //h = new Hyperlink(new Run(display));
                        h = new Hyperlink(new Run(valueString));
                        h.NavigateUri = new Uri(valueString, valueString.ToLower().StartsWith("http://") ? UriKind.Absolute : UriKind.Relative);
                        h.Click += new RoutedEventHandler(HyperlinkClicked);
                        //p.TextIndent = -20;
                        //p.Margin = new Thickness(20, p.Margin.Top, p.Margin.Right, p.Margin.Bottom);
                    }
                    UIElement valueElement;
                    if (img != null)
                    {
                        valueElement = img;
                    }
                    else if (h != null)
                    {
                        valueElement = new TextBlock(h);
                    }
                    else
                    {
                        valueElement = new TextBox()
                        {
                            Text = valueString,
                            Style = FindResource("FauxLabel") as Style,
                        };
                    }
                    valueElement.SetValue(MarginProperty, new Thickness(5, 1, 5, 1));
                    sp.Children.Add(valueElement);
                }
            }
        }

        private string GetUserAccountControl(int value)
        {
            string userAccountControl = value.ToString();
            foreach (int val in Enum.GetValues(typeof(Enums.ADS_USER_FLAG_ENUM)))
            {
                if ((value & val) != 0)
                    userAccountControl += ", " + Enum.GetName(typeof(Enums.ADS_USER_FLAG_ENUM), val);
            }
            return userAccountControl;
        }

        private Paragraph PropParagraph(string name, System.Collections.ICollection props)
        {
            Paragraph p = new Paragraph();
            bool appendSeparator = false;

            if (props != null)
            {
                foreach (dynamic value in props.Cast<object>().OrderBy(x => x.ToString()))
                {
                    if (appendSeparator)
                        p.Inlines.Add("\r\n");
                    else
                        appendSeparator = true;
                    string valueString = ((Object)value).ToString();
                    if (name == "userAccountControl:")
                        p.Inlines.Add(new Run(GetUserAccountControl(value)));
                    else if (valueString == "System.__ComObject")
                        try
                        {
                            // Try an ADSI Large Integer
                            Object o = value;
                            Object low = o.GetType().InvokeMember("LowPart", System.Reflection.BindingFlags.GetProperty, null, o, null);
                            Object high = o.GetType().InvokeMember("HighPart", System.Reflection.BindingFlags.GetProperty, null, o, null);

                            //long dateValue = (value.HighPart * &H100000000) + value.LowPart;
                            long dateValue = ((long)((int)high) << 32) + (long)((int)low);
                            DateTime dt = DateTime.FromFileTime(dateValue);
                            if (dt.ToString("dd-MMM-yyyy HH:mm") != "01-Jan-1601 11:00")
                                if (dt.Year == 1601)
                                    p.Inlines.Add(dt.ToString("HH:mm"));
                                else
                                    p.Inlines.Add(dt.ToString("dd-MMM-yyyy HH:mm"));
                        }
                        catch
                        {
                            p.Inlines.Add(valueString);
                        }
                    else if (valueString == "System.Byte[]")
                    {
                        byte[] bytes = value as byte[];
                        if (bytes.Length == 16)
                        {
                            Guid guid = new Guid(bytes);
                            p.Inlines.Add(guid.ToString("B"));
                        }
                        else if (bytes.Length == 28)
                        {
                            try
                            {
                                System.Security.Principal.SecurityIdentifier sid = new System.Security.Principal.SecurityIdentifier(bytes, 0);
                                p.Inlines.Add(sid.ToString());
                            }
                            catch
                            {
                                p.Inlines.Add(valueString);
                            }
                        }
                        else
                        {
                            System.IO.MemoryStream ms = new System.IO.MemoryStream(bytes);
                            try
                            {
                                BitmapImage photoBitmap = new BitmapImage();
                                photoBitmap.BeginInit();
                                photoBitmap.StreamSource = ms;
                                photoBitmap.EndInit();
                                Image img = new Image();
                                img.Source = photoBitmap;
                                img.Stretch = Stretch.None;
                                p.Inlines.Add(img);
                            }
                            catch
                            {
                                p.Inlines.Add(valueString);
                            }
                        }
                    }
                    else
                    {
                        if (valueString.StartsWith("CN=") || valueString.ToLower().StartsWith("http://"))
                        {
                            //string display = Regex.Match(valueString + ",", "CN=(.*?),").Groups[1].Captures[0].Value;
                            //Hyperlink h = new Hyperlink(new Run(display));
                            Hyperlink h = new Hyperlink(new Run(valueString));
                            h.NavigateUri = new Uri(valueString, valueString.ToLower().StartsWith("http://") ? UriKind.Absolute : UriKind.Relative);
                            h.Click += new RoutedEventHandler(HyperlinkClicked);
                            //p.TextIndent = -20;
                            //p.Margin = new Thickness(20, p.Margin.Top, p.Margin.Right, p.Margin.Bottom);
                            p.Inlines.Add(h);
                        }
                        else
                            p.Inlines.Add(new Run(valueString));
                    }
                }
            }

            return p;
        }

        void HyperlinkClicked(object sender, RoutedEventArgs e)
        {
            Hyperlink h = sender as Hyperlink;
            if (h != null)
            {
                if (h.NavigateUri.IsAbsoluteUri)
                    Process.Start(h.NavigateUri.AbsoluteUri);
                else
                    DisplayEntry("LDAP://" + h.NavigateUri.OriginalString, true);
            }
        }

        private void BrowseBackCommand_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = backStack.Count > 0;
        }

        private void BrowseBackCommand_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            string name = backStack.Pop();
            DisplayEntry(name, false);
            if (backStack.Count == 0)
                CommandManager.InvalidateRequerySuggested();
        }

        // The following shouldn't be necessary as BrowseBack is a standard command
        //private void Window_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        //{
        //    if (e.ChangedButton == MouseButton.XButton1)
        //    {
        //       NavigationCommands.BrowseBack.Execute(null, backButton);
        //    }
        //}


    }
}
