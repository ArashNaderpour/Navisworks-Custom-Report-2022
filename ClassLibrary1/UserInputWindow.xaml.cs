using System.Windows;
using System.IO;
using System.Windows.Forms;
using System.Reflection;

using Autodesk.Navisworks.Api;
using ComBridge = Autodesk.Navisworks.Api.ComApi.ComApiBridge;
using ComApi = Autodesk.Navisworks.Api.Interop.ComApi;
using System.Linq;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.Text.RegularExpressions;
using Microsoft.VisualBasic;
using System.Web.Script.Serialization;
using System.Drawing;

namespace CustomReport2020
{
    /// <summary>
    /// Interaction logic for UserInputWindow.xaml
    /// </summary>
    public partial class UserInputWindow : Window
    {
        public int widthResolution;
        public int heightResolution;

        string svpName;

        public float headerSizeInput;
        public float commentsSizeInput;
        public float reviewTextSizeInput;

        int numOfColumns;
        int rowHeight;

        int counter = 0;

        private string saveFolderPath;

        private VPGroup savedViewPoints = new VPGroup("Root");

        private Dictionary<string, string> viewpointReviewTexts = new Dictionary<string, string>();

        private Dictionary<string, string> viewpointComments = new Dictionary<string, string>();

        Document oDoc = Autodesk.Navisworks.Api.Application.ActiveDocument;

        private List<string> userTextInputs = new List<string>();

        private List<System.Windows.Media.Color> userColorInputs = new List<System.Windows.Media.Color>();

        //private List<string> userColorInputs = new List<string>();

        private Dictionary<string, string> backgroundColorImage = new Dictionary<string, string>();

        private string nameTextColor = "rgb(30, 30, 30)";

        public UserInputWindow()
        {
            InitializeComponent();
        }

        public void Add_Clicked(object sender, RoutedEventArgs e)
        {
            UIMethods.AddGroup(this.GroupSettingGrid);
        }

        public void Remove_Clicked(object sender, RoutedEventArgs e)
        {
            UIMethods.DeleteGroup(this.GroupSettingGrid);
        }

        public void Export_Clicked(object sender, RoutedEventArgs e)
        {
            try
            {
                widthResolution = int.Parse(WidthInput.Text);

                if (widthResolution <= 0)
                {
                    throw new System.Exception();
                }
            }
            catch
            {
                System.Windows.Forms.MessageBox.Show("Please enter a number larger than 0 as the width.");
                return;
            }

            try
            {
                heightResolution = int.Parse(HeightInput.Text);

                if (heightResolution <= 0)
                {
                    throw new System.Exception();
                }
            }
            catch
            {
                System.Windows.Forms.MessageBox.Show("Please enter a number larger than 0 as the height.");
                return;
            }

            try
            {
                this.numOfColumns = int.Parse(this.ColumnCountInput.Text);

                if (this.numOfColumns <= 0)
                {
                    throw new System.Exception();
                }
            }
            catch
            {
                System.Windows.Forms.MessageBox.Show("Please enter a number larger than 0 as the number of columns.");
                return;
            }

            try
            {
                this.rowHeight = int.Parse(this.RowHeightInput.Text);

                if (this.rowHeight <= 0)
                {
                    throw new System.Exception();
                }
            }
            catch
            {
                System.Windows.Forms.MessageBox.Show("Please enter a number larger than 0 as the row height.");
                return;
            }

            try
            {
                this.headerSizeInput = float.Parse(this.HeaderSizeInput.Text);

                if (this.headerSizeInput <= 0)
                {
                    throw new System.Exception();
                }
            }
            catch
            {
                System.Windows.Forms.MessageBox.Show("Please enter a number larger than 0 as the header font size.");
                return;
            }

            try
            {
                this.commentsSizeInput = float.Parse(this.CommentsSizeInput.Text);

                if (this.commentsSizeInput <= 0)
                {
                    throw new System.Exception();
                }
            }
            catch
            {
                System.Windows.Forms.MessageBox.Show("Please enter a number larger than 0 as the comments font size.");
                return;
            }

            try
            {
                this.reviewTextSizeInput = float.Parse(this.ReviewTextSizeInput.Text);

                if (this.reviewTextSizeInput <= 0)
                {
                    throw new System.Exception();
                }
            }
            catch
            {
                System.Windows.Forms.MessageBox.Show("Please enter a number larger than 0 as the reviewText font size.");
                return;
            }

            foreach (System.Windows.Controls.Grid group in this.GroupSettingGrid.Children)
            {
                foreach (UIElement element in group.Children)
                {
                    if (element.GetType().ToString() == "System.Windows.Controls.TextBox")
                    {
                        try
                        {
                            string userTextInput = ((System.Windows.Controls.TextBox)element).Text.ToLower();

                            if (userTextInput.Replace(" ", "") == string.Empty)
                            {
                                System.Windows.Forms.MessageBox.Show("It is not possible to leave any of the text inputs empty! Please fill the empty ones or remove them.");

                                ((System.Windows.Controls.TextBox)element).Text = "Empty";

                                return;
                            }
                            else if (userTextInput[0] == ' ')
                            {
                                userTextInput = userTextInput.Substring(1, userTextInput.Length - 1);

                                if (this.userTextInputs.Contains(userTextInput))
                                {
                                    System.Windows.Forms.MessageBox.Show("Two groups with the name \"" + userTextInput + "\" exist. Please remove one of them.");

                                    this.userTextInputs.Clear();
                                    this.userColorInputs.Clear();

                                    return;
                                }
                                else
                                {
                                    this.userTextInputs.Add(userTextInput);

                                    ((System.Windows.Controls.TextBox)element).Text = userTextInput;
                                }
                            }
                            else if (userTextInput.Last() == ' ')
                            {
                                userTextInput = userTextInput.Substring(0, userTextInput.Length - 2);

                                if (this.userTextInputs.Contains(userTextInput))
                                {
                                    System.Windows.Forms.MessageBox.Show("Two groups with the name \"" + userTextInput + "\" exist. Please remove one of them.");

                                    this.userTextInputs.Clear();
                                    this.userColorInputs.Clear();

                                    return;
                                }
                                else
                                {
                                    this.userTextInputs.Add(userTextInput);

                                    ((System.Windows.Controls.TextBox)element).Text = userTextInput;
                                }
                            }
                            else
                            {
                                if (this.userTextInputs.Contains(userTextInput))
                                {
                                    System.Windows.Forms.MessageBox.Show("Two groups with the name \"" + userTextInput + "\" exist. Please remove one of them.");

                                    this.userTextInputs.Clear();
                                    this.userColorInputs.Clear();

                                    return;
                                }
                                else
                                {
                                    this.userTextInputs.Add(userTextInput);
                                }
                            }
                        }
                        catch (System.Exception ex)
                        {
                            System.Windows.Forms.MessageBox.Show(ex.ToString());
                            return;
                        }

                    }
                    else
                    {
                        Xceed.Wpf.Toolkit.ColorPicker cp = (Xceed.Wpf.Toolkit.ColorPicker)element;

                        string userColorInput = "rgb(" + cp.SelectedColor.Value.R + "," + cp.SelectedColor.Value.G + "," + cp.SelectedColor.Value.B + ")";

                        //userColorInputs.Add(userColorInput);

                        userColorInputs.Add(cp.SelectedColor.Value);
                    }
                }
            }

            if (this.rowHeight <= this.heightResolution)
            {
                System.Windows.Forms.MessageBox.Show("It is recommended for \"Row Height\" to be bigger than \"Image Height\".");
            }

            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == System.Windows.Forms.DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    string[] files = Directory.GetFiles(fbd.SelectedPath);

                    saveFolderPath = fbd.SelectedPath;
                }
                else
                {
                    this.userTextInputs.Clear();
                    this.userColorInputs.Clear();

                    return;
                }
            }

            try
            {
                dumpSavedVP(widthResolution, heightResolution);
                SaveLogo();
                GenerateHTML();

                System.Windows.Forms.MessageBox.Show("Successful.");
            }

            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.ToString());
            }
        }

        private void dumpSavedVP(int width, int height)
        {
            // get the state of COM
            ComApi.InwOpState10 oState = ComBridge.State;

            // get the IO plugin for image
            ComApi.InwOaPropertyVec options = oState.GetIOPluginOptions("lcodpimage");

            // configure the option "export.image.format" to export png
            foreach (ComApi.InwOaProperty opt in options.Properties())
            {
                if (opt.name == "export.image.format")
                    opt.value = "lcodpexpng";
                if (opt.name == "export.image.width")
                    opt.value = width;
                if (opt.name == "export.image.height")
                    opt.value = height;
            }

            this.GenerateViewPointImage(oDoc.SavedViewpoints.Value, oState, options, this.savedViewPoints);
        }

        private void GenerateHTML()
        {
            string htmlTemplate = Properties.Resources.HTMLTemplate;

            string tempText = string.Empty;

            List<string> lines = htmlTemplate.Split('\n').ToList<string>();
            string reportName = oDoc.FileName.Split('\\')[oDoc.FileName.Split('\\').Length - 1].Split('.')[0];

            string reportHeader = string.Format("<h1>{0}</h1>", "Viewpoints__" + oDoc.FileName.Split('\\')[oDoc.FileName.Split('\\').Length - 1]);

            string cellSize = ((100 / this.numOfColumns) - 2).ToString();

            int insertLine = 9;

            tempText = "h2, h3, h4, h5 { font-family: arial, sans-serif; text-transform: capitalize; margin: 1em; font-size: " + this.headerSizeInput.ToString() + "pt; }";
            insertLine++;
            lines.Insert(insertLine, tempText);

            tempText = ".viewpointImage { max-width: " + this.ImageScale.Value.ToString() + "%; max-height: " + this.ImageScale.Value.ToString() + "%; margin: 1em; float: left; }";
            insertLine++;
            lines.Insert(insertLine, tempText);

            tempText = "span.name { color: black; display: inline-block; font-size: " + this.commentsSizeInput.ToString() + "pt; margin-top: 10pt; }";
            insertLine++;
            lines.Insert(insertLine, tempText);

            tempText = "p.value { color: " + nameTextColor + "; display: inline; margin-left: 1em; font-size: " + this.commentsSizeInput.ToString() + "pt; }";
            insertLine++;
            lines.Insert(insertLine, tempText);

            tempText = "div.otherViewPoints { background-color: white; border: 1px solid black; width:" +
                    cellSize + "%; float: left; margin-bottom: 10px; height:" + this.rowHeight.ToString() + "px; page-break-inside: avoid; break-inside: avoid; overflow:hidden; position:relative; }";
            insertLine++;
            lines.Insert(insertLine, tempText);

            tempText = ".bg { position: absolute; height: " + this.rowHeight.ToString() + "px; width: 100%; z-index: -1; }";
            insertLine++;
            lines.Insert(insertLine, tempText);

            for (int i = 0; i < this.userTextInputs.Count; i++)
            {
                var bitmap = new Bitmap(1, 1);

                if (this.userColorInputs[i].ToString() != "#00FFFFFF")
                {
                    using (var g = System.Drawing.Graphics.FromImage(bitmap))
                    {
                        using (var brush = new SolidBrush(System.Drawing.Color.FromArgb(this.userColorInputs[i].R,
                            this.userColorInputs[i].G, this.userColorInputs[i].B)))
                        {
                            g.FillRectangle(brush, new Rectangle(0, 0, bitmap.Width, bitmap.Height));
                        }
                        using (var ms = new MemoryStream())
                        {
                            bitmap.Save(ms, ImageFormat.Png);
                            string SigBase64 = System.Convert.ToBase64String(ms.GetBuffer()); //Get Base64

                            this.backgroundColorImage.Add(this.userTextInputs[i], SigBase64);
                        }
                    }

                    tempText = "div." + this.userTextInputs[i].Replace(" ", "_") + "{ border: 1px solid black; width:" +
                        cellSize + "%; float: left; margin-bottom: 10px; height:" + this.rowHeight.ToString() + "px; page-break-inside: avoid; break-inside: avoid; overflow:hidden; position:relative; }";
                    insertLine++;
                    lines.Insert(insertLine, tempText);
                }
                else
                {
                    this.backgroundColorImage.Add(this.userTextInputs[i], "");

                    tempText = "div." + this.userTextInputs[i].Replace(" ", "_") + "{ border: 1px solid black; display: none; width:" +
                        cellSize + "%; float: left; margin-bottom: 10px; height:" + this.rowHeight.ToString() + "px; page-break-inside: avoid; break-inside: avoid; overflow:hidden; position:relative; }";
                    insertLine++;
                    lines.Insert(insertLine, tempText);
                }
            }

            tempText = "span.reviewText { color: " + nameTextColor + "; display: inline; margin-left: 1em; margin-right: 5em; font-size: " + this.reviewTextSizeInput.ToString() + "pt; }";
            insertLine++;
            lines.Insert(insertLine, tempText);

            lines.Add(reportHeader);

            GenerateViewpointHTML(savedViewPoints, string.Empty, lines, viewpointComments, viewpointReviewTexts, this.counter);

            lines.Add("</div>");
            lines.Add("</body>");
            lines.Add("</html>");

            File.WriteAllLines(saveFolderPath + "\\" + reportName + ".html", lines);
        }

        private void GenerateViewPointImage(SavedItemCollection savedItem, ComApi.InwOpState10 state,
            ComApi.InwOaPropertyVec options, VPGroup savedViewPoints)
        {
            foreach (SavedItem oSVI in savedItem)
            {
                if (oSVI.GetType().ToString() == "Autodesk.Navisworks.Api.SavedViewpoint")
                {
                    string removableChars = Regex.Escape(@"|<>*/\?:");
                    string pattern = "[" + removableChars + "]";

                    svpName = Regex.Replace(oSVI.DisplayName, pattern, "").Replace("\"", string.Empty);

                    // set the current viewpoint  
                    oDoc.SavedViewpoints.CurrentSavedViewpoint = oSVI;

                    Viewpoint oCurVP = oDoc.CurrentViewpoint;

                    // Extract the ReviewText Json string from active viewpoint and deserialize it to the ReviewText custom object
                    string reviewTextJson = oDoc.ActiveView.GetRedlines();

                    JavaScriptSerializer javaScriptSerializer = new JavaScriptSerializer();

                    ReviewTextRootObject reviewTextRootObject = javaScriptSerializer.Deserialize<ReviewTextRootObject>(reviewTextJson);

                    // A List to store ReviewText texts of each saved viewpoint
                    string reviewTexts = "";

                    if (this.ReviewTextInclude.IsChecked == true)
                    {
                        for (int i = 0; i < reviewTextRootObject.Values.Count; i++)
                        {
                            if (i > 0 && string.IsNullOrWhiteSpace(reviewTextRootObject.Values[i].Text) == false)
                            {
                                reviewTexts += "<br>";
                            }
                            reviewTexts += reviewTextRootObject.Values[i].Text;
                        }
                    }

                    // A List to store comments of each saved viewpoint
                    string comments = "";

                    if (this.CommentsInclude.IsChecked == true)
                    {
                        for (int i = 0; i < oSVI.Comments.Count; i++)
                        {
                            if (i > 0)
                            {
                                comments += "<br>";
                            }
                            comments += oSVI.Comments[i].CreationDate;
                            comments += " /--/ ";
                            comments += oSVI.Comments[i].Author;
                            comments += " /--/ ";
                            comments += oSVI.Comments[i].Body;
                        }
                    }

                    this.SaveViewpointData(comments, reviewTexts);

                    if (string.IsNullOrWhiteSpace(svpName))
                    {
                        System.Windows.Forms.MessageBox.Show("The process was canceled.");
                        return;
                    }

                    else
                    {
                        savedViewPoints.children.Add(svpName);
                    }

                    // the image file name
                    string imageFileName = saveFolderPath + "\\" + svpName + ".png";

                    // delete the existing image if there is.
                    if (System.IO.File.Exists(imageFileName))
                        System.IO.File.Delete(imageFileName);
                    try
                    {
                        //export the viewpoint to the image
                        state.DriveIOPlugin("lcodpimage", imageFileName, options);
                    }

                    catch (System.Exception ex)
                    {
                        System.Windows.Forms.MessageBox.Show(ex.ToString());
                    }
                }

                if (oSVI.GetType().ToString() == "Autodesk.Navisworks.Api.FolderItem")
                {
                    VPGroup vpGroup = new VPGroup(oSVI.DisplayName);

                    savedViewPoints.subGroups.Add(vpGroup);

                    GenerateViewPointImage(((GroupItem)oSVI).Children, state, options, vpGroup);
                }
            }
        }

        private void SaveLogo()
        {
            try
            {
                Properties.Resources.logo.Save(saveFolderPath + "\\" + "logo.png", ImageFormat.Png);
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.ToString());
            }
        }

        private string ViewpointHTML(string vPName, string vPTitle, string className, string comments, string reviewText)
        {
            string backgroundColor = string.Empty;

            if (this.backgroundColorImage.Keys.Contains(className) && this.backgroundColorImage[className] != string.Empty)
            {
                backgroundColor = string.Format("<img class=\"bg\" src='data:image/png;base64,{0}'/>", this.backgroundColorImage[className]);
            }

            string result = string.Format("<div class=" + className.Replace(" ", "_") + ">\n" + backgroundColor + "<h2>{0}</h2><img class=\"viewpointImage\" src = \"{1}\"><span class=\"name\"><strong>Review Text: </strong></span><span class=\"reviewText\">{3}</span><br><span class=\"name\"><strong>Comments: </strong></span><p class=\"value\">{2}</p><div class=\"spacer\"></div>\n</div>",
                vPTitle, vPName + ".png", comments, reviewText);


            return result;
        }

        private void GenerateViewpointHTML(VPGroup vPGroup, string title, List<string> lines, Dictionary<string, string> comments, Dictionary<string, string> reviewText, int counter)
        {

            string baseTitle = title + vPGroup.groupName + " => ";

            for (int i = 0; i < vPGroup.children.Count; i++)
            {
                string vPName = vPGroup.children[i];
                string vPTitle = baseTitle + vPName;

                if (counter != 0 && counter % this.numOfColumns == 0)
                {
                    lines.Add("</div>");
                    counter = 0;
                }

                if (counter == 0)
                {
                    lines.Add("<div class=\"row\">");
                }

                string className = string.Empty;

                foreach (string folderName in baseTitle.ToLower().Split(new string[] { " => " }, System.StringSplitOptions.None))
                {
                    if (folderName != "root")
                    {
                        if (this.userTextInputs.Contains(folderName))
                        {
                            className = folderName;
                        }
                    }
                }

                if (className != string.Empty)
                {
                    lines.Add(this.ViewpointHTML(vPName, vPTitle, className, comments[vPName], reviewText[vPName]));
                }
                else
                {
                    lines.Add(this.ViewpointHTML(vPName, vPTitle, "otherViewPoints", comments[vPName], reviewText[vPName]));
                }

                counter++;
                this.counter = counter;
            }

            foreach (VPGroup subGroup in vPGroup.subGroups)
            {
                GenerateViewpointHTML(subGroup, baseTitle, lines, comments, reviewText, this.counter);
            }
        }

        private void SaveViewpointData(string comments, string redText)
        {
            try
            {
                this.viewpointComments.Add(svpName, comments);

                this.viewpointReviewTexts.Add(svpName, redText);
            }

            catch
            {
                string message = string.Format("All viewpoints are required to have unique names for export process to continue." +
                    " Please enter a new name for the second \"{0}\" viewpoint in the box below and press OK. Press Cancel to abandon the export process", svpName);

                svpName = Interaction.InputBox(message, "Title", "", -1, -1);

                if (svpName == "")
                {
                    return;
                }

                else
                {
                    SaveViewpointData(comments, redText);
                }
            }
        }
    }
}



