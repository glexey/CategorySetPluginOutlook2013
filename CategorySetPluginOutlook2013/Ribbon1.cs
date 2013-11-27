using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Ribbon;

namespace CategorySetPluginOutlook2013 {
    public partial class Ribbon1 {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e) {

        }

        private void gallery1_ButtonClick(object sender, RibbonControlEventArgs e) {
            Globals.ThisAddIn.Application.ActiveExplorer().CommandBars.ExecuteMso("AllCategories");
        }

        static String[] separators = new String[] {", ", ","};

        private String[] cSplit(String s) {
            // Split category string to list of strings
            if (s == null) return new String[0];
            return s.Split(separators, StringSplitOptions.RemoveEmptyEntries);
        }

        private String getCat(Object item) {
            String ret = null;
            if (item is MailItem) ret = (item as MailItem).Categories;
            else if (item is MeetingItem) ret = (item as MeetingItem).Categories;
            else if (item is AppointmentItem) ret = (item as AppointmentItem).Categories;
            else if (item is TaskItem) ret = (item as TaskItem).Categories;
            else if (item is RemoteItem) ret = (item as RemoteItem).Categories;
            else if (item is DocumentItem) ret = (item as DocumentItem).Categories;
            else if (item is PostItem) ret = (item as PostItem).Categories;
            else if (item is SharingItem) ret = (item as SharingItem).Categories;
            else if (item is ContactItem) ret = (item as ContactItem).Categories;
            else if (item is DistListItem) ret = (item as DistListItem).Categories;
            else if (item is TaskRequestAcceptItem) ret = (item as TaskRequestAcceptItem).Categories;
            else if (item is TaskRequestDeclineItem) ret = (item as TaskRequestDeclineItem).Categories;
            else if (item is TaskRequestUpdateItem) ret = (item as TaskRequestUpdateItem).Categories;
            else if (item is ReportItem) ret = (item as ReportItem).Categories;
            // complete list at http://msdn.microsoft.com/en-us/library/office/ff861539.aspx
            else return null; // Unsupported object type
            if (ret == null) ret = ""; // Supported type, but Categories list is empty
            return ret;
        }

        private void setCat(Object item, String value) {
            if (item is MailItem) { (item as MailItem).Categories = value; (item as MailItem).Save(); }
            else if (item is MeetingItem) { (item as MeetingItem).Categories = value; (item as MeetingItem).Save(); }
            else if (item is AppointmentItem) { (item as AppointmentItem).Categories = value; (item as AppointmentItem).Save(); }
            else if (item is TaskItem) { (item as TaskItem).Categories = value; (item as TaskItem).Save(); }
            else if (item is TaskRequestItem) { (item as TaskRequestItem).Categories = value; (item as TaskRequestItem).Save(); }
            else if (item is RemoteItem) { (item as RemoteItem).Categories = value; (item as RemoteItem).Save(); }
            else if (item is DocumentItem) { (item as DocumentItem).Categories = value; (item as DocumentItem).Save(); }
            else if (item is PostItem) { (item as PostItem).Categories = value; (item as PostItem).Save(); }
            else if (item is SharingItem) { (item as SharingItem).Categories = value; (item as SharingItem).Save(); }
            else if (item is ContactItem) { (item as ContactItem).Categories = value; (item as ContactItem).Save(); }
            else if (item is DistListItem) { (item as DistListItem).Categories = value; (item as DistListItem).Save(); }
            else if (item is TaskRequestAcceptItem) { (item as TaskRequestAcceptItem).Categories = value; (item as TaskRequestAcceptItem).Save(); }
            else if (item is TaskRequestDeclineItem) { (item as TaskRequestDeclineItem).Categories = value; (item as TaskRequestDeclineItem).Save(); }
            else if (item is TaskRequestUpdateItem) { (item as TaskRequestUpdateItem).Categories = value; (item as TaskRequestUpdateItem).Save(); }
            else if (item is ReportItem) { (item as ReportItem).Categories = value; (item as ReportItem).Save(); }
        }

        private bool hasCat(Object item, String targetCat) {
            String str = getCat(item);
            if (str == null) return true; // Unsupported outlook item type
            return cSplit(str).Contains(targetCat);
        }

        private void removeCat(Object item, String targetCat) {
            String str = getCat(item);
            if (str == null) return; // Unsupported outlook item type
            setCat(item, String.Join(",", cSplit(str).Where(s => s != targetCat).ToList()));
        }

        private void addCat(Object item, String targetCat) {
            if (hasCat(item, targetCat)) return; // Already has this category, or unsupported outlook item type
            String str = getCat(item);
            setCat(item, str + "," + targetCat);
        }

        private void gallery1_Click(object sender, RibbonControlEventArgs e) {
            RibbonGallery rg = (RibbonGallery)sender;
            RibbonDropDownItem dd = rg.SelectedItem;
            Explorer ex = Globals.ThisAddIn.Application.ActiveExplorer();
            String targetCategory = dd.Label;
            if (ex != null && dd != null) {
                var items = ex.Selection.Cast<Object>();
                if (items.All(item => hasCat(item, targetCategory)))
                    // If all items already contain the category, then remove it
                    items.AsParallel().ForAll(item => removeCat(item, targetCategory));
                else
                    // Otherwise, add the category
                    items.AsParallel().ForAll(item => addCat(item, targetCategory));
            }

        }

        private void gallery1_ItemLoad(object sender, RibbonControlEventArgs e) {
            RibbonGallery dd = (RibbonGallery)sender;
            dd.Items.Clear();
            NameSpace objNameSpace = Globals.ThisAddIn.Application.GetNamespace("MAPI");
            if (objNameSpace.Categories.Count > 0) {
                // Check the categories which are already set on the selected items
                Explorer ex = Globals.ThisAddIn.Application.ActiveExplorer();
                HashSet<String> catsAllHave = new HashSet<String>();  // Categories that all messages have
                HashSet<String> catsSomeHave = new HashSet<String>(); // Categories that some messages have
                if (ex != null) {
                    int itemnum = 0;
                    foreach (Object item in ex.Selection) {
                        String[] cats = cSplit(getCat(item));
                        foreach (String cat in cats) catsSomeHave.Add(cat);
                        if (itemnum++ == 0)
                            foreach (String cat in cats) catsAllHave.Add(cat);
                        catsAllHave.RemoveWhere(cat => !cats.Contains<String>(cat));
                    }
                }
                // Draw the categories
                var sorted_cats = from Category cat in objNameSpace.Categories orderby cat.Name select cat;
                foreach (Category cat in sorted_cats) {
                    RibbonDropDownItem rdi = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                    rdi.Label = cat.Name;
                    Rectangle rect = new Rectangle(0, 0, 20, 20);
                    Bitmap img = new Bitmap(rect.Width, rect.Height);
                    Graphics gr = Graphics.FromImage(img);

                    // draw the interior with category color
                    rect.Inflate(-4, -4);
                    Brush b1 = new SolidBrush(ColorTranslator.FromOle((int)cat.CategoryGradientBottomColor));
                    gr.FillRectangle(b1, rect);
                    Pen p1 = new Pen(ColorTranslator.FromOle((int)cat.CategoryBorderColor), 1);
                    gr.DrawRectangle(p1, rect);

                    rect.Inflate(3, 3);
                    Pen p2 = new Pen(Color.MediumBlue, 1);
                    if (catsAllHave.Contains(cat.Name)) {
                        gr.DrawRectangle(p2, rect);
                    }
                    else if (catsSomeHave.Contains(cat.Name)) {
                        p2.DashStyle = System.Drawing.Drawing2D.DashStyle.Dot;
                        gr.DrawRectangle(p2, rect);
                    }
                    rdi.Image = img;
                    dd.Items.Add(rdi);
                }
            }
            //MessageBox.Show(strOutput);
        }
    }
}
