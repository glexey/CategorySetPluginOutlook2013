using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Ribbon;
using Office = Microsoft.Office.Core;

namespace CategorySetPluginOutlook2013 {
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility {
        private Office.IRibbonUI ribbon;

        public Ribbon() {
        }

        #region Working with categories

        static String[] separators = new String[] { ", ", "," };

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
            else if (item is NoteItem) ret = (item as NoteItem).Categories;
            else if (item is JournalItem) ret = (item as JournalItem).Categories;
            else if (item is SharingItem) ret = (item as SharingItem).Categories;
            else if (item is ContactItem) ret = (item as ContactItem).Categories;
            else if (item is DistListItem) ret = (item as DistListItem).Categories;
            else if (item is TaskRequestAcceptItem) ret = (item as TaskRequestAcceptItem).Categories;
            else if (item is TaskRequestDeclineItem) ret = (item as TaskRequestDeclineItem).Categories;
            else if (item is TaskRequestUpdateItem) ret = (item as TaskRequestUpdateItem).Categories;
            else if (item is ReportItem) ret = (item as ReportItem).Categories;
            // ^^^ the list comes from http://msdn.microsoft.com/en-us/library/office/ff861539.aspx
            else if (item is ConversationHeader) {
                Conversation conv = (item as ConversationHeader).GetConversation();
                ret = conv.GetAlwaysAssignCategories(ex.CurrentFolder.Store);
            }
            else { Debug.WriteLine("unsupported item type " + item); return null; } // Unsupported object type
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
            else if (item is NoteItem) { (item as NoteItem).Categories = value; (item as NoteItem).Save(); }
            else if (item is JournalItem) { (item as JournalItem).Categories = value; (item as JournalItem).Save(); }
            else if (item is SharingItem) { (item as SharingItem).Categories = value; (item as SharingItem).Save(); }
            else if (item is ContactItem) { (item as ContactItem).Categories = value; (item as ContactItem).Save(); }
            else if (item is DistListItem) { (item as DistListItem).Categories = value; (item as DistListItem).Save(); }
            else if (item is TaskRequestAcceptItem) { (item as TaskRequestAcceptItem).Categories = value; (item as TaskRequestAcceptItem).Save(); }
            else if (item is TaskRequestDeclineItem) { (item as TaskRequestDeclineItem).Categories = value; (item as TaskRequestDeclineItem).Save(); }
            else if (item is TaskRequestUpdateItem) { (item as TaskRequestUpdateItem).Categories = value; (item as TaskRequestUpdateItem).Save(); }
            else if (item is ReportItem) { (item as ReportItem).Categories = value; (item as ReportItem).Save(); }
            else if (item is ConversationHeader) {
                Conversation conv = (item as ConversationHeader).GetConversation();
                if (value != "")
                    conv.SetAlwaysAssignCategories(value, ex.CurrentFolder.Store);
            }
        }

        private bool hasCat(Object item, String targetCat) {
            String str = getCat(item);
            if (str == null) return true; // Unsupported outlook item type
            return cSplit(str).Contains(targetCat);
        }

        private void removeCat(Object item, String targetCat) {
            String str = getCat(item);
            if (str == null) return; // Unsupported outlook item type
            if (item is ConversationHeader)
                (item as ConversationHeader).GetConversation().ClearAlwaysAssignCategories(ex.CurrentFolder.Store);
            setCat(item, String.Join(",", cSplit(str).Where(s => s != targetCat).ToList()));
        }

        private void addCat(Object item, String targetCat) {
            if (hasCat(item, targetCat)) return; // Already has this category, or unsupported outlook item type
            String str = getCat(item);
            setCat(item, str + "," + targetCat);
        }

        #endregion

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID) {
            return GetResourceText("CategorySetPluginOutlook2013.Ribbon.xml");
        }

        #endregion

        #region Ribbon Callbacks

        public void Ribbon_Load(Office.IRibbonUI ribbonUI) {
            this.ribbon = ribbonUI;
        }

        HashSet<String> catsAllHave;  // Categories that all messages have
        HashSet<String> catsSomeHave; // Categories that some messages have
        IEnumerable<Category> sorted_cats; // Categories sorted [alphabetically]
        Explorer ex;

        private NameSpace objNameSpace = null;

        // [Hack] Assumption :: getItemCount() always gets called before other get* callbacks
        public int catGallery_getItemCount(Office.IRibbonControl control) {

            //Debug.WriteLine("catGallery_getItemCount(Office.IRibbonControl control)");
            if (objNameSpace == null)
                objNameSpace = Globals.ThisAddIn.Application.GetNamespace("MAPI");

            // Check the categories which are already set on the selected items
            ex = Globals.ThisAddIn.Application.ActiveExplorer();
            catsAllHave = new HashSet<String>();
            catsSomeHave = new HashSet<String>();
            if (ex != null) {
                int itemnum = 0;
                Selection convHeaders = ex.Selection.GetSelection(OlSelectionContents.olConversationHeaders) as Selection;
                if (convHeaders.Count > 0) {
                    Debug.WriteLine("getcount: convHeaders.Count > 0");
                    // If a conversation header is selected, we'll examine/set "always set" categories on a conversation
                    foreach (ConversationHeader item in convHeaders) {
                        String[] cats = cSplit(getCat(item));
                        foreach (String cat in cats) catsSomeHave.Add(cat);
                        if (itemnum++ == 0)
                            foreach (String cat in cats) catsAllHave.Add(cat);
                        catsAllHave.RemoveWhere(cat => !cats.Contains<String>(cat));
                    }
                }
                else {
                    Debug.WriteLine("getcount: convHeaders.Count > 0");
                    // If conversation header is not selected, we'll examine/set categories on individual items
                    foreach (Object item in ex.Selection) {
                        String[] cats = cSplit(getCat(item));
                        foreach (String cat in cats) catsSomeHave.Add(cat);
                        if (itemnum++ == 0)
                            foreach (String cat in cats) catsAllHave.Add(cat);
                        catsAllHave.RemoveWhere(cat => !cats.Contains<String>(cat));
                    }
                }
            }

            // Sort the categories and save for later access by get* methods
            sorted_cats = from Category cat in objNameSpace.Categories orderby cat.Name select cat;

            return objNameSpace.Categories.Count;
        }

        public string catGallery_getItemLabel(Office.IRibbonControl control, int index) {
            return sorted_cats.ElementAt(index).Name;
        }

        public stdole.IPictureDisp catGallery_getItemImage(Office.IRibbonControl control, int index) {
            // On our last get* call I invalidate the control, so that next time it's invoked
            // our get* methods would be called all over again. We need this since any selection change 
            // invalidates the category images
            if (index == sorted_cats.Count() - 1 )
                this.ribbon.InvalidateControl("SetCategory");

            Category cat = sorted_cats.ElementAt(index);
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

            return AxHostConverter.ImageToPictureDisp(img);
        }
        
        public void catGallery_clicked(Office.IRibbonControl control, string selectedId, int selectedIndex) {
            String targetCategory = sorted_cats.ElementAt(selectedIndex).Name;
            ex = Globals.ThisAddIn.Application.ActiveExplorer();
            if (ex != null) {
                Selection convHeaders = ex.Selection.GetSelection(OlSelectionContents.olConversationHeaders) as Selection;
                IEnumerable<Object> items = (convHeaders.Count > 0) ?
                    convHeaders.Cast<ConversationHeader>() : ex.Selection.Cast<Object>();

                if (items.All(item => hasCat(item, targetCategory)))
                    // If all items already contain the category, then remove it
                    items.AsParallel().ForAll(item => removeCat(item, targetCategory));
                else
                    // Otherwise, add the category
                    items.AsParallel().ForAll(item => addCat(item, targetCategory));
            }
        }

        #endregion

        #region GetResourceText Helper

        private static string GetResourceText(string resourceName) {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i) {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0) {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i]))) {
                        if (resourceReader != null) {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
