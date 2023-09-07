using Microsoft.Office.Interop.Word;
using System.Reflection;
using System.Reflection.Metadata;
using System.Windows.Forms.Design;
using Word = Microsoft.Office.Interop.Word;


namespace WordManager
{
    public partial class Form1 : Form
    {
        public object oMissing = System.Reflection.Missing.Value;
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */

            //Start Word and create a new document.
            Word._Application oWord;
            Word._Document oDoc;
            oWord = new Word.Application();
            oWord.Visible = true;
            object oTemplate = "C:\\Users\\students\\Downloads\\tf16392716_win32.dotx";
            oDoc = oWord.Documents.Add(oTemplate, oMissing, oMissing, oMissing);


            Replace("ВИ", textBox1.Text.Substring(0, 1) + textBox2.Text.Substring(0, 1), oDoc);
            Replace("Ваше имя", textBox1.Text + " " + textBox2.Text, oDoc);

            Word.InlineShape oShape;
            object oClassType = "Button";
            var wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oShape = wrdRng.InlineShapes.AddOLEObject(ref oClassType, ref oMissing,
            ref oMissing, ref oMissing, ref oMissing,
            ref oMissing, ref oMissing, ref oMissing);
        }

        private void Replace(string find, string replace, Word._Document oDoc)
        {
            Word.Find findObject = oDoc.Application.Selection.Find;
            findObject.ClearFormatting();
            findObject.Text = find;
            findObject.Replacement.ClearFormatting();
            findObject.Replacement.Text = replace;

            object replaceAll = Word.WdReplace.wdReplaceAll;
            findObject.Execute(ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
        ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
        ref replaceAll, ref oMissing, ref oMissing, ref oMissing, ref oMissing);
        }
    }
}