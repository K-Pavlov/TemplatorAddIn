namespace AddionalPluginUtilities
{
    using System;
    using System.Data;
    using System.Data.SqlClient;
    using System.Text.RegularExpressions;
    using System.Windows.Forms;
    using Extensibility;
    using EnvDTE;
    using EnvDTE80;
    using Microsoft.VisualStudio.CommandBars;
    using System.Runtime.InteropServices;

    public partial class TemplateForm : Form
    {
        [DllImportAttribute("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);

        [DllImportAttribute("user32.dll")]
        public static extern bool ReleaseCapture();

        private const int WM_NCLBUTTONDOWN = 0xA1;
        private const int HT_CAPTION = 0x2;
        private const int WM_NCLBUTTONDBLCLK = 0x00A3;

        private SqlConnection connection;
        private DTE2 applicationObject;

        public DTE2 ApplicationObject
        {
            get
            {
                return this.applicationObject;
            }

            private set
            {
                this.applicationObject = value;
            }
        }


        public TemplateForm(DTE2 applicationObject)
        {
            InitializeComponent();
            this.connection = new SqlConnection("Server=localhost;" +
            "Integrated security=SSPI;" +
            "database=master;");

            this.ApplicationObject = applicationObject;
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_NCLBUTTONDBLCLK)
            {
                m.Result = IntPtr.Zero;
                return;
            }

            base.WndProc(ref m);
        }

        private void writeButton_Click(object sender, EventArgs e)
        {
            string template = this.textBoxTemplate.Text,
                module = this.textBoxModule.Text,
                method = this.textBoxMethod.Text,
                language = this.textBoxLanguage.Text,
                text = this.textBoxText.Text,
                id = this.textBoxId.Text,
                templatedText;

            var document = this.applicationObject.ActiveDocument;

            if (document == null)
            {
                MessageBox.Show("No document opened");
                return;
            }

            if (template == "")
            {
                templatedText = "\"" + module + " " + method + " "
                    + language + " " + text + " " + id + "\"";
            }
            else
            {
                templatedText = "\"";
                templatedText += Regex.Replace(template, "{{module}}", module, RegexOptions.IgnoreCase);
                templatedText = Regex.Replace(templatedText, "{{method}}", method, RegexOptions.IgnoreCase);
                templatedText = Regex.Replace(templatedText, "{{language}}", language, RegexOptions.IgnoreCase);
                templatedText = Regex.Replace(templatedText, "{{text}}", text, RegexOptions.IgnoreCase);
                templatedText = Regex.Replace(templatedText, "{{id}}", id, RegexOptions.IgnoreCase);
                templatedText += "\"";
            }

            this.InsertTemplate(document, templatedText);

            //
            // 

            try
            {
                this.connection.Open();
                SqlCommand writeCommand = PrepareWriteCommand(new Values(module, method, language, text, id));
                writeCommand.ExecuteNonQuery();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Addin", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            finally
            {
                if (this.connection.State == ConnectionState.Open)
                {
                    this.connection.Close();
                }
            }
        }

        private void InsertTemplate(Document document, string toInsert)
        {
            var textDocument = document.Object() as TextDocument;

            textDocument.StartPoint.CreateEditPoint();
            textDocument.Selection.Insert(toInsert);
        }

        private void closeButton_Click(object sender, EventArgs e)
        {
            this.Hide();
        }

        private SqlCommand PrepareWriteCommand(Values values)
        {
            SqlCommand command = new SqlCommand(null, this.connection);
            SqlParameter module = new SqlParameter("@module", SqlDbType.VarChar, 50),
                method = new SqlParameter("@method", SqlDbType.VarChar, 50),
                language = new SqlParameter("@language", SqlDbType.VarChar, 50),
                text = new SqlParameter("@text", SqlDbType.Text),
                id = new SqlParameter("@id", SqlDbType.Int);

            command.CommandText =
                "INSERT INTO AddinInfo.dbo.addinContent (module, method, language, text, id) "
                + "VALUES (@module, @method, @language, @text, @id)";

            module.Value = values.Module;
            method.Value = values.Method;
            language.Value = values.Language;
            text.Value = values.Text;
            id.Value = values.Id;

            command.Parameters.Add(module);
            command.Parameters.Add(method);
            command.Parameters.Add(language);
            command.Parameters.Add(text);
            command.Parameters.Add(id);

            command.Prepare();

            return command;
        }

        private void TemplateForm_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }

        private struct Values
        {
            private string module;
            private string method;
            private string language;
            private string text;
            private int id;

            public string Text
            {
                get
                {
                    return this.text;
                }

                private set
                {
                    this.text = value;
                }
            }


            public int Id
            {
                get
                {
                    return this.id;
                }

                private set
                {
                    this.id = value;
                }
            }


            public string Language
            {
                get
                {
                    return this.language;
                }

                private set
                {
                    this.language = value;
                }
            }

            public string Method
            {
                get
                {
                    return this.method;
                }

                private set
                {
                    this.method = value;
                }
            }


            public string Module
            {
                get
                {
                    return this.module;
                }

                private set
                {
                    this.module = value;
                }
            }

            public Values(string module, string method, string language, string text, string id)
            {
                this.module = "";
                this.method = "";
                this.language = "";
                this.text = "";
                this.id = 0;

                this.Module = module;
                this.Method = method;
                this.Language = language;
                this.text = text;
                this.Id = validateId(id);
            }

            private int validateId(string id)
            {
                int number;
                bool result;

                result = int.TryParse(id, out number);

                if (result)
                {
                    return number;
                }
                else
                {
                    throw new ArgumentException("Id must be a number");
                }
            }
        }
    }
}