using System.Windows.Forms;
namespace DDM.ExcelAddIn
{
    partial class DDMWindow
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(DDMWindow));
            this.elementHost1 = new System.Windows.Forms.Integration.ElementHost();
            this.ddmControl1 = new DDMControl();
            this.SuspendLayout();
            // 
            // elementHost1
            // 
            resources.ApplyResources(this.elementHost1, "elementHost1");
            this.elementHost1.Name = "elementHost1";
            this.elementHost1.Child = this.ddmControl1;
            // 
            // DDMWindow
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.elementHost1);
            this.MaximizeBox = false;
            this.Name = "DDMWindow";
            this.Load += new System.EventHandler(this.DDMWindow_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Integration.ElementHost elementHost1;
        private DDMControl ddmControl1;
    }
}