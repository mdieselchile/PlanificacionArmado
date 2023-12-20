namespace PlanificacionArmado
{
    partial class FormSeg
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
            this.Comentario = new System.Windows.Forms.TextBox();
            this.BotonOK = new System.Windows.Forms.Button();
            this.BotonCancel = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // Comentario
            // 
            this.Comentario.Location = new System.Drawing.Point(0, 2);
            this.Comentario.Multiline = true;
            this.Comentario.Name = "Comentario";
            this.Comentario.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.Comentario.Size = new System.Drawing.Size(296, 248);
            this.Comentario.TabIndex = 0;
            // 
            // BotonOK
            // 
            this.BotonOK.Location = new System.Drawing.Point(69, 256);
            this.BotonOK.Name = "BotonOK";
            this.BotonOK.Size = new System.Drawing.Size(75, 23);
            this.BotonOK.TabIndex = 1;
            this.BotonOK.Text = "Aceptar";
            this.BotonOK.UseVisualStyleBackColor = true;
            // 
            // BotonCancel
            // 
            this.BotonCancel.Location = new System.Drawing.Point(150, 256);
            this.BotonCancel.Name = "BotonCancel";
            this.BotonCancel.Size = new System.Drawing.Size(75, 23);
            this.BotonCancel.TabIndex = 2;
            this.BotonCancel.Text = "Cancelar";
            this.BotonCancel.UseVisualStyleBackColor = true;
            // 
            // FormSeg
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(295, 288);
            this.Controls.Add(this.BotonCancel);
            this.Controls.Add(this.BotonOK);
            this.Controls.Add(this.Comentario);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.KeyPreview = true;
            this.Name = "FormSeg";
            this.Text = "Comentario seguimiento";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox Comentario;
        private System.Windows.Forms.Button BotonOK;
        private System.Windows.Forms.Button BotonCancel;
    }
}