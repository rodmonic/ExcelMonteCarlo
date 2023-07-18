namespace Campagna
{
    partial class ProgressBar
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
            this.progressBarLevelOne = new System.Windows.Forms.ProgressBar();
            this.progressBarLevelTwo = new System.Windows.Forms.ProgressBar();
            this.levelOneLabel = new System.Windows.Forms.Label();
            this.levelTwoLabel = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // progressBarLevelOne
            // 
            this.progressBarLevelOne.Location = new System.Drawing.Point(12, 52);
            this.progressBarLevelOne.Name = "progressBarLevelOne";
            this.progressBarLevelOne.Size = new System.Drawing.Size(644, 26);
            this.progressBarLevelOne.TabIndex = 0;
            // 
            // progressBarLevelTwo
            // 
            this.progressBarLevelTwo.Location = new System.Drawing.Point(12, 111);
            this.progressBarLevelTwo.Name = "progressBarLevelTwo";
            this.progressBarLevelTwo.Size = new System.Drawing.Size(644, 26);
            this.progressBarLevelTwo.TabIndex = 1;
            // 
            // levelOneLabel
            // 
            this.levelOneLabel.AutoSize = true;
            this.levelOneLabel.Location = new System.Drawing.Point(12, 32);
            this.levelOneLabel.Name = "levelOneLabel";
            this.levelOneLabel.Size = new System.Drawing.Size(20, 17);
            this.levelOneLabel.TabIndex = 2;
            this.levelOneLabel.Text = "...";
            // 
            // levelTwoLabel
            // 
            this.levelTwoLabel.AutoSize = true;
            this.levelTwoLabel.Location = new System.Drawing.Point(12, 91);
            this.levelTwoLabel.Name = "levelTwoLabel";
            this.levelTwoLabel.Size = new System.Drawing.Size(20, 17);
            this.levelTwoLabel.TabIndex = 3;
            this.levelTwoLabel.Text = "...";
            // 
            // ProgressBar
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(668, 168);
            this.Controls.Add(this.levelTwoLabel);
            this.Controls.Add(this.levelOneLabel);
            this.Controls.Add(this.progressBarLevelTwo);
            this.Controls.Add(this.progressBarLevelOne);
            this.Name = "ProgressBar";
            this.Text = "Running...";
            this.Load += new System.EventHandler(this.ProgressBar_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ProgressBar progressBarLevelOne;
        private System.Windows.Forms.ProgressBar progressBarLevelTwo;
        private System.Windows.Forms.Label levelOneLabel;
        private System.Windows.Forms.Label levelTwoLabel;
    }
}

