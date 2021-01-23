namespace SpreadsheetGUI
{
    partial class DataGraph
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
            System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea1 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
            System.Windows.Forms.DataVisualization.Charting.Legend legend1 = new System.Windows.Forms.DataVisualization.Charting.Legend();
            this.dataGraphChart = new System.Windows.Forms.DataVisualization.Charting.Chart();
            ((System.ComponentModel.ISupportInitialize)(this.dataGraphChart)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGraphChart
            // 
            chartArea1.Name = "ChartArea1";
            this.dataGraphChart.ChartAreas.Add(chartArea1);
            this.dataGraphChart.Dock = System.Windows.Forms.DockStyle.Fill;
            legend1.Name = "Legend1";
            this.dataGraphChart.Legends.Add(legend1);
            this.dataGraphChart.Location = new System.Drawing.Point(0, 0);
            this.dataGraphChart.Name = "dataGraphChart";
            this.dataGraphChart.Size = new System.Drawing.Size(784, 561);
            this.dataGraphChart.TabIndex = 0;
            this.dataGraphChart.Text = "Data Graph";
            // 
            // DataGraph
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.ClientSize = new System.Drawing.Size(784, 561);
            this.Controls.Add(this.dataGraphChart);
            this.Name = "DataGraph";
            this.Text = "DataGraph";
            this.Load += new System.EventHandler(this.DataGraph_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGraphChart)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataVisualization.Charting.Chart dataGraphChart;
    }
}