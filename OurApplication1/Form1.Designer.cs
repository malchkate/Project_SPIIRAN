namespace OurApplication1
{
    partial class Form1
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
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.label14 = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.Latex_button = new System.Windows.Forms.Button();
            this.Excel = new System.Windows.Forms.Button();
            this.Word_List = new System.Windows.Forms.Button();
            this.Word_Table = new System.Windows.Forms.Button();
            this.Conferentions_List = new System.Windows.Forms.ListBox();
            this.Years_List = new System.Windows.Forms.ListBox();
            this.Authors_List = new System.Windows.Forms.ListBox();
            this.Categories_List = new System.Windows.Forms.ListBox();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.label = new System.Windows.Forms.ComboBox();
            this.category = new System.Windows.Forms.ComboBox();
            this.button3 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.info = new System.Windows.Forms.TextBox();
            this.way = new System.Windows.Forms.TextBox();
            this.pages = new System.Windows.Forms.TextBox();
            this.year = new System.Windows.Forms.TextBox();
            this.name = new System.Windows.Forms.TextBox();
            this.label12 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.tableView = new System.Windows.Forms.DataGridView();
            this.pathLabel = new System.Windows.Forms.Label();
            this.BDButton = new System.Windows.Forms.Button();
            this.AuthorsLabel = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.Authors = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage2.SuspendLayout();
            this.tabPage1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.tableView)).BeginInit();
            this.tabControl1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.label14);
            this.tabPage2.Controls.Add(this.label13);
            this.tabPage2.Controls.Add(this.label7);
            this.tabPage2.Controls.Add(this.label2);
            this.tabPage2.Controls.Add(this.Latex_button);
            this.tabPage2.Controls.Add(this.Excel);
            this.tabPage2.Controls.Add(this.Word_List);
            this.tabPage2.Controls.Add(this.Word_Table);
            this.tabPage2.Controls.Add(this.Conferentions_List);
            this.tabPage2.Controls.Add(this.Years_List);
            this.tabPage2.Controls.Add(this.Authors_List);
            this.tabPage2.Controls.Add(this.Categories_List);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(560, 489);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Извлечение";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(41, 253);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(115, 13);
            this.label14.TabIndex = 21;
            this.label14.Text = "Список конференций";
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(41, 200);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(76, 13);
            this.label13.TabIndex = 20;
            this.label13.Text = "Список годов";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(39, 109);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(88, 13);
            this.label7.TabIndex = 19;
            this.label7.Text = "Список авторов";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(39, 28);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(99, 13);
            this.label2.TabIndex = 18;
            this.label2.Text = "Список категорий";
            // 
            // Latex_button
            // 
            this.Latex_button.Location = new System.Drawing.Point(368, 340);
            this.Latex_button.Name = "Latex_button";
            this.Latex_button.Size = new System.Drawing.Size(157, 33);
            this.Latex_button.TabIndex = 17;
            this.Latex_button.Text = "LaTeX";
            this.Latex_button.UseVisualStyleBackColor = true;
            this.Latex_button.MouseClick += new System.Windows.Forms.MouseEventHandler(this.Latex_button_MouseClick);
            // 
            // Excel
            // 
            this.Excel.Location = new System.Drawing.Point(164, 379);
            this.Excel.Name = "Excel";
            this.Excel.Size = new System.Drawing.Size(165, 25);
            this.Excel.TabIndex = 15;
            this.Excel.Text = "Excel";
            this.Excel.UseVisualStyleBackColor = true;
            this.Excel.MouseClick += new System.Windows.Forms.MouseEventHandler(this.Excel_MouseClick);
            // 
            // Word_List
            // 
            this.Word_List.Location = new System.Drawing.Point(164, 347);
            this.Word_List.Name = "Word_List";
            this.Word_List.Size = new System.Drawing.Size(165, 26);
            this.Word_List.TabIndex = 14;
            this.Word_List.Text = "Word List";
            this.Word_List.UseVisualStyleBackColor = true;
            this.Word_List.MouseClick += new System.Windows.Forms.MouseEventHandler(this.Word_List_MouseClick);
            // 
            // Word_Table
            // 
            this.Word_Table.Location = new System.Drawing.Point(164, 313);
            this.Word_Table.Name = "Word_Table";
            this.Word_Table.Size = new System.Drawing.Size(165, 28);
            this.Word_Table.TabIndex = 13;
            this.Word_Table.Text = "Word Table";
            this.Word_Table.UseVisualStyleBackColor = true;
            this.Word_Table.MouseClick += new System.Windows.Forms.MouseEventHandler(this.Word_Table_MouseClick);
            // 
            // Conferentions_List
            // 
            this.Conferentions_List.FormattingEnabled = true;
            this.Conferentions_List.Location = new System.Drawing.Point(165, 245);
            this.Conferentions_List.Name = "Conferentions_List";
            this.Conferentions_List.ScrollAlwaysVisible = true;
            this.Conferentions_List.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple;
            this.Conferentions_List.Size = new System.Drawing.Size(221, 30);
            this.Conferentions_List.TabIndex = 12;
            // 
            // Years_List
            // 
            this.Years_List.FormattingEnabled = true;
            this.Years_List.Location = new System.Drawing.Point(164, 198);
            this.Years_List.Name = "Years_List";
            this.Years_List.ScrollAlwaysVisible = true;
            this.Years_List.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple;
            this.Years_List.Size = new System.Drawing.Size(222, 30);
            this.Years_List.TabIndex = 11;
            // 
            // Authors_List
            // 
            this.Authors_List.FormattingEnabled = true;
            this.Authors_List.Location = new System.Drawing.Point(165, 110);
            this.Authors_List.Name = "Authors_List";
            this.Authors_List.ScrollAlwaysVisible = true;
            this.Authors_List.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple;
            this.Authors_List.Size = new System.Drawing.Size(221, 82);
            this.Authors_List.TabIndex = 10;
            // 
            // Categories_List
            // 
            this.Categories_List.FormattingEnabled = true;
            this.Categories_List.Location = new System.Drawing.Point(164, 22);
            this.Categories_List.Name = "Categories_List";
            this.Categories_List.ScrollAlwaysVisible = true;
            this.Categories_List.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple;
            this.Categories_List.Size = new System.Drawing.Size(222, 82);
            this.Categories_List.TabIndex = 9;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.label);
            this.tabPage1.Controls.Add(this.category);
            this.tabPage1.Controls.Add(this.button3);
            this.tabPage1.Controls.Add(this.button2);
            this.tabPage1.Controls.Add(this.info);
            this.tabPage1.Controls.Add(this.way);
            this.tabPage1.Controls.Add(this.pages);
            this.tabPage1.Controls.Add(this.year);
            this.tabPage1.Controls.Add(this.name);
            this.tabPage1.Controls.Add(this.label12);
            this.tabPage1.Controls.Add(this.label11);
            this.tabPage1.Controls.Add(this.label10);
            this.tabPage1.Controls.Add(this.label9);
            this.tabPage1.Controls.Add(this.label8);
            this.tabPage1.Controls.Add(this.label6);
            this.tabPage1.Controls.Add(this.button1);
            this.tabPage1.Controls.Add(this.label5);
            this.tabPage1.Controls.Add(this.tableView);
            this.tabPage1.Controls.Add(this.pathLabel);
            this.tabPage1.Controls.Add(this.BDButton);
            this.tabPage1.Controls.Add(this.AuthorsLabel);
            this.tabPage1.Controls.Add(this.label4);
            this.tabPage1.Controls.Add(this.Authors);
            this.tabPage1.Controls.Add(this.label3);
            this.tabPage1.Controls.Add(this.label1);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(560, 489);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Добавление";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // label
            // 
            this.label.FormattingEnabled = true;
            this.label.Items.AddRange(new object[] {
            "печ.",
            "рук. ",
            "электр.",
            "нет метки"});
            this.label.Location = new System.Drawing.Point(116, 152);
            this.label.Name = "label";
            this.label.Size = new System.Drawing.Size(82, 21);
            this.label.TabIndex = 57;
            // 
            // category
            // 
            this.category.FormattingEnabled = true;
            this.category.Items.AddRange(new object[] {
            "тезисы",
            "доклад на конференции",
            "отчет",
            "препринт"});
            this.category.Location = new System.Drawing.Point(180, 105);
            this.category.Name = "category";
            this.category.Size = new System.Drawing.Size(129, 21);
            this.category.TabIndex = 56;
            this.category.SelectedIndexChanged += new System.EventHandler(this.category_SelectedIndexChanged);
            // 
            // button3
            // 
            this.button3.Enabled = false;
            this.button3.Location = new System.Drawing.Point(344, 225);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(137, 23);
            this.button3.TabIndex = 55;
            this.button3.Text = "Добавить поля отчета";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button2
            // 
            this.button2.Enabled = false;
            this.button2.Location = new System.Drawing.Point(22, 234);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(287, 38);
            this.button2.TabIndex = 54;
            this.button2.Text = "Добавить запись";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // info
            // 
            this.info.Location = new System.Drawing.Point(22, 199);
            this.info.Name = "info";
            this.info.Size = new System.Drawing.Size(287, 20);
            this.info.TabIndex = 53;
            // 
            // way
            // 
            this.way.Location = new System.Drawing.Point(204, 152);
            this.way.Name = "way";
            this.way.Size = new System.Drawing.Size(105, 20);
            this.way.TabIndex = 51;
            // 
            // pages
            // 
            this.pages.Location = new System.Drawing.Point(24, 152);
            this.pages.Name = "pages";
            this.pages.Size = new System.Drawing.Size(83, 20);
            this.pages.TabIndex = 47;
            // 
            // year
            // 
            this.year.Location = new System.Drawing.Point(24, 106);
            this.year.Name = "year";
            this.year.Size = new System.Drawing.Size(137, 20);
            this.year.TabIndex = 32;
            // 
            // name
            // 
            this.name.Location = new System.Drawing.Point(24, 64);
            this.name.Name = "name";
            this.name.Size = new System.Drawing.Size(285, 20);
            this.name.TabIndex = 29;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(19, 179);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(101, 13);
            this.label12.TabIndex = 52;
            this.label12.Text = "Доп. информация:";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(201, 134);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(49, 13);
            this.label11.TabIndex = 50;
            this.label11.Text = "Ссылка:";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(113, 134);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(42, 13);
            this.label10.TabIndex = 48;
            this.label10.Text = "Метка:";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(19, 134);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(88, 13);
            this.label9.TabIndex = 46;
            this.label9.Text = "Кол-во страниц:";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(177, 90);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(63, 13);
            this.label8.TabIndex = 44;
            this.label8.Text = "Категория:";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(339, 179);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(142, 13);
            this.label6.TabIndex = 43;
            this.label6.Text = "*выбранная конференция*";
            // 
            // button1
            // 
            this.button1.Enabled = false;
            this.button1.Location = new System.Drawing.Point(344, 149);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(137, 23);
            this.button1.TabIndex = 41;
            this.button1.Text = "Выбрать конференцию";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click_2);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(341, 90);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(113, 13);
            this.label5.TabIndex = 40;
            this.label5.Text = "*выбранные авторы*";
            // 
            // tableView
            // 
            this.tableView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.tableView.Location = new System.Drawing.Point(22, 288);
            this.tableView.Name = "tableView";
            this.tableView.Size = new System.Drawing.Size(503, 179);
            this.tableView.TabIndex = 39;
            // 
            // pathLabel
            // 
            this.pathLabel.AutoSize = true;
            this.pathLabel.Location = new System.Drawing.Point(177, 27);
            this.pathLabel.Name = "pathLabel";
            this.pathLabel.Size = new System.Drawing.Size(132, 13);
            this.pathLabel.TabIndex = 38;
            this.pathLabel.Text = "*название базы данных*";
            // 
            // BDButton
            // 
            this.BDButton.Location = new System.Drawing.Point(24, 22);
            this.BDButton.Name = "BDButton";
            this.BDButton.Size = new System.Drawing.Size(137, 23);
            this.BDButton.TabIndex = 37;
            this.BDButton.Text = "Выбор базы данных";
            this.BDButton.UseVisualStyleBackColor = true;
            this.BDButton.Click += new System.EventHandler(this.BDButton_Click);
            // 
            // AuthorsLabel
            // 
            this.AuthorsLabel.AutoSize = true;
            this.AuthorsLabel.Location = new System.Drawing.Point(81, 72);
            this.AuthorsLabel.Name = "AuthorsLabel";
            this.AuthorsLabel.Size = new System.Drawing.Size(0, 13);
            this.AuthorsLabel.TabIndex = 36;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(75, 72);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(0, 13);
            this.label4.TabIndex = 35;
            // 
            // Authors
            // 
            this.Authors.Enabled = false;
            this.Authors.Location = new System.Drawing.Point(344, 62);
            this.Authors.Name = "Authors";
            this.Authors.Size = new System.Drawing.Size(137, 23);
            this.Authors.TabIndex = 33;
            this.Authors.Text = "Выбрать авторов";
            this.Authors.UseVisualStyleBackColor = true;
            this.Authors.Click += new System.EventHandler(this.Authors_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(21, 90);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(90, 13);
            this.label3.TabIndex = 31;
            this.label3.Text = "Год публикации:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(21, 48);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(60, 13);
            this.label1.TabIndex = 30;
            this.label1.Text = "Название:";
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Location = new System.Drawing.Point(-1, -1);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(568, 515);
            this.tabControl1.TabIndex = 0;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(565, 514);
            this.Controls.Add(this.tabControl1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.tableView)).EndInit();
            this.tabControl1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Button Latex_button;
        private System.Windows.Forms.Button Excel;
        private System.Windows.Forms.Button Word_List;
        private System.Windows.Forms.Button Word_Table;
        private System.Windows.Forms.ListBox Conferentions_List;
        private System.Windows.Forms.ListBox Years_List;
        private System.Windows.Forms.ListBox Authors_List;
        private System.Windows.Forms.ListBox Categories_List;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.ComboBox label;
        private System.Windows.Forms.ComboBox category;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.TextBox info;
        private System.Windows.Forms.TextBox way;
        private System.Windows.Forms.TextBox pages;
        private System.Windows.Forms.TextBox year;
        private System.Windows.Forms.TextBox name;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.DataGridView tableView;
        private System.Windows.Forms.Label pathLabel;
        private System.Windows.Forms.Button BDButton;
        private System.Windows.Forms.Label AuthorsLabel;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button Authors;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label2;


    }
}

