﻿
namespace MainForm
{
    partial class Window
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.MainMenu = new System.Windows.Forms.MenuStrip();
            this.execute = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.label2 = new System.Windows.Forms.Label();
            this.discipline = new System.Windows.Forms.TextBox();
            this.workBox = new System.Windows.Forms.ComboBox();
            this.disciplineBox = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.theme = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.pictureBox2 = new System.Windows.Forms.PictureBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.prepod = new System.Windows.Forms.TextBox();
            this.prepodIniz = new System.Windows.Forms.TextBox();
            this.prepodInfo = new System.Windows.Forms.TextBox();
            this.pictureBox3 = new System.Windows.Forms.PictureBox();
            this.label9 = new System.Windows.Forms.Label();
            this.args = new System.Windows.Forms.TextBox();
            this.pictureBox4 = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox4)).BeginInit();
            this.SuspendLayout();
            // 
            // MainMenu
            // 
            this.MainMenu.Location = new System.Drawing.Point(0, 0);
            this.MainMenu.Name = "MainMenu";
            this.MainMenu.Size = new System.Drawing.Size(530, 24);
            this.MainMenu.TabIndex = 0;
            this.MainMenu.Text = "menuStrip1";
            // 
            // execute
            // 
            this.execute.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.execute.Location = new System.Drawing.Point(460, 345);
            this.execute.Name = "execute";
            this.execute.Size = new System.Drawing.Size(58, 23);
            this.execute.TabIndex = 1;
            this.execute.Text = "Запуск";
            this.execute.UseVisualStyleBackColor = true;
            this.execute.Click += new System.EventHandler(this.execute_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Calibri", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.label1.Location = new System.Drawing.Point(12, 24);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(220, 23);
            this.label1.TabIndex = 2;
            this.label1.Text = "Конфигурация генератора";
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackColor = System.Drawing.SystemColors.ControlDark;
            this.pictureBox1.Location = new System.Drawing.Point(12, 50);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(506, 1);
            this.pictureBox1.TabIndex = 3;
            this.pictureBox1.TabStop = false;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.label2.Location = new System.Drawing.Point(12, 60);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(73, 14);
            this.label2.TabIndex = 4;
            this.label2.Text = "Дисциплина:";
            this.label2.UseMnemonic = false;
            // 
            // discipline
            // 
            this.discipline.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.discipline.Location = new System.Drawing.Point(91, 57);
            this.discipline.Name = "discipline";
            this.discipline.Size = new System.Drawing.Size(427, 22);
            this.discipline.TabIndex = 5;
            this.discipline.TextChanged += new System.EventHandler(this.discChanged);
            // 
            // workBox
            // 
            this.workBox.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.workBox.FormattingEnabled = true;
            this.workBox.Items.AddRange(new object[] {
            "Эссе",
            "Реферат",
            "Контрольная работа",
            "Расчётно-графическая работа",
            "Курсовой проект"});
            this.workBox.Location = new System.Drawing.Point(91, 85);
            this.workBox.Name = "workBox";
            this.workBox.Size = new System.Drawing.Size(183, 22);
            this.workBox.TabIndex = 6;
            this.workBox.SelectedIndexChanged += new System.EventHandler(this.wbChanged);
            // 
            // disciplineBox
            // 
            this.disciplineBox.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.disciplineBox.FormattingEnabled = true;
            this.disciplineBox.Items.AddRange(new object[] {
            "Дисциплина",
            "Междисциплинарный курс",
            "Модуль"});
            this.disciplineBox.Location = new System.Drawing.Point(370, 85);
            this.disciplineBox.Name = "disciplineBox";
            this.disciplineBox.Size = new System.Drawing.Size(148, 22);
            this.disciplineBox.TabIndex = 7;
            this.disciplineBox.SelectedIndexChanged += new System.EventHandler(this.dbChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.label3.Location = new System.Drawing.Point(12, 88);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(70, 14);
            this.label3.TabIndex = 8;
            this.label3.Text = "Тип работы:";
            this.label3.UseMnemonic = false;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.label4.Location = new System.Drawing.Point(280, 88);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(84, 14);
            this.label4.TabIndex = 9;
            this.label4.Text = "Тип предмета:";
            this.label4.UseMnemonic = false;
            // 
            // theme
            // 
            this.theme.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.theme.Location = new System.Drawing.Point(91, 113);
            this.theme.Multiline = true;
            this.theme.Name = "theme";
            this.theme.Size = new System.Drawing.Size(427, 69);
            this.theme.TabIndex = 10;
            this.theme.TextChanged += new System.EventHandler(this.themeChanged);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.label5.Location = new System.Drawing.Point(12, 116);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(37, 14);
            this.label5.TabIndex = 11;
            this.label5.Text = "Тема:";
            this.label5.UseMnemonic = false;
            // 
            // pictureBox2
            // 
            this.pictureBox2.BackColor = System.Drawing.SystemColors.ControlDark;
            this.pictureBox2.Location = new System.Drawing.Point(12, 188);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new System.Drawing.Size(506, 1);
            this.pictureBox2.TabIndex = 12;
            this.pictureBox2.TabStop = false;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.label6.Location = new System.Drawing.Point(12, 198);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(84, 14);
            this.label6.TabIndex = 13;
            this.label6.Text = "Руководитель:";
            this.label6.UseMnemonic = false;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.label7.Location = new System.Drawing.Point(300, 198);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(64, 14);
            this.label7.TabIndex = 14;
            this.label7.Text = "Инициалы:";
            this.label7.UseMnemonic = false;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.label8.Location = new System.Drawing.Point(12, 220);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(199, 14);
            this.label8.TabIndex = 15;
            this.label8.Text = "Должность, учёная степень, звание:";
            this.label8.UseMnemonic = false;
            // 
            // prepod
            // 
            this.prepod.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.prepod.Location = new System.Drawing.Point(112, 195);
            this.prepod.Name = "prepod";
            this.prepod.Size = new System.Drawing.Size(182, 22);
            this.prepod.TabIndex = 16;
            this.prepod.TextChanged += new System.EventHandler(this.advisorChanged);
            // 
            // prepodIniz
            // 
            this.prepodIniz.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.prepodIniz.Location = new System.Drawing.Point(370, 195);
            this.prepodIniz.Name = "prepodIniz";
            this.prepodIniz.Size = new System.Drawing.Size(148, 22);
            this.prepodIniz.TabIndex = 17;
            this.prepodIniz.TextChanged += new System.EventHandler(this.inizChanged);
            // 
            // prepodInfo
            // 
            this.prepodInfo.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.prepodInfo.Location = new System.Drawing.Point(12, 237);
            this.prepodInfo.Multiline = true;
            this.prepodInfo.Name = "prepodInfo";
            this.prepodInfo.Size = new System.Drawing.Size(506, 46);
            this.prepodInfo.TabIndex = 18;
            this.prepodInfo.TextChanged += new System.EventHandler(this.infoChanged);
            // 
            // pictureBox3
            // 
            this.pictureBox3.BackColor = System.Drawing.SystemColors.ControlDark;
            this.pictureBox3.Location = new System.Drawing.Point(12, 289);
            this.pictureBox3.Name = "pictureBox3";
            this.pictureBox3.Size = new System.Drawing.Size(506, 1);
            this.pictureBox3.TabIndex = 19;
            this.pictureBox3.TabStop = false;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.label9.Location = new System.Drawing.Point(12, 293);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(205, 14);
            this.label9.TabIndex = 20;
            this.label9.Text = "Индивидуальные аргументы запуска:";
            this.label9.UseMnemonic = false;
            // 
            // args
            // 
            this.args.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.args.Location = new System.Drawing.Point(12, 310);
            this.args.Name = "args";
            this.args.Size = new System.Drawing.Size(506, 22);
            this.args.TabIndex = 21;
            this.args.TextChanged += new System.EventHandler(this.argsChanged);
            // 
            // pictureBox4
            // 
            this.pictureBox4.BackColor = System.Drawing.SystemColors.ControlDark;
            this.pictureBox4.Location = new System.Drawing.Point(12, 338);
            this.pictureBox4.Name = "pictureBox4";
            this.pictureBox4.Size = new System.Drawing.Size(506, 1);
            this.pictureBox4.TabIndex = 22;
            this.pictureBox4.TabStop = false;
            // 
            // Window
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(530, 379);
            this.Controls.Add(this.pictureBox4);
            this.Controls.Add(this.args);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.pictureBox3);
            this.Controls.Add(this.prepodInfo);
            this.Controls.Add(this.prepodIniz);
            this.Controls.Add(this.prepod);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.pictureBox2);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.theme);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.disciplineBox);
            this.Controls.Add(this.workBox);
            this.Controls.Add(this.discipline);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.execute);
            this.Controls.Add(this.MainMenu);
            this.MainMenuStrip = this.MainMenu;
            this.Name = "Window";
            this.Text = "STO Generator 2.0";
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox4)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip MainMenu;
        private System.Windows.Forms.Button execute;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox discipline;
        private System.Windows.Forms.ComboBox workBox;
        private System.Windows.Forms.ComboBox disciplineBox;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox theme;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.PictureBox pictureBox2;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox prepod;
        private System.Windows.Forms.TextBox prepodIniz;
        private System.Windows.Forms.TextBox prepodInfo;
        private System.Windows.Forms.PictureBox pictureBox3;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.TextBox args;
        private System.Windows.Forms.PictureBox pictureBox4;
    }
}

