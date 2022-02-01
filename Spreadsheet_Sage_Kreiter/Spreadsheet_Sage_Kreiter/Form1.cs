// <copyright file="Form1.cs" company="Sage Krieter - 11659212">
// Copyright (c) Sage Krieter. All rights reserved.
// </copyright>

namespace Spreadsheet_Sage_Kreiter
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Data;
    using System.Drawing;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using System.Windows.Forms;
    using System.Xml;
    using SpreadsheetEngine;

    /// <summary>
    /// Class to control the form and spreadsheet.
    /// </summary>
    public partial class Form1 : Form
    {
        private Spreadsheet spreadsheet;

        /// <summary>
        /// Initializes a new instance of the <see cref="Form1"/> class.
        /// </summary>
        public Form1()
        {
            this.spreadsheet = new Spreadsheet(50, 26);
            this.spreadsheet.PropertyChanged += this.SpreadsheetPropertyChanged;
            this.InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.dataGridView1.ColumnCount = 26;
            this.dataGridView1.RowCount = 50;
            this.undoToolStripMenuItem.Enabled = false;
            this.redoToolStripMenuItem.Enabled = false;
            int j = 0;
            for (char i = 'A'; i <= 'Z'; i++, j++)
            {
                this.dataGridView1.Columns[j].Name = i.ToString();
                this.dataGridView1.Columns[j].HeaderText = i.ToString();
            }

            this.dataGridView1.RowHeadersWidth = 60;
            for (int i = 1; i <= 50; i++)
            {
                this.dataGridView1.Rows[i - 1].HeaderCell.Value = i.ToString();
            }
        }

        /// <summary>
        /// When button is clicked program comes here.
        /// </summary>
        /// <param name="sender">Button clicked.</param>
        /// <param name="e">Event notifier.</param>
        private void button1_Click(object sender, EventArgs e)
        {
            var rand = new Random();
            int rows = 0, columns = 0;
            char columnLetter = 'a';

            for (int i = 0; i < 50; i++)
            {
                rows = rand.Next(0, 49);
                columns = rand.Next(0, 25);
                columns += 65;
                columnLetter = Convert.ToChar(columns);
                this.spreadsheet.SetCell(rows, columnLetter, "Hello!", false);
            }

            for (int i = 1; i <= 50; i++)
            {
                this.spreadsheet.SetCell(i - 1, 'B', "This is cell B" + i, false);
            }

            for (int i = 1; i <= 50; i++)
            {
                this.spreadsheet.SetCell(i - 1, 'A', "=B" + i, false);
            }
        }

        private void SpreadsheetPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            this.dataGridView1.Rows[(sender as CellHelp).GetRow()].Cells[(sender as Cell).GetColumn()].Value = (sender as Cell).GetValue();
            this.dataGridView1.Rows[(sender as CellHelp).GetRow()].Cells[(sender as Cell).GetColumn()].Style.BackColor = Color.FromArgb((int)(sender as Cell).Color);
        }

        private void dataGridView1_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            this.dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = this.spreadsheet.GetCell(e.RowIndex, e.ColumnIndex).GetText();
            this.spreadsheet.AddUndoText(this.spreadsheet.GetCell(e.RowIndex, e.ColumnIndex));
            this.undoToolStripMenuItem.Text = "Undo Text Change";
            this.undoToolStripMenuItem.Enabled = true;
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            char col = 'a';
            for (int i = 0; i < e.ColumnIndex; i++)
            {
                col++;
            }

            int error = 0;
            List<Cell> originalRefer = new List<Cell>(this.spreadsheet.GetCell(e.RowIndex, e.ColumnIndex).Refer);

            if (this.dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value == null)
            {
                error = this.spreadsheet.SetCell(e.RowIndex, col, string.Empty, false);
            }
            else
            {
                error = this.spreadsheet.SetCell(e.RowIndex, col, this.dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString(), false);
            }

            this.dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = this.spreadsheet.GetValue(e.RowIndex, col);
            List<Cell> list = new List<Cell>();

            if (error == 0)
            {
                this.UpdateCells(originalRefer);
            }
        }

        private void UpdateCells(List<Cell> list)
        {
            foreach (Cell cell in list)
            {
                char column = (char)(cell.GetColumn() + 65);
                List<Cell> refers = new List<Cell>(this.spreadsheet.GetCell(cell.GetRow(), cell.GetColumn()).Refer);

                this.spreadsheet.SetCell(cell.GetRow(), column, cell.GetText(), true);
                this.dataGridView1.Rows[cell.GetRow()].Cells[cell.GetColumn()].Value = this.spreadsheet.GetValue(cell.GetRow(), column);
                this.UpdateCells(refers); // Recursion to check cells that refer to this cell
            }
        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void colorsToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void undoRedoToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void changeColorToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ColorDialog newColor = new ColorDialog();
            if (newColor.ShowDialog() == DialogResult.OK)
            {
                int cellSelect = this.dataGridView1.GetCellCount(DataGridViewElementStates.Selected);
                if (cellSelect > 0)
                {
                    Stack<Tuple<int, int>> cells = new Stack<Tuple<int, int>>();
                    Stack<uint> colors = new Stack<uint>();
                    for (int i = 0; i < cellSelect; i++)
                    {
                        int row = this.dataGridView1.SelectedCells[i].RowIndex;
                        int col = this.dataGridView1.SelectedCells[i].ColumnIndex;

                        Tuple<int, int> rowCol = new Tuple<int, int>(row, col);
                        cells.Push(rowCol);
                        colors.Push(this.spreadsheet.GetCell(row, col).Color);

                        this.spreadsheet.SetColor((uint)newColor.Color.ToArgb(), row, col);
                    }

                    this.spreadsheet.AddUndoColor(colors, cells);
                    this.undoToolStripMenuItem.Enabled = true;
                    this.undoToolStripMenuItem.Text = "Undo Color Change";
                }
            }
        }

        private void undoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string title = this.spreadsheet.Undo();
            if (this.spreadsheet.IsUndoEmpty())
            {
                this.undoToolStripMenuItem.Enabled = false;
            }

            this.redoToolStripMenuItem.Enabled = true;

            string redoText = this.undoToolStripMenuItem.Text;
            redoText = redoText.Substring(4);
            this.redoToolStripMenuItem.Text = "Redo" + redoText;
            this.undoToolStripMenuItem.Text = title;
        }

        private void redoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string title = this.spreadsheet.Redo();
            if (this.spreadsheet.IsRedoEmpty())
            {
                this.redoToolStripMenuItem.Enabled = false;
            }

            this.redoToolStripMenuItem.Text = title;
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.DefaultExt = "xml";

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                FileStream file = new FileStream(saveFileDialog.FileName, FileMode.Create, FileAccess.Write);
                this.spreadsheet.SaveXML(file);
                file.Dispose();
            }
        }

        private void loadToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                FileStream file = new FileStream(openFileDialog.FileName, FileMode.Open, FileAccess.Read);
                this.spreadsheet.LoadXML(file);
                file.Dispose();
            }
        }
    }
}
