// <copyright file="Spreadsheet.cs" company="Sage Kreiter - 11659212">
// Copyright (c) Sage Kreiter. All rights reserved.
// </copyright>

namespace SpreadsheetEngine
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Drawing;
    using System.IO;
    using System.Linq;
    using System.Runtime.CompilerServices;
    using System.Text;
    using System.Text.RegularExpressions;
    using System.Xml;

    /// <summary>
    /// Class that hold the whole sheet.
    /// </summary>
    public class Spreadsheet : INotifyPropertyChanged
    {
        private CellHelp[,] sheet;
        private int rows;
        private int columns;
        private Stack<UndoRedoCollection> undos;
        private Stack<UndoRedoCollection> redos;

        /// <summary>
        /// Initializes a new instance of the <see cref="Spreadsheet"/> class.
        /// </summary>
        /// <param name="newRows">How many rows there will be.</param>
        /// <param name="newColumns">How many columns there will be.</param>
        public Spreadsheet(int newRows, int newColumns)
        {
            this.undos = new Stack<UndoRedoCollection>();
            this.redos = new Stack<UndoRedoCollection>();
            this.rows = newRows;
            this.columns = newColumns;
            this.sheet = new CellHelp[this.rows, this.columns];

            for (int i = 0; i < newRows; i++)
            {
                for (int j = 0; j < newColumns; j++)
                {
                    this.sheet[i, j] = new CellHelp(i, j);
                    this.sheet[i, j].SetText(string.Empty);
                }
            }
        }

        /// <summary>
        /// Event for when a cell is changed.
        /// </summary>
        public event PropertyChangedEventHandler PropertyChanged;

        /// <summary>
        /// Gets whole cell.
        /// </summary>
        /// <param name="row1">Row requested.</param>
        /// <param name="column1">Column Requested.</param>
        /// <returns>The cell.</returns>
        public Cell GetCell(int row1, int column1)
        {
            if (row1 < this.rows && column1 < this.columns)
            {
                return this.sheet[row1, column1];
            }

            return null;
        }

        /// <summary>
        /// Sets the value of a cell.
        /// </summary>
        /// <param name="row1">Row requested.</param>
        /// <param name="columnLetter">Column Requested.</param>
        /// <param name="value">Value to be set into cell.</param>
        /// <param name="refBool">Says whether or not references need to be updated.</param>
        /// <returns>The error value.</returns>
        public int SetCell(int row1, char columnLetter, string value, bool refBool)
        {
            bool evaluate = false;
            int error = 0;
            char check = '0';
            int column1 = this.CharToNum(columnLetter);
            string words = string.Empty;
            int colToGet = 0, rowToGet = 0;
            List<Cell> myList = new List<Cell>();
            this.sheet[row1, column1].ReferTo.Clear();
            this.sheet[row1, column1].Refer.Clear();
            if (row1 < this.rows && column1 < this.columns)
            {
                this.sheet[row1, column1].Changed = true;
                this.sheet[row1, column1].SetText(value);
                if (value != string.Empty && value[0] == '=')
                {
                    // Enter this if statement if there is a self reference detected and skip rest of evaluation
                    if (this.CheckSelfReference(row1, columnLetter, value))
                    {
                        this.sheet[row1, column1].SetText(value);
                        error = 1;
                    }
                    else
                    {
                        check = value[0];
                        evaluate = true;
                        value = value.Substring(1);
                        this.sheet[row1, column1].ExpressionTree = value; // Setting the expression tree to a new expression

                        // Getting all the values other than operators out of the expression and entering cell values as variables in the tree
                        string var = string.Empty;
                        for (int i = 0; i < value.Length; i++)
                        {
                            if (char.IsLetter(value[i]))
                            {
                                var = string.Empty;
                                var += value[i]; // Starting a new variable to be inserted
                            }
                            else if (char.IsDigit(value[i]) && var != string.Empty)
                            {
                                var += value[i];
                                colToGet = this.CharToNum(var[0]);
                                rowToGet = (int)char.GetNumericValue(var[1]);
                                double cellToGet = 0;

                                // Got cells value to set this cell to
                                if (rowToGet - 1 < this.rows && colToGet < this.columns && rowToGet - 1 >= 0 && colToGet >= 0 && !double.TryParse(this.sheet[rowToGet - 1, colToGet].GetValue(), out cellToGet))
                                {
                                    evaluate = false;
                                    words = this.sheet[rowToGet - 1, colToGet].GetValue();
                                }

                                this.sheet[row1, column1].SetVariables(var, cellToGet);
                                var = string.Empty;

                                // Adding references to the cell if it is a valid entry into the dict
                                if (rowToGet - 1 < this.rows && colToGet < this.columns && !refBool && rowToGet - 1 >= 0 && colToGet >= 0)
                                {
                                    this.sheet[rowToGet - 1, colToGet].Refer.Add(this.sheet[row1, column1]);
                                    this.sheet[row1, column1].ReferTo.Add(this.sheet[rowToGet - 1, colToGet]);
                                }
                            }
                        }
                    }
                }
                else if (value == string.Empty)
                {
                    this.sheet[row1, column1].SetText(value);
                }

                if (this.CheckBadReference(row1, column1))
                {
                    evaluate = false;
                    error = 2;
                }
                else if (this.CheckCircularReference(row1, column1))
                {
                    evaluate = false;
                    error = 3;
                }

                this.NotifyCellPropertyChanged(this.sheet[row1, column1], new PropertyChangedEventArgs("CellText"), evaluate, words, false, error);
                return error;
            }

            return 0;
        }

        /// <summary>
        /// Checks if a cell references itself by checking its text before setting the text in SetCell().
        /// </summary>
        /// <param name="row">Cells row.</param>
        /// <param name="column">Cells column.</param>
        /// <param name="text">Text to be set into the cell, starts with =.</param>
        /// <returns>True if text contains the cell and False if not.</returns>
        public bool CheckSelfReference(int row, char column, string text)
        {
            string cell = column + Convert.ToString(row + 1);
            string cellcap = char.ToUpper(column) + Convert.ToString(row + 1);
            if (text.Contains(cell) || text.Contains(cellcap))
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// Checking if reference contains a value outside of the spreadsheet.
        /// </summary>
        /// <param name="row">Cells row.</param>
        /// <param name="column">Cells column.</param>
        /// <returns>True if cell references a bad cell and False if not.</returns>
        public bool CheckBadReference(int row, int column)
        {
            ExpressionTree tree = this.sheet[row, column].GetTree;
            foreach (KeyValuePair<string, double> value in tree.Dict)
            {
                string var = value.Key;
                string colString = string.Empty;
                int i = 0, rowVar = 0, colVar = 0;
                Regex r = new Regex("[a-zA-Z0-9]");

                if (!r.IsMatch(var))
                {
                    return true;
                }

                // Getting the column value from the dictionary word
                for (; i < var.Length && char.IsLetter(var[i]); i++)
                {
                    colString += var[i];
                }

                // If this is true it means no row number was entered or the column letter was entered in the wromg order so return true
                if (i == var.Length || i == 0)
                {
                    return true;
                }

                // Getting the column number the letters refer to
                for (int j = 1; j < i; j++)
                {
                    colVar += 25;
                }

                colVar += this.CharToNum(var[i - 1]);

                rowVar = Convert.ToInt32(var.Substring(i)) - 1;

                if ((rowVar >= this.rows || colVar >= this.columns) || (rowVar < 0 || colVar < 0))
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Cehcks if there is a circular reference.
        /// </summary>
        /// <param name="row">Cell row.</param>
        /// <param name="column">Cell column.</param>
        /// <returns>True if there is an error otherwise false.</returns>
        public bool CheckCircularReference(int row, int column)
        {
            foreach (Cell cell in this.sheet[row, column].ReferTo)
            {
                if (cell.ReferTo.Contains(this.sheet[row, column]))
                {
                    return true;
                }
                else if (this.CheckCircularReferenceHelp(cell.GetRow(), cell.GetColumn(), row, column))
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Helps the circular reference checker with recursion.
        /// </summary>
        /// <param name="row">Cell row.</param>
        /// <param name="column">Cell column.</param>
        /// <param name="lookRow">Row to check.</param>
        /// <param name="lookCol">Column to check.</param>
        /// <returns>True if there is an error otherwise false.</returns>
        public bool CheckCircularReferenceHelp(int row, int column, int lookRow, int lookCol)
        {
            foreach (Cell cell in this.sheet[row, column].ReferTo)
            {
                if (cell.ReferTo.Contains(this.sheet[lookRow, lookCol]))
                {
                    return true;
                }
                else if (this.CheckCircularReferenceHelp(cell.GetRow(), cell.GetColumn(), lookRow, lookCol))
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Sets the cell to a new color.
        /// </summary>
        /// <param name="color">New color to be set.</param>
        /// <param name="row">Row color to be set.</param>
        /// <param name="col">Column color to be set.</param>
        public void SetColor(uint color, int row, int col)
        {
            this.sheet[row, col].Changed = true;
            this.sheet[row, col].Color = color;
            this.NotifyCellPropertyChanged(this.sheet[row, col], new PropertyChangedEventArgs("CellText"), false, string.Empty, true, 0);
        }

        /// <summary>
        /// Gets the value of a cell.
        /// </summary>
        /// <param name="row1">Row requested.</param>
        /// <param name="columnLetter">Column requested.</param>
        /// <returns>String value of cell.</returns>
        public string GetValue(int row1, char columnLetter)
        {
            int columnNum = this.CharToNum(columnLetter);
            if (row1 < this.rows && columnNum < this.columns)
            {
                return this.sheet[row1, columnNum].GetValue();
            }

            return null;
        }

        /// <summary>
        /// Transfers a column char value to an int.
        /// </summary>
        /// <param name="character">Character to be transferred.</param>
        /// <returns>Column number.</returns>
        public int CharToNum(char character)
        {
            int val = character;
            if (val > 96)
            {
                return val - 97;
            }
            else if (val < 58)
            {
                return val - 48;
            }

            return val - 65;
        }

        /// <summary>
        /// Gets number of columns.
        /// </summary>
        /// <returns>Number of Columns.</returns>
        public int ColumnCount()
        {
            return this.columns;
        }

        /// <summary>
        /// Gets number of rows.
        /// </summary>
        /// <returns>Number of rows.</returns>
        public int RowCount()
        {
            return this.rows;
        }

        /// <summary>
        /// Adds an UndoText object to the stack of undos.
        /// </summary>
        /// <param name="cell">Cell that was changed.</param>
        public void AddUndoText(Cell cell)
        {
                UndoText undo = new UndoText("Text Change", cell.GetText(), cell.GetRow(), cell.GetColumn());
                this.undos.Push(undo);
        }

        /// <summary>
        /// Adds an UndoText object to the stack of redos.
        /// </summary>
        /// <param name="cell">Cell that was changed.</param>
        public void AddRedoText(Cell cell)
        {
            UndoText undo = new UndoText("Text Change", cell.GetText(), cell.GetRow(), cell.GetColumn());
            this.undos.Push(undo);
        }

        /// <summary>
        /// Adds an UndoColor object to the stack of undos.
        /// </summary>
        /// <param name="color">The original colors of all changed cells.</param>
        /// <param name="newCells">The rows and columns of the changed cells.</param>
        public void AddUndoColor(Stack<uint> color, Stack<Tuple<int, int>> newCells)
        {
            UndoColor undo = new UndoColor(color, newCells);
            this.undos.Push(undo);
        }

        /// <summary>
        /// Adds an UndoColor object to the stack of redos.
        /// </summary>
        /// <param name="color">The original colors of all changed cells.</param>
        /// <param name="newCells">The rows and columns of the changed cells.</param>
        public void AddRedoColor(Stack<uint> color, Stack<Tuple<int, int>> newCells)
        {
            UndoColor redo = new UndoColor(color, newCells);
            this.redos.Push(redo);
        }

        /// <summary>
        /// Pops the top of the undo stack and undoes what happened.
        /// </summary>
        /// <returns>String of what the next undo type is.</returns>
        public string Undo()
        {
            Type type = this.undos.Peek().GetType();

            if (type.Equals(typeof(UndoText)))
            {
                UndoText cell = (UndoText)this.undos.Pop();

                UndoText redoCell = new UndoText("Text Change", this.GetCell(cell.Row, cell.Column).GetText(), cell.Row, cell.Column);

                this.redos.Push(redoCell);
                char colLetter = (char)(cell.Column + 65);

                this.SetCell(cell.Row, colLetter, cell.Text, false);
            }
            else if (type.Equals(typeof(UndoColor)))
            {
                UndoColor cells = (UndoColor)this.undos.Pop();
                Stack<uint> redoColors = new Stack<uint>();
                Stack<Tuple<int, int>> redoCells = new Stack<Tuple<int, int>>();

                for (int i = 0; i < cells.Count; i++)
                {
                    Tuple<int, int> rowCol = cells.Cells;
                    redoColors.Push(this.sheet[rowCol.Item1, rowCol.Item2].Color);
                    Tuple<int, int> redoRowCol = new Tuple<int, int>(rowCol.Item1, rowCol.Item2);
                    redoCells.Push(redoRowCol);
                    this.SetColor(cells.Color, rowCol.Item1, rowCol.Item2);
                }

                UndoColor pusher = new UndoColor(redoColors, redoCells);
                this.redos.Push(pusher);
            }

            if (this.undos.Count != 0)
            {
                return "Undo " + this.undos.Peek().Title;
            }
            else
            {
                return "Undo";
            }
        }

        /// <summary>
        /// Pops the top of the redo stack and redoes what happened.
        /// </summary>
        /// <returns>String of what the next redo type is.</returns>
        public string Redo()
        {
            Type type = this.redos.Peek().GetType();

            if (type.Equals(typeof(UndoText)))
            {
                UndoText cell = (UndoText)this.redos.Pop();
                char colLetter = (char)(cell.Column + 65);

                this.SetCell(cell.Row, colLetter, cell.Text, false);
            }
            else if (type.Equals(typeof(UndoColor)))
            {
                UndoColor cells = (UndoColor)this.redos.Pop();

                for (int i = 0; i < cells.Count; i++)
                {
                    Tuple<int, int> rowCol = cells.Cells;
                    this.SetColor(cells.Color, rowCol.Item1, rowCol.Item2);
                }
            }

            if (this.redos.Count != 0)
            {
                return "Redo " + this.redos.Peek().Title;
            }
            else
            {
                return "Redo";
            }
        }

        /// <summary>
        /// Checks if the undo stack is empty.
        /// </summary>
        /// <returns>True if stack is empty.</returns>
        public bool IsUndoEmpty()
        {
            if (this.undos.Count == 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Checks if the redo stack is empty.
        /// </summary>
        /// <returns>True if stack is empty.</returns>
        public bool IsRedoEmpty()
        {
            if (this.redos.Count == 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Clears the undo and redo stacks.
        /// </summary>
        public void ClearUndoRedo()
        {
            this.undos.Clear();
            this.redos.Clear();
        }

        /// <summary>
        /// Saves the spreadsheet as an xml file.
        /// </summary>
        /// <param name="file">File to save item to.</param>
        public void SaveXML(FileStream file)
        {
            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Indent = true;
            settings.NewLineOnAttributes = true;
            settings.OmitXmlDeclaration = true;

            using (XmlWriter writer = XmlWriter.Create(file, settings))
            {
                writer.WriteStartElement("Spreadsheet"); // Start spreadsheet
                foreach (Cell cell in this.sheet)
                {
                    if (cell.Changed == true)
                    {
                        char columnChar = Convert.ToChar(cell.GetColumn() + 65);
                        string postion = Convert.ToString(columnChar) + Convert.ToString(cell.GetRow() + 1);

                        writer.WriteStartElement("Cell"); // Start cell
                        writer.WriteAttributeString("Position", postion);

                        writer.WriteStartElement("Text"); // Start text portion
                        writer.WriteString(cell.GetText());
                        writer.WriteEndElement(); // End text portion

                        writer.WriteStartElement("Color"); // Start color portion
                        writer.WriteString(Convert.ToString(cell.Color));
                        writer.WriteEndElement(); // End color portion

                        writer.WriteEndElement(); // End whole cell
                    }
                }

                writer.WriteEndElement(); // End spreadsheet portion
                writer.Close();
            }
        }

        /// <summary>
        /// Loads an xml document file.
        /// </summary>
        /// <param name="xml">File to be loaded.</param>
        public void LoadXML(FileStream xml)
        {
            this.ClearSpreadsheet();
            XmlReaderSettings settings = new XmlReaderSettings();
            settings.DtdProcessing = DtdProcessing.Parse;
            XmlReader reader = XmlReader.Create(xml, settings);

            string text = string.Empty;
            string position = string.Empty;
            uint color = 0;
            this.ClearUndoRedo();

            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element)
                {
                    if (reader.Name == "Cell")
                    {
                        reader.MoveToNextAttribute(); // Moves to next item in xml
                        position = reader.Value;
                    }
                    else if (reader.Name == "Text")
                    {
                        reader.Read();
                        text = reader.Value; // Sets the text
                    }
                    else if (reader.Name == "Color")
                    {
                        reader.Read();
                        color = Convert.ToUInt32(reader.Value); // Sets the color
                    }
                }
                else if (reader.NodeType == XmlNodeType.EndElement)
                {
                    if (reader.Name == "Cell")
                    {
                        char colChar = Convert.ToChar(position.Substring(0, 1));
                        int row = Convert.ToInt32(position.Substring(1)) - 1;

                        if (text != string.Empty)
                        {
                            this.SetCell(row, colChar, text, false); // Sets the cell text to the loaded text if it is not empty
                        }

                        if (color != 0)
                        {
                            int col = Convert.ToInt32(colChar) - 65;
                            this.SetColor(color, row, col); // Sets the cell color to the loaded text if it is not empty
                        }

                        text = string.Empty;
                        position = string.Empty;
                        color = 0;
                    }
                }
            }
        }

        /// <summary>
        /// Clears the spreadsheet.
        /// </summary>
        public void ClearSpreadsheet()
        {
            for (int i = 0; i < this.rows; i++)
            {
                for (int j = 0; j < this.columns; j++)
                {
                    char col = Convert.ToChar(j + 65);
                    this.SetCell(i, col, string.Empty, false);
                    this.SetColor(4294967295, i, j);
                    this.sheet[i, j].Changed = false;
                }
            }
        }

        private void NotifyCellPropertyChanged(object sender, PropertyChangedEventArgs e, bool eval, string word, bool color, int error)
        {
            if (!color && error == 0)
            {
                if (eval)
                {
                    (sender as Cell).Evaluate();
                }
                else if (word == string.Empty)
                {
                    (sender as Cell).SetValue((sender as Cell).GetText());
                }
                else
                {
                    (sender as Cell).SetValue(word);
                }
            }
            else if (error != 0)
            {
                switch (error)
                {
                    case 1:
                        (sender as Cell).SetValue("!(Self Reference)");
                        break;
                    case 2:
                        (sender as Cell).SetValue("!(Bad Reference)");
                        break;
                    case 3:
                        (sender as Cell).SetValue("!(Circular Reference)");
                        break;
                    default:
                        break;
                }
            }

            if (this.PropertyChanged != null)
            {
                this.PropertyChanged(sender as CellHelp, new PropertyChangedEventArgs("CellText"));
            }
        }
    }
}
