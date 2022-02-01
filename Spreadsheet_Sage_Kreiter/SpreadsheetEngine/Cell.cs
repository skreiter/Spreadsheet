// <copyright file="Cell.cs" company="Sage Kreiter - 11659212">
// Copyright (c) Sage Krieter. All rights reserved.
// </copyright>

namespace SpreadsheetEngine
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Runtime.CompilerServices;

    /// <summary>
    /// Cell class that keeps track of a single cell.
    /// </summary>
    public abstract class Cell : INotifyPropertyChanged
    {
        /// <summary>
        /// Values that refer to this cell.
        /// </summary>
        public List<Cell> Refer;

        /// <summary>
        /// Values that this cell referes to.
        /// </summary>
        public List<Cell> ReferTo;

        /// <summary>
        /// Actual value of the cell.
        /// </summary>
        protected string value;

        /// <summary>
        /// Text entered into the cell.
        /// </summary>
        protected string text;

        private int rowIndex;
        private int columnIndex;
        private ExpressionTree expressionTree;
        private uint BGColor;
        private bool changed;

        /// <summary>
        /// Initializes a new instance of the <see cref="Cell"/> class.
        /// </summary>
        /// <param name="row">Value to set the row.</param>
        /// <param name="column">Value to set the column.</param>
        public Cell(int row, int column)
        {
            this.rowIndex = row;
            this.columnIndex = column;
            this.expressionTree = new ExpressionTree(string.Empty);
            this.Refer = new List<Cell>();
            this.ReferTo = new List<Cell>();
            this.BGColor = 0xFFFFFFFF;
            this.changed = false;
        }

        /// <summary>
        /// Event handler for when value is changed.
        /// </summary>
        public event PropertyChangedEventHandler PropertyChanged;

        /// <summary>
        /// Gets or sets the color.
        /// </summary>
        public uint Color
        {
            get { return this.BGColor; }

            set { this.BGColor = value; }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the cell has been changed.
        /// </summary>
        public bool Changed
        {
            get { return this.changed; }
            set { this.changed = value; }
        }

        /// <summary>
        /// Sets the expression tree with a string.
        /// </summary>
        public string ExpressionTree
        {
            set
            {
                this.expressionTree = new ExpressionTree(value);
            }
        }

        /// <summary>
        /// Gets the expressionTree.
        /// </summary>
        public ExpressionTree GetTree
        {
            get { return this.expressionTree; }
        }

        /// <summary>
        /// Sets a variable in the tree.
        /// </summary>
        /// <param name="key">Key to be set.</param>
        /// <param name="value">Value to be set for the key.</param>
        public void SetVariables(string key, double value)
        {
            this.expressionTree.SetVariable(key, value);
        }

        /// <summary>
        /// Gets the row.
        /// </summary>
        /// <returns>The row.</returns>
        public int GetRow()
        {
            return this.rowIndex;
        }

        /// <summary>
        /// Gets the column.
        /// </summary>
        /// <returns>The column.</returns>
        public int GetColumn()
        {
            return this.columnIndex;
        }

        /// <summary>
        /// Evaluates the expression and gives a value.
        /// </summary>
        public void Evaluate()
        {
            this.value = this.expressionTree.Evaluate().ToString();
        }

        /// <summary>
        /// Gets the text.
        /// </summary>
        /// <returns>The text.</returns>
        public string GetText()
        {
            return this.text;
        }

        /// <summary>
        /// Checks if reference list contains a specified cell.
        /// </summary>
        /// <param name="row">Row to check.</param>
        /// <param name="column">Column to check.</param>
        /// <returns>True if value is in list false otherwise.</returns>
        public bool ContainReference(int row, int column)
        {
            foreach (Cell reference in this.Refer)
            {
                if (reference.rowIndex == row && reference.columnIndex == column)
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Sets the text value of the cell.
        /// </summary>
        /// <param name="newText">Value to be set.</param>
        public void SetText(string newText)
        {
            if (newText != this.text)
            {
                this.text = newText;
                this.OnPropertyChanged();
            }
        }

        /// <summary>
        /// Gets the value portion of the cell.
        /// </summary>
        /// <returns>The value.</returns>
        public string GetValue()
        {
            return this.value;
        }

        /// <summary>
        /// Sets the value.
        /// </summary>
        /// <param name="newVal">Value to be set.</param>
        public void SetValue(string newVal)
        {
            this.value = newVal;
        }

        /// <summary>
        /// Event handler that changes the value.
        /// </summary>
        /// <param name="sender">String to set value to.</param>
        /// <param name="e">Event sender.</param>
        public void ChangeValue(object sender, PropertyChangedEventArgs e)
        {
            this.value = sender as string;
        }

        /// <summary>
        /// Event for when value is changed.
        /// </summary>
        /// <param name="name">Name of value changing.</param>
        protected void OnPropertyChanged([CallerMemberName] string name = null)
        {
            this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }

    /// <summary>
    /// Helps access the cell.
    /// </summary>
    public class CellHelp : Cell
    {
        private string helpValue;

        /// <summary>
        /// Initializes a new instance of the <see cref="CellHelp"/> class.
        /// </summary>
        /// <param name="newRow">Row to be set.</param>
        /// <param name="newColumn">Column to be set.</param>
        public CellHelp(int newRow, int newColumn)
            : base(newRow, newColumn)
        {
        }

        /// <summary>
        /// Sets the value.
        /// </summary>
        /// <param name="newVal">Value to be set.</param>
        public void SetHelpValue(string newVal)
        {
            this.helpValue = newVal;
        }
    }
}
