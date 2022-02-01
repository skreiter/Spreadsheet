// <copyright file="UndoText.cs" company="Sage Kreiter - 11659212">
// Copyright (c) Sage Kreiter. All rights reserved.
// </copyright>

namespace SpreadsheetEngine
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Class for undoing text changes.
    /// </summary>
    public class UndoText : UndoRedoCollection
    {
        private string originalText;
        private int row;
        private int column;

        /// <summary>
        /// Initializes a new instance of the <see cref="UndoText"/> class.
        /// </summary>
        /// <param name="newTitle">Title of change.</param>
        /// <param name="newText">Text that the cell originally contained.</param>
        /// <param name="newRow">Row of cell that was changed.</param>
        /// <param name="newColumn">Column of cell that was changed.</param>
        public UndoText(string newTitle, string newText, int newRow, int newColumn)
        {
            this.Title = newTitle;
            this.originalText = newText;
            this.row = newRow;
            this.column = newColumn;
        }

        /// <summary>
        /// Gets the text of the original cell.
        /// </summary>
        public string Text
        {
            get { return this.originalText; }
        }

        /// <summary>
        /// Gets the row of the original cell.
        /// </summary>
        public int Row
        {
            get { return this.row; }
        }

        /// <summary>
        /// Gets the column of the original cell.
        /// </summary>
        public int Column
        {
            get { return this.column; }
        }
    }
}
