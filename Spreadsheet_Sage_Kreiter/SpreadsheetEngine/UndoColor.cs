// <copyright file="UndoColor.cs" company="Sage Kreiter - 11659212">
// Copyright (c) Sage Kreiter. All rights reserved.
// </copyright>

namespace SpreadsheetEngine
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Class to handle Undoing color changes.
    /// </summary>
    public class UndoColor : UndoRedoCollection
    {
        private Stack<Tuple<int, int>> cells; // A stack of pairs where each pair has a row and column that corresponds to a cell that needs to be undone
        private Stack<uint> colors; // Stack of all colors which will be in the same order as the cells stack
        private int numItems;

        /// <summary>
        /// Initializes a new instance of the <see cref="UndoColor"/> class.
        /// </summary>
        /// <param name="newColor">Stack of colors that were changed.</param>
        /// <param name="newCells">Stack of cell rows and columns that were changed.</param>
        public UndoColor(Stack<uint> newColor, Stack<Tuple<int, int>> newCells)
        {
            this.colors = new Stack<uint>();
            this.colors = newColor;
            this.cells = new Stack<Tuple<int, int>>();
            this.cells = newCells;
            this.Title = "Color Change";
            this.numItems = this.colors.Count;
        }

        /// <summary>
        /// Gets a cell row and column from the top of the stack.
        /// </summary>
        public Tuple<int, int> Cells
        {
            get { return this.cells.Pop(); }
        }

        /// <summary>
        /// Gets the color off the top of the stack.
        /// </summary>
        public uint Color
        {
            get { return this.colors.Pop(); }
        }

        /// <summary>
        /// Gets number of items in the stacks.
        /// </summary>
        public int Count
        {
            get { return this.numItems; }
        }
    }
}
