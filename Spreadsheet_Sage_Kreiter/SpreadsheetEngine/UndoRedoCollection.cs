// <copyright file="UndoRedoCollection.cs" company="Sage Kreiter - 11659212">
// Copyright (c) Sage Kreiter. All rights reserved.
// </copyright>

namespace SpreadsheetEngine
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Base class for undoing and redoing.
    /// </summary>
    public class UndoRedoCollection
    {
        private string title;

        /// <summary>
        /// Gets or sets the title of the undo or redo.
        /// </summary>
        public string Title
        {
            get { return this.title; }
            set { this.title = value; }
        }
    }
}
