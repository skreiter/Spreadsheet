// <copyright file="Node.cs" company="Sage Kreiter - 11659212">
// Copyright (c) Sage Kreiter. All rights reserved.
// </copyright>

namespace SpreadsheetEngine
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Class for other nodes to inherit from.
    /// </summary>
    public abstract class Node
    {
        /// <summary>
        /// Evaluation overridable function.
        /// </summary>
        /// <param name="dic">Dictionary of variables.</param>
        /// <returns>Value of expression.</returns>
        public abstract double Evaluate(Dictionary<string, double> dic);
    }
}
