// <copyright file="Node.cs" company="Sage Kreiter - 11659212">
// Copyright (c) Sage Kreiter. All rights reserved.
// </copyright>

namespace SpreadsheetEngine
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Leaf Node to hold constant values.
    /// </summary>
    public class ConstantNode : Node
    {
        /// <summary>
        /// Gets or Sets Value of the constant.
        /// </summary>
        public double Value { get; set; }

        /// <summary>
        /// Helps to evalueate the expression tree by returning nodes value.
        /// </summary>
        /// <param name="dic">Dictionary of variables.</param>
        /// <returns>Value of node.</returns>
        public override double Evaluate(Dictionary<string, double> dic)
        {
            return this.Value;
        }
    }
}
