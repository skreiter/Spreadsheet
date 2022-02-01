// <copyright file="ExpressionTreeTests.cs" company="Sage Kreiter - 11659212">
// Copyright (c) Sage Kreiter. All rights reserved.
// </copyright>

namespace NUnitTestSpreadsheet
{
    using System;
    using System.Collections.Generic;
    using System.Text;
    using NUnit.Framework;
    using SpreadsheetEngine;

    /// <summary>
    /// Tests teh expression tree.
    /// </summary>
    public class ExpressionTreeTests
    {
        /// <summary>
        /// Tests the construction of the tree as well as evaluate.
        /// </summary>
        [Test]
        public void TestExpressionTreeConstruction()
        {
            string expression = "A1+2+3";

            ExpressionTree tree = new ExpressionTree(expression);

            Assert.AreEqual(5, tree.Evaluate());
        }

        /// <summary>
        /// Tests the construction of the tree as well as evaluate.
        /// </summary>
        [Test]
        public void TestExpressionTreeConstructionEdgeCase()
        {
            string expression = string.Empty;

            ExpressionTree tree = new ExpressionTree(expression);

            Assert.AreEqual(0, tree.Evaluate());
        }

        /// <summary>
        /// Tests setting a variable in an expression tree.
        /// </summary>
        [Test]
        public void TestVariableSetting()
        {
            ExpressionTree tree = new ExpressionTree("A1 + 2 * 3");
            tree.SetVariable("A1", 5);

            Assert.AreEqual(11, tree.Evaluate());
        }

        /// <summary>
        /// Tests setting a variable in an expression tree.
        /// </summary>
        [Test]
        public void TestVariableSettingWithExtraVar()
        {
            ExpressionTree tree = new ExpressionTree("A1 + B3 + 3");
            tree.SetVariable("A1", 5);
            tree.SetVariable("B3", 10);

            Assert.AreEqual(18, tree.Evaluate());
        }

        /// <summary>
        /// Testing the expression tree when parenthesis are used.
        /// </summary>
        [Test]
        public void TestExpressionWithParanthesis()
        {
            ExpressionTree tree = new ExpressionTree("3*(4+1)");

            Assert.AreEqual(15, tree.Evaluate());
        }

        /// <summary>
        /// Testing the expression tree when multiple parenthesis are used.
        /// </summary>
        [Test]
        public void TestExpressionWithTwoParanthesis()
        {
            ExpressionTree tree = new ExpressionTree("(A1+3)*(B3-4)");
            tree.SetVariable("A1", 5);
            tree.SetVariable("B3", 6);

            Assert.AreEqual(16, tree.Evaluate());
        }
    }
}
