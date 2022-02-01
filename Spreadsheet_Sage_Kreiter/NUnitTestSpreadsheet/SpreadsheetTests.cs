// <copyright file="SpreadsheetTests.cs" company="Sage Kreiter - 11659212">
// Copyright (c) Sage Kreiter. All rights reserved.
// </copyright>

namespace NUnitTestSpreadsheet
{
    using System.Collections.Generic;
    using System.IO;
    using System.Xml;
    using NUnit.Framework;
    using SpreadsheetEngine;

    /// <summary>
    /// Tests for methods.
    /// </summary>
    public class SpreadsheetTests
    {
        private Spreadsheet grid;

        /// <summary>
        /// Setting up spreadsheet for tests.
        /// </summary>
        [SetUp]
        public void Setup()
        {
            this.grid = new Spreadsheet(3, 3);
        }

        /// <summary>
        /// Testing if when a cells text is changed, its value is changed.
        /// </summary>
        [Test]
        public void TestCellTextChange()
        {
            this.grid.SetCell(0, 'A', "Test", false);
            Assert.AreEqual("Test", this.grid.GetValue(0, 'A'));
        }

        /// <summary>
        /// Testing if when a cells text starts with = if the value is changed to another cells value.
        /// </summary>
        [Test]
        public void TestCellTextChangeWithEquals()
        {
            this.grid.SetCell(0, 'A', "Test", false);
            this.grid.SetCell(0, 'b', "=A1", false);
            Assert.AreEqual("Test", this.grid.GetValue(0, 'b'));
        }

        /// <summary>
        /// Testing if an expression in a spreadsheet will be evaulated correctly.
        /// </summary>
        [Test]
        public void TestExpressionHandling()
        {
            this.grid.SetCell(0, 'a', "2", false);
            this.grid.SetCell(1, 'a', "1", false);
            this.grid.SetCell(0, 'b', "=a1+a2", false);
            Assert.AreEqual("3", this.grid.GetValue(0, 'b'));
        }

        /// <summary>
        /// Testing undo functionality.
        /// </summary>
        [Test]
        public void TestUndo()
        {
            this.grid.AddUndoText(this.grid.GetCell(0, 0));
            this.grid.SetCell(0, 'a', "2", false);
            this.grid.Undo();

            Assert.AreEqual(string.Empty, this.grid.GetValue(0, 'a'));
        }

        /// <summary>
        /// Testing redo functionality.
        /// </summary>
        [Test]
        public void TestRedo()
        {
            this.grid.AddUndoText(this.grid.GetCell(0, 0));
            this.grid.SetCell(0, 'a', "2", false);
            this.grid.Undo();
            this.grid.Redo();

            Assert.AreEqual("2", this.grid.GetValue(0, 'a'));
        }

        /// <summary>
        /// Test for saving and loading.
        /// </summary>
        [Test]
        public void TestSaveLoad()
        {
            this.grid.SetCell(0, 'a', "Test", false);
            this.grid.SetColor(4294967168, 0, 1);

            Spreadsheet holder = new Spreadsheet(3, 3);
            holder = this.grid;

            FileStream file = new FileStream("test.xml", FileMode.Create, FileAccess.Write);
            this.grid.SaveXML(file);
            file.Close();

            file = new FileStream("test.xml", FileMode.Open, FileAccess.Read);
            this.grid.ClearSpreadsheet();
            this.grid.LoadXML(file);

            Assert.AreEqual(this.grid, holder);
        }

        /// <summary>
        /// Test for a reference to a cell outside of the sheet.
        /// </summary>
        [Test]
        public void TestBadReference()
        {
            this.grid.SetCell(0, 'a', "=D3", false);

            Assert.AreEqual(this.grid.GetValue(0, 'a'), "!(Bad Reference)");
        }

        /// <summary>
        /// Test for a self reference error.
        /// </summary>
        [Test]
        public void TestSelfReference()
        {
            this.grid.SetCell(0, 'a', "=A1", false);

            Assert.AreEqual(this.grid.GetValue(0, 'a'), "!(Self Reference)");
        }

        /// <summary>
        /// Test for circular reference.
        /// </summary>
        [Test]
        public void TestCircularReference()
        {
            this.grid.SetCell(0, 'B', "=A1", false);
            this.grid.SetCell(0, 'a', "=B1", false);

            Assert.AreEqual(this.grid.GetValue(0, 'a'), "!(Circular Reference)");
        }
    }
}