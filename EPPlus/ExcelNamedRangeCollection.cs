/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See https://github.com/JanKallman/EPPlus for details.
 *
 * Copyright (C) 2011  Jan Källman
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "as is" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
 * Author							Change						Date
 * ******************************************************************************
 * Jan Källman		Added this class		        2010-01-28
 * Jan Källman		License changed GPL-->LGPL 2011-12-27
 *******************************************************************************/
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;

namespace OfficeOpenXml
{
    /// <summary>
    /// Collection for named ranges
    /// </summary>
    public class ExcelNamedRangeCollection : IEnumerable<ExcelNamedRange>
    {
        private readonly ExcelWorksheet _ws;
        private readonly ExcelWorkbook _wb;
        public ExcelNamedRangeCollection(ExcelWorkbook wb)
        {
            _wb = wb;
            _ws = null;
        }
        internal ExcelNamedRangeCollection(ExcelWorkbook wb, ExcelWorksheet ws)
        {
            _wb = wb;
            _ws = ws;
        }

        private readonly List<ExcelNamedRange> _list = new();
        private readonly Dictionary<string, int> _dic = new(StringComparer.OrdinalIgnoreCase);
        /// <summary>
        /// Add a new named range
        /// </summary>
        /// <param name="name">The name</param>
        /// <param name="range">The range</param>
        /// <returns></returns>
        public ExcelNamedRange Add(string name, ExcelRangeBase range)
        {
            if (!ExcelAddressUtil.IsValidName(name))
            {
                throw new ArgumentException($"Name {name} contains invalid characters");  //Issue 458
            }
            var item = range.IsName 
                ? new ExcelNamedRange(name, _wb, _ws, _dic.Count) 
                : new ExcelNamedRange(name, _ws, range.Worksheet, range.Address, _dic.Count);

            AddName(name, item);

            return item;
        }

        private void AddName(string name, ExcelNamedRange item)
        {
            _dic.Add(name, _list.Count);
            _list.Add(item);
        }
        /// <summary>
        /// Add a defined name referencing value
        /// </summary>
        /// <param name="name"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public ExcelNamedRange AddValue(string name, object value)
        {
            var item = new ExcelNamedRange(name, _wb, _ws, _dic.Count);
            item.NameValue = value;
            AddName(name, item);
            return item;
        }

        /// <summary>
        /// Add a defined name referencing a formula
        /// </summary>
        /// <param name="name"></param>
        /// <param name="formula"></param>
        /// <returns></returns>
        public ExcelNamedRange AddFormula(string name, string formula)
        {
            var item = new ExcelNamedRange(name, _wb, _ws, _dic.Count);
            item.NameFormula = formula;
            AddName(name, item);
            return item;
        }

        internal void Insert(int rowFrom, int colFrom, int rows, int cols)
        {
            Insert(rowFrom, colFrom, rows, cols, n => true);
        }

        internal void Insert(int rowFrom, int colFrom, int rows, int cols, Func<ExcelNamedRange, bool> filter)
        {
            var namedRanges = _list.Where(filter);
            foreach (var namedRange in namedRanges)
            {
                InsertRows(rowFrom, rows, namedRange);
                InsertColumns(colFrom, cols, namedRange);
            }
        }
        internal void Delete(int rowFrom, int colFrom, int rows, int cols)
        {
            Delete(rowFrom, colFrom, rows, cols, n => true);
        }
        internal void Delete(int rowFrom, int colFrom, int rows, int cols, Func<ExcelNamedRange, bool> filter)
        {
            var namedRanges = _list.Where(filter);
            foreach (var namedRange in namedRanges)
            {
                ExcelAddressBase adr;
                if (cols > 0 && rowFrom == 0 && rows >= ExcelPackage.MaxRows)   //Issue 15554. Check
                {
                    adr = namedRange.DeleteColumn(colFrom, cols);
                }
                else
                {
                    adr = namedRange.DeleteRow(rowFrom, rows);
                }
                namedRange.Address = adr == null ? "#REF!" : adr.Address;
            }
        }
        private void InsertColumns(int colFrom, int cols, ExcelNamedRange namedRange)
        {
            if (colFrom > 0)
            {
                if (colFrom <= namedRange.Start.Column)
                {
                    var newAddress = ExcelCellBase.GetAddress(namedRange.Start.Row, namedRange.Start.Column +cols, namedRange.End.Row, namedRange.End.Column + cols);
                    namedRange.Address = BuildNewAddress(namedRange, newAddress);
                }
                else if (colFrom <= namedRange.End.Column && namedRange.End.Column + cols < ExcelPackage.MaxColumns)
                {
                    var newAddress = ExcelCellBase.GetAddress(namedRange.Start.Row, namedRange.Start.Column, namedRange.End.Row, namedRange.End.Column + cols);
                    namedRange.Address = BuildNewAddress(namedRange, newAddress);
                }
            }
        }

        private static string BuildNewAddress(ExcelNamedRange namedRange, string newAddress)
        {
            if (!namedRange.FullAddress.Contains('!')) return newAddress;
            var worksheet = namedRange.FullAddress.Split('!')[0];
            worksheet = worksheet.Trim('\'');
            newAddress = ExcelCellBase.GetFullAddress(worksheet, newAddress);
            return newAddress;
        }

        private void InsertRows(int rowFrom, int rows, ExcelNamedRange namedRange)
        {
            if (rows > 0)
            {
                if (rowFrom <= namedRange.Start.Row)
                {
                    var newAddress = ExcelCellBase.GetAddress(namedRange.Start.Row + rows, namedRange.Start.Column, namedRange.End.Row + rows, namedRange.End.Column);
                    namedRange.Address = BuildNewAddress(namedRange, newAddress);
                }
                else if (rowFrom <= namedRange.End.Row && namedRange.End.Row+rows <= ExcelPackage.MaxRows)
                {
                    var newAddress = ExcelCellBase.GetAddress(namedRange.Start.Row, namedRange.Start.Column, namedRange.End.Row + rows, namedRange.End.Column);
                    namedRange.Address = BuildNewAddress(namedRange, newAddress);
                }
            }
        }

        /// <summary>
        /// Remove a defined name from the collection
        /// </summary>
        /// <param name="name">The name</param>
        public void Remove(string name)
        {
            if (_dic.ContainsKey(name))
            {
                var ix = _dic[name];

                for (var i = ix+1; i < _list.Count; i++)
                {
                    _dic.Remove(_list[i].Name);
                    _list[i].Index--;
                    _dic.Add(_list[i].Name, _list[i].Index);
                }
                _dic.Remove(name);
                _list.RemoveAt(ix);
            }
        }
        /// <summary>
        /// Checks collection for the presence of a key
        /// </summary>
        /// <param name="key">key to search for</param>
        /// <returns>true if the key is in the collection</returns>
        public bool ContainsKey(string key)
        {
            return _dic.ContainsKey(key);
        }
        /// <summary>
        /// The current number of items in the collection
        /// </summary>
        public int Count => _dic.Count;

        /// <summary>
        /// Name indexer
        /// </summary>
        /// <param name="name">The name (key) for a Named range</param>
        /// <returns>a reference to the range</returns>
        /// <remarks>
        /// Throws a KeyNotFoundException if the key is not in the collection.
        /// </remarks>
        public ExcelNamedRange this[string name] => _list[_dic[name]];

        public ExcelNamedRange this[int index] => _list[index];

        #region "IEnumerable"
        #region IEnumerable<ExcelNamedRange> Members
        /// <summary>
        /// Implement interface method IEnumerator&lt;ExcelNamedRange&gt; GetEnumerator()
        /// </summary>
        /// <returns></returns>
        public IEnumerator<ExcelNamedRange> GetEnumerator()
        {
            return _list.GetEnumerator();
        }
        #endregion
        #region IEnumerable Members
        /// <summary>
        /// Implement interface method IEnumeratable GetEnumerator()
        /// </summary>
        /// <returns></returns>
        IEnumerator IEnumerable.GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        #endregion
        #endregion

        internal void Clear()
        {
            while (Count>0)
            {
                Remove(_list[0].Name);
            }
        }
    }
}
