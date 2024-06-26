﻿using ClosedXML.Excel.Caching;

namespace ClosedXML.Excel
{
    internal class XLBorderValue
    {
        private static readonly XLBorderRepository Repository = new XLBorderRepository(key => new XLBorderValue(key));

        public static XLBorderValue FromKey(ref XLBorderKey key)
        {
            return Repository.GetOrCreate(ref key);
        }

        private static readonly XLBorderKey DefaultKey = new XLBorderKey
        {
            BottomBorder = XLBorderStyleValues.None,
            DiagonalBorder = XLBorderStyleValues.None,
            DiagonalDown = false,
            DiagonalUp = false,
            LeftBorder = XLBorderStyleValues.None,
            RightBorder = XLBorderStyleValues.None,
            TopBorder = XLBorderStyleValues.None,
            BottomBorderColor = XLColor.Black.Key,
            DiagonalBorderColor = XLColor.Black.Key,
            LeftBorderColor = XLColor.Black.Key,
            RightBorderColor = XLColor.Black.Key,
            TopBorderColor = XLColor.Black.Key
        };

        internal static readonly XLBorderValue Default = FromKey(ref DefaultKey);

        public XLBorderKey Key { get; private set; }

        public XLBorderStyleValues LeftBorder => Key.LeftBorder;

        public XLColor LeftBorderColor { get; private set; }

        public XLBorderStyleValues RightBorder => Key.RightBorder;

        public XLColor RightBorderColor { get; private set; }

        public XLBorderStyleValues TopBorder => Key.TopBorder;

        public XLColor TopBorderColor { get; private set; }

        public XLBorderStyleValues BottomBorder => Key.BottomBorder;

        public XLColor BottomBorderColor { get; private set; }

        public XLBorderStyleValues DiagonalBorder => Key.DiagonalBorder;

        public XLColor DiagonalBorderColor { get; private set; }

        public bool DiagonalUp => Key.DiagonalUp;

        public bool DiagonalDown => Key.DiagonalDown;

        private XLBorderValue(XLBorderKey key)
        {
            Key = key;
            var leftBorderColor = Key.LeftBorderColor;
            var rightBorderColor = Key.RightBorderColor;
            var topBorderColor = Key.TopBorderColor;
            var bottomBorderColor = Key.BottomBorderColor;
            var diagonalBorderColor = Key.DiagonalBorderColor;
            LeftBorderColor = XLColor.FromKey(ref leftBorderColor);
            RightBorderColor = XLColor.FromKey(ref rightBorderColor);
            TopBorderColor = XLColor.FromKey(ref topBorderColor);
            BottomBorderColor = XLColor.FromKey(ref bottomBorderColor);
            DiagonalBorderColor = XLColor.FromKey(ref diagonalBorderColor);
        }

        public override bool Equals(object obj)
        {
            return obj is XLBorderValue cached &&
                   Key.Equals(cached.Key);
        }

        public override int GetHashCode()
        {
            return -280332839 + Key.GetHashCode();
        }
    }
}
