using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace GroupSpecVarB
{
    //
    //  Ячейка формы спецификации
    //
    public class SpecFormCell
    {
        string _text;
        UnderliningFormat _underlining;
        AlignFormat _align;

        /**
         *  Параметр подчеркивания
         */
        public enum UnderliningFormat
        {
            NotUnderlined,
            Underlined
        }

        /**
         *  Параметр выравнивания
         */
        public enum AlignFormat
        {
            Left,
            Center,
            Right
        }

        public SpecFormCell(string text, UnderliningFormat underlining, AlignFormat align)
        {
            _text = text;
            _underlining = underlining;
            _align = align;
        }

        public SpecFormCell(string text, AlignFormat align)
            : this(text, UnderliningFormat.NotUnderlined, align)
        {
        }

        public SpecFormCell(string text)
            : this(text, UnderliningFormat.NotUnderlined, AlignFormat.Left)
        {
        }

        public override string ToString()
        {
            return _text;
        }

        public string Text
        {
            get
            {
                return _text;
            }

            set
            {
                _text = value;
            }
        }

        public UnderliningFormat Underlining
        {
            get { return _underlining; }
        }

        public AlignFormat Align
        {
            get { return _align; }
        }

        public static SpecFormCell Empty
        {
            get
            {
                return new SpecFormCell("");
            }
        }
    }
}
