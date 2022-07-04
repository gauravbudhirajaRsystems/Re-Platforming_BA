using System;
using System.ComponentModel;
using System.ComponentModel.Design.Serialization;
using System.Globalization;

namespace LevitJames.Win32
{
    internal class WindowPlacementConverter : TypeConverter
    {
        public override bool CanConvertFrom(ITypeDescriptorContext context, Type sourceType) =>
                ((sourceType == typeof(string)) || base.CanConvertFrom(context, sourceType));

        public override bool CanConvertTo(ITypeDescriptorContext context, Type destinationType) =>
            ((destinationType == typeof(InstanceDescriptor)) || base.CanConvertTo(context, destinationType));

        public override object ConvertFrom(ITypeDescriptorContext context, CultureInfo culture, object value)
        {
            var strValue = (value as string)?.Trim();
            if (string.IsNullOrEmpty(strValue))
                return base.ConvertFrom(context, culture, value);

            if (culture == null)
                culture = CultureInfo.CurrentCulture;

            var ch = culture.TextInfo.ListSeparator[0];
            var separator = new char[] { ch };
            var strArray = strValue.Split(separator);
            var numArray = new int[strArray.Length];
            TypeConverter converter = TypeDescriptor.GetConverter(typeof(int));

            for (int i = 0; i < numArray.Length; i++)
            {
                numArray[i] = (int)converter.ConvertFromString(context, culture, strArray[i]);
            }
            if (numArray.Length == 10)
            {
                return new WindowPlacement()
                {
                    Flags = numArray[0],
                    ShowCmd = numArray[1],
                    MinPosition = new PointI(numArray[2], numArray[3]),
                    MaxPosition = new PointI(numArray[4], numArray[5]),
                    NormalPosition = new RectangleI(numArray[6], numArray[7], numArray[8], numArray[9])
                };


            }

            throw new ArgumentException(nameof(WindowPlacementConverter) + " Failed", strValue);
        }
 
        public override object ConvertTo(ITypeDescriptorContext context, CultureInfo culture, object value, Type destinationType)
        {
                
            if (destinationType == null)
                throw new ArgumentNullException(nameof(destinationType));

            if (value is WindowPlacement && destinationType == typeof(string))
            {
                WindowPlacement wp = (WindowPlacement)value;
                if (culture == null)
                    culture = CultureInfo.CurrentCulture;

                string separator = culture.TextInfo.ListSeparator + " ";
                TypeConverter converter = TypeDescriptor.GetConverter(typeof(int));
                var strArray = new string[10];
                strArray[0] = converter.ConvertToString(context, culture, wp.Flags);
                strArray[1] = converter.ConvertToString(context, culture, wp.ShowCmd);
                strArray[2] = converter.ConvertToString(context, culture, wp.MinPosition.X);
                strArray[3] = converter.ConvertToString(context, culture, wp.MinPosition.Y);
                strArray[4] = converter.ConvertToString(context, culture, wp.MaxPosition.X);
                strArray[5] = converter.ConvertToString(context, culture, wp.MaxPosition.Y);
                strArray[6] = converter.ConvertToString(context, culture, wp.NormalPosition.X);
                strArray[7] = converter.ConvertToString(context, culture, wp.NormalPosition.Y);
                strArray[8] = converter.ConvertToString(context, culture, wp.NormalPosition.Right);
                strArray[9] = converter.ConvertToString(context, culture, wp.NormalPosition.Bottom);

                return string.Join(separator, strArray);
            }

            return base.ConvertTo(context, culture, value, destinationType);
        }

    }
}