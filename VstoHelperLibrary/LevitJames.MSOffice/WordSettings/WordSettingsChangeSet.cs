// © Copyright 2018 Levit & James, Inc.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Linq.Expressions;
using Microsoft.Office.Interop.Word;

namespace LevitJames.MSOffice.MSWord
{
    /// <summary>
    ///     Class that tracks and applies requested changes to Word settings.
    ///     <note type="inheritinfo">Inherits from <c>Dictionary(Of String, Object)</c></note>
    /// </summary>
    /// <remarks>
    ///     <para>
    ///         Changes to Word settings are requested through the <c>ApplySetting</c> method.
    ///         The <c>WordSettingsChangeSet</c> tracks all such requests since its creation so
    ///         that original values may be restored using the <c>RestoreSettings</c> method.
    ///     </para>
    ///     <para>Word settings are tracked using dictionary entries.</para>
    ///     <para>
    ///         <c>WordSettingsChangeSet</c> objects can be managed via the <seealso cref="WordSettingsChangeSetController" />
    ///         class. This class maintains a stack of <c>WordSettingsChangeSet</c> objects.
    ///     </para>
    /// </remarks>
    [Serializable]
    public class WordSettingsChangeSet
    {
        [NonSerialized] private readonly WordPropertyHandler _wordProp;

        private Dictionary<string, object> _settings;


        internal WordSettingsChangeSet(Document wordDoc, string name)
        {
            _wordProp = new WordPropertyHandler(wordDoc);
            _settings = new Dictionary<string, object>();
            Name = name;
        }


        // PUBLIC MEMBERS
        public string Name { get; }


        /// <summary>
        ///     Method for requesting changes to Word settings.
        /// </summary>
        /// <param name="propertyPath">Fully-qualified object/property key.</param>
        /// <param name="value">Requested value for Word setting.</param>
        /// <remarks>
        ///     <para>
        ///         When a change to a Word setting is requested with the <c>ApplySetting</c>
        ///         method, the following is performed:
        ///     </para>
        ///     <list type="bullet">
        ///         <item>
        ///             A check is made to see if the requested value is different than
        ///             the current value. If not, <c>ApplySetting</c> exits without changing the value.
        ///         </item>
        ///         <item>
        ///             A check is made to see if the requested setting has already been
        ///             entered in the dictionary. If not, then an entry is created with the
        ///             objPropKey key value and the original setting value is stored.
        ///         </item>
        ///         <item>
        ///             The requested setting change is made through a <seealso cref="WordPropertyHandler" />
        ///             object.
        ///         </item>
        ///     </list>
        ///     <para></para>
        ///     <para>
        ///         Supported Word objects:
        ///         <list type="bullet">
        ///             <item>Application</item>
        ///             <item>Application.AutoCorrect</item>
        ///             <item>Application.Browser</item>
        ///             <item>Application.Options</item>
        ///             <item>Application.Selection.Find</item>
        ///             <item>Application.Selection.Find.Replacement</item>
        ///             <item>Document</item>
        ///             <item>Document.ActiveWindow</item>
        ///             <item>Document.ActiveWindow.View</item>
        ///             <item>Document.ActiveWindow.View.Zoom</item>
        ///         </list>
        ///     </para>
        ///     <example>
        ///         <c>ApplySetting("Document.ActiveWindow.View.Type", "wdPrintView")</c>
        ///     </example>
        ///     <para></para>
        /// </remarks>
        public void SetValue(string propertyPath, object value)
        {
            AddSetting(propertyPath, value, () => _wordProp.GetValue(propertyPath));
        }

        /// <summary>
        ///     Method for requesting changes a Word Document instance or its child instances.
        /// </summary>
        /// <param name="expression">An expression representing the value to set.</param>
        /// <param name="value">The value of the property to set.</param>
        public void SetValue(Expression<Func<Document, object>> expression, object value)
        {
            var propertyPath = PropertyPathFromExpression(expression);
            if (propertyPath == null)
                throw new ArgumentException(@"Invalid expression passed", nameof(expression));

            AddSetting(propertyPath, value, () => expression.Compile()(_wordProp.Document));
        }

        /// <summary>
        ///     Method for requesting changes to Word settings.
        /// </summary>
        /// <param name="expression">An expression representing the value to set.</param>
        /// <param name="value">The value of the property to set.</param>
        public void SetApplicationValue(Expression<Func<Application, object>> expression, object value)
        {
            var propertyPath = PropertyPathFromExpression(expression);
            if (propertyPath == null)
                throw new ArgumentException(@"Invalid expression passed", nameof(expression));

            AddSetting(propertyPath, value, () => expression.Compile()(_wordProp.Document.Application));
        }


        private void AddSetting(string propertyPath, object newValue, Func<object> getCurValue)
        {
            if (_settings == null)
                _settings = new Dictionary<string, object>();

            // If there isn't an entry in the dictionary, create one and fill it with the current value
            if (!_settings.ContainsKey(propertyPath))
            {
                var curValue = getCurValue();
                // Don't do anything if Word setting property is already set to the requested value
                if (curValue == newValue)
                    return;

                _settings.Add(propertyPath, curValue);
            }


            // Apply the new value
            _wordProp.SetValue(propertyPath, newValue);
        }


        /// <summary>
        ///     Restores all Word setting changes made since the <c>WordSettingsChangeSet</c> object was created.
        /// </summary>
        /// <remarks>
        ///     For each entry in the dictionary, the original value of the setting
        ///     is only restored if it is different from the current setting value.
        /// </remarks>
        internal void Restore()
        {
            Trace.TraceInformation("Restoring change set: " + Name);

            foreach (var key in _settings.Keys.Reverse())
            {
                var val = _settings[key];
                _wordProp.SetValue(key, val);
            }

            _settings.Clear();
        }


        private string PropertyPathFromExpression(LambdaExpression expression)
        {
            MemberExpression expr;
            switch (expression.Body.NodeType)
            {
            case ExpressionType.Convert:
            case ExpressionType.ConvertChecked:
                var ue = expression.Body as UnaryExpression;
                expr = ue?.Operand as MemberExpression;
                break;
            default:
                expr = expression.Body as MemberExpression;
                break;
            }

            var items = new List<string>();
            while (expr != null)
            {
                items.Add(expr.Member.Name);
                expr = expr.Expression as MemberExpression;
            }

            items.Reverse();
            return expression.Parameters[0].Type.Name + "." + string.Join(".", items);
        }
    }
}