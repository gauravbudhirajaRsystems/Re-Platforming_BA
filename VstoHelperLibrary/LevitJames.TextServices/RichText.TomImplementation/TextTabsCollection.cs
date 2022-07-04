using System.Collections;
using System.Collections.Generic;

namespace LevitJames.TextServices
{
    /// <summary>
    /// Represents a collection of TextParagraphTabs.
    /// </summary>
    public class TextTabsCollection : IReadOnlyList<TextTab>
    {
        private readonly ITextPara _para;

        internal TextTabsCollection(ITextPara para)
        {
            _para = para;
        }

        /// <summary>
        ///     Adds a tab at the displacement tabPos, with type tabAlign, and leader style, tabLeader.
        /// </summary>
        /// <param name="position">New tab displacement, in floating-point points.</param>
        /// <param name="alignment">A ParagraphTabAlign options for the tab position.</param>
        /// <param name="leader">
        ///     A ParagraphAddTabLeader enum value. A leader character is the character that is used to fill
        ///     the space taken by a tab character.
        /// </param>
        public void Add(float position, ParagraphTabAlignment alignment, ParagraphAddTabLeader leader)
        {
            _para.AddTab(position, alignment, leader);
        }

        /// <summary>
        ///     Adds a tab at the displacement tabPos, with type tabAlign, and leader style, tabLeader.
        /// </summary>
        public void Add(TextTab tab)
        {
            _para.AddTab(tab.Position, tab.Alignment, tab.Leader);
        }

        /// <summary>
        ///     Deletes a tab at a specified displacement.
        /// </summary>
        /// <param name="tabPos">Displacement, in floating-point points, at which a tab should be deleted.</param>
        public void Delete(float tabPos)
        {
            _para.DeleteTab(tabPos);
        }

        /// <summary>
        ///     Deletes a tab.
        /// </summary>
        public void Delete(TextTab tab)
        {
            _para.DeleteTab(tab.Position);
        }

        /// <summary>
        ///     Clears all tabs, reverting to equally spaced tabs with the default tab spacing.
        /// </summary>
        public void Clear()
        {
            _para.ClearAllTabs();
        }

        /// <inheritdoc />
        public IEnumerator<TextTab> GetEnumerator()
        {
            for (var i = 0; i < _para.TabCount; i++)
            {
                _para.GetTab((ParagraphTabIndex)i, out var position, out var alignment, out var leader);
                yield return new TextTab(position, alignment, leader);
            }
        }

        /// <inheritdoc />
        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        /// <inheritdoc />
        public int Count => _para.TabCount;

        /// <inheritdoc />
        public TextTab this[int index]
        {
            get
            {
                _para.GetTab((ParagraphTabIndex)index, out var position, out var alignment, out var leader);
                return new TextTab(position, alignment, leader);
            }
        }
    }
}