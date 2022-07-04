using System.Diagnostics;

namespace LevitJames.TextServices
{
    /// <summary>
    ///     Class to
    /// </summary>
    [DebuggerStepThrough]
    public class TextTab
    {

        internal TextTab(float position, ParagraphTabAlignment alignment, ParagraphAddTabLeader leader)
        {
            Position = position;
            Alignment = alignment;
            Leader = leader;
        }

        /// <summary>
        ///     Tab position, in points.
        /// </summary>
        public float Position { get; }

        /// <summary>
        ///     Tab alignment. Value is a ParagraphTabAlignment enum value.
        /// </summary>
        public ParagraphTabAlignment Alignment { get; }

        /// <summary>
        ///     Tab leader. Value is a ParagraphAddTabLeader enum value.
        /// </summary>
        public ParagraphAddTabLeader Leader { get; }
    }
}