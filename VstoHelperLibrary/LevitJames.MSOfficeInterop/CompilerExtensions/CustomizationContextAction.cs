// © Copyright 2018 Levit & James, Inc.

using Microsoft.Office.Interop.Word;

namespace LevitJames.MSOffice.MSWord
{
    /// <summary>
    ///     A delegate used by the <see cref="Extensions.LJCustomizationContext" /> method.
    /// </summary>
    /// <typeparam name="T">The Type used for the return value from the delegate call.</typeparam>
    /// <param name="document">The Document on which the context will act upon.</param>
    /// <param name="previousContext">The previous customization context of the Document.</param>
    /// <returns>Any value required to be returned to the calling code.</returns>
    public delegate T CustomizationContextAction<out T>(Document document, object previousContext);

    public delegate void CustomizationContextAction(Document document, object previousContext);
}