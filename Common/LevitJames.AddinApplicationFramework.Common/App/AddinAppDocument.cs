using LevitJames.AddinApplicationFramework.Common.WordAddinApplication;
using System.Runtime.Serialization;

namespace LevitJames.AddinApplicationFramework.Common.App
{
    /// <summary>
    ///     An abstract class which wraps a Word Document, providing additional functionality.
    /// </summary>
    public abstract class AddinAppDocument : WordAddinDocument, IAddinAppProvider, ISerializable, IAddinAppDirty
    {
        public void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            throw new System.NotImplementedException();
        }
    }
}
