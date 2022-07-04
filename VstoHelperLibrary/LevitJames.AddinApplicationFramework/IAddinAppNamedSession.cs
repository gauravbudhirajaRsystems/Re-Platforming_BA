// © Copyright 2018 Levit & James, Inc.


namespace LevitJames.AddinApplicationFramework
{
    public interface IAddinAppNamedSession
    {
        AddinAppDocument Document { get; }

        string Name { get; }

        bool Closing { get; }

        void Close();
    }
}