// © Copyright 2018 Levit & James, Inc.

namespace LevitJames.AddinApplicationFramework
{


    public interface IAddinAppDirty 
    {
        bool Dirty { get; }
        int DirtyCookie { get; set; }

    }
}