// © Copyright 2018 Levit & James, Inc.

namespace LevitJames.Core
{
    /// <summary>
    ///     A delegate used by the UserAction class to facilitate the updating of UserAction states.
    /// </summary>
    /// <param name="userAction">A UserAction instance on which the provided propertyName was changed</param>
    /// <param name="propertyName">The name of the property that changed</param>
    /// <param name="context">The context provided by the UserActionManagers ActiveTarget.Context property.</param>
    public delegate void UserActionUpdateDelegate(UserAction userAction, string propertyName, object context);
}