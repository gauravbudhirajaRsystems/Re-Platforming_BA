// © Copyright 2018 Levit & James, Inc.

using System.Collections.Generic;
using System.Linq;
using LevitJames.AddinApplicationFramework.Properties;
using LevitJames.Core;

namespace LevitJames.AddinApplicationFramework
{
    public sealed class AddinAppSession : AddinAppBase
    {
        private readonly Stack<IAddinAppNamedSession> _namedSessionStack = new Stack<IAddinAppNamedSession>();
        private AddinAppDocument _doc;

        public IAddinAppNamedSession NamedSession => !InNamedSession() ? null : _namedSessionStack.Peek();

        public string SessionName => !InNamedSession() ? string.Empty : _namedSessionStack.Peek().Name;

        public bool InSession { get; internal set; }

        public AddinAppTransactionBase ActiveTransaction { get; internal set; }

        internal IAddinAppNamedSession Current => InNamedSession() ? _namedSessionStack.Peek() : null;

        protected override void Initialize(object data)
        {
            base.Initialize(data);
            _doc = (AddinAppDocument) data;
        }

        internal void PushNamedSession(IAddinAppNamedSession namedSession)
        {
            _namedSessionStack.Push(namedSession);
            UpdateNamedSessionVariable();
        }

        internal bool IsNamedSessionActive(IAddinAppNamedSession namedSession) => namedSession == _namedSessionStack.Peek();

        public bool InNamedSession() => _namedSessionStack.Count != 0;
        public bool InNamedSession(string sessionName) => sessionName != null && InSession && _namedSessionStack.Any(p => p.Name == sessionName);
        public bool InNamedSession(IAddinAppNamedSession session) => InNamedSession(session?.Name);
        internal IAddinAppNamedSession GetNamedSession(string sessionName) => _namedSessionStack.FirstOrDefault(p => p.Name == sessionName);

        internal int NamedSessionCount() => _namedSessionStack.Count;

        private void UpdateNamedSessionVariable()
        {
            _doc.Store.Set(AddinAppDocumentStorage.ActiveNamedSessionVariableName, SessionName);
        }

        internal void CloseAllSessions()
        {
            do
            {
                if (!InNamedSession())
                    break;
                CloseNamedSession();
            } while (true);
        }
 
        public void CloseNamedSession()
        {
            var curSession = Current;
            if (curSession == null)
                return;

            if ((ActiveTransaction == null || !InNamedSession(ActiveTransaction.NamedSession)) && !curSession.Closing)
            {
                curSession.Close();
            }

            if (_namedSessionStack.Count == 0 || Current != curSession)
                return;

            _namedSessionStack.Pop();

            App.ViewService.OnOwnerClosing(curSession);

            UpdateNamedSessionVariable();

            if (_namedSessionStack.Count == 0 && ActiveTransaction == null)
            {
                App.TransactionManager.EndEditSession(_doc);
            }
        }

        internal static void AddCloseSessionUserActions()
        {
            //Special UserActions for closing Named Sessions
            if (UserActionManager.UserActions.Contains(AddinAppUserActionConstants.CloseSessionGroup))
                return;

            var ua = new RibbonUserAction(AddinAppUserActionConstants.CloseSession, visible: false, enabled: false , 
                                          updateDelegate: UpdateCloseSessionUserActionStateCallback);

            UserActionManager.UserActions.Add(ua);
            UserActionManager.BindToUserAction(ua, ua.Id);

            ua = new RibbonUserAction(AddinAppUserActionConstants.CloseSessionGroup, visible: false, enabled: false, 
                                      updateDelegate: UpdateCloseSessionUserActionStateCallback);

            UserActionManager.UserActions.Add(ua);
            UserActionManager.BindToUserAction(ua, ua.Id);
        }

        private static void UpdateCloseSessionUserActionStateCallback(UserAction userAction, string propertyName, object context)
        {
            var app = (IAddinApplicationInternal)context;
            if (app == null)
                return;

            var session = app.ActiveDocument?.Session;

            var visibleAndEnabled = session?.InNamedSession() == true;
 
            if (userAction.Id == AddinAppUserActionConstants.CloseSession)
            {
                var text = app.GetStringResource(nameof(Resources.aafCloseSessionUserActionText));
                text += visibleAndEnabled ? " " + session.Current.Name : null;

                userAction.Enabled = visibleAndEnabled;
                userAction.Visible = visibleAndEnabled;
                userAction.Text = text;
                return;
            }

            if (userAction.Id == AddinAppUserActionConstants.CloseSessionGroup)
            {
                userAction.Enabled = visibleAndEnabled;
                userAction.Visible = visibleAndEnabled;
            }
 
        }
        
    }
}