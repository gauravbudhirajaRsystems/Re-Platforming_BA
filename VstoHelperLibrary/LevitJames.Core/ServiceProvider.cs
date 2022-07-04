using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace LevitJames.Core
{
    /// <summary>
    ///     Defines a mechanism for retrieving a service object; that is, an object that provides custom support to other
    ///     objects.
    /// </summary>
    public class ServiceProvider : IServiceProvider
    {
        private readonly object _serviceLock;
        private readonly Dictionary<Type, object> _services;

        /// <summary>
        ///     Creates a new instance of ServiceProvider.
        /// </summary>
        public ServiceProvider()
        {
            _services = new Dictionary<Type, object>();
            _serviceLock = new object();
        }

        /// <summary>
        ///     Gets the service instance of the type supplied, or returns null if the service does not exist.
        /// </summary>
        /// <param name="serviceType">The type of service to retrieve.</param>
        public virtual object GetService(Type serviceType)
        {
            object service;
            // ReSharper disable once InconsistentlySynchronizedField
            _services.TryGetValue(serviceType, out service);
            return service;
        }

        /// <summary>
        ///     Adds the provided service instance to the Provider. If a service of the same type already exists, it is replaced.
        /// </summary>
        /// <typeparam name="TServiceContract">The type of service to Add</typeparam>
        /// <param name="implementation">The service instance to add.</param>
        public void AddService<TServiceContract>(TServiceContract implementation) where TServiceContract : class
        {
            lock (_serviceLock)
            {
                if (implementation == null)
                {
                    _services.Remove(typeof(TServiceContract));
                }
                else
                {
                    _services[typeof(TServiceContract)] = implementation;
                }
            }
        }

        /// <summary>
        ///     Gets the service instance of the generic type supplied, or returns null if the service does not exist.
        /// </summary>
        /// <typeparam name="TServiceContract"></typeparam>
        [DebuggerStepThrough]
        public TServiceContract GetService<TServiceContract>() where TServiceContract : class
        {
            return GetService(typeof(TServiceContract)) as TServiceContract;
        }
    }
}
