// © Copyright 2018 Levit & James, Inc.

namespace LevitJames.MSOffice
{
    /// <summary>
    ///     A class that exposes the Word Events collections that can be suppresses. Use via the WordExtensions.SuppressEvents
    ///     property.
    /// </summary>
    public sealed class SuppressWordEvents
    {
        private SuppressWordApplicationEvents _wordAppEvents;

        internal SuppressWordEvents() { }

        /// <summary>
        ///     Returns an SuppressWordApplicationEvents instance used for suppressing Application Events.
        /// </summary>
        public SuppressWordApplicationEvents ApplicationEvents =>
            _wordAppEvents ?? (_wordAppEvents = new SuppressWordApplicationEvents());
    }
}