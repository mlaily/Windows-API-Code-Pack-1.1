//Copyright (c) Microsoft Corporation.  All rights reserved.

using System;
using System.Collections.Generic;
using Microsoft.WindowsAPICodePack.Shell;

namespace Microsoft.WindowsAPICodePack.Controls
{

    /// <summary>
    /// Event argument for The NavigationPending event
    /// </summary>
    public class NavigationPendingEventArgs : EventArgs
    {
        /// <summary>
        /// The location being navigated to
        /// </summary>
        public ShellObject PendingLocation { get; set; }

        /// <summary>
        /// Set to 'True' to cancel the navigation.
        /// </summary>
        public bool Cancel { get; set; }

    }

    /// <summary>
    /// Event argument for The NavigationComplete event
    /// </summary>
    public class NavigationCompleteEventArgs : EventArgs
    {
        /// <summary>
        /// The new location of the explorer browser
        /// </summary>
        public ShellObject NewLocation { get; set; }
    }

    /// <summary>
    /// Event argument for the NavigatinoFailed event
    /// </summary>
    public class NavigationFailedEventArgs : EventArgs
    {
        /// <summary>
        /// The location the the browser would have navigated to.
        /// </summary>
        public ShellObject FailedLocation { get; set; }
    }

    /// <summary>
    /// Event argument for the ExecutingDefaultCommand event.
    /// </summary>
    public class ExecutingDefaultCommandEventArgs : EventArgs
    {
        /// <summary>
        /// When set, prevents the default opening action from happening.
        /// </summary>
        public bool Cancel { get; set; }

        /// <summary>
        /// Selected items in the view at the moment of the event.
        /// </summary>
        public IEnumerable<ShellObject> SelectedItems { get; private set; }

        public ExecutingDefaultCommandEventArgs(IEnumerable<ShellObject> selectedItems)
        {
            SelectedItems = new List<ShellObject>(selectedItems);
        }
    }

    /// <summary>
    /// Event argument for the IncludingObject event.
    /// </summary>
    public class IncludingObjectEventArgs : EventArgs
    {
        /// <summary>
        /// When set, prevents the object from being included in the view.
        /// </summary>
        public bool Cancel { get; set; }

        /// <summary>
        /// Object being included in the view.
        /// </summary>
        public ShellObject ShellObject { get; private set; }

        public IncludingObjectEventArgs(ShellObject shellObject)
        {
            ShellObject = shellObject;
        }
    }
}