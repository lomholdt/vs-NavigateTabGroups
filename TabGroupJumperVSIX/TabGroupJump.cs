//------------------------------------------------------------------------------
// <copyright file="TabGroupJump.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using System.Globalization;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using EnvDTE80;
using EnvDTE;
using System.Collections.Generic;
using System.Linq;
using Windows = EnvDTE.Windows;

namespace TabGroupJumperVSIX
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class TabGroupJump
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandIdJumpLeft = 0x0100;
        public const int CommandIdJumpRight = 0x0101;
        public const int CommandIdJumpUp = 0x0102;
        public const int CommandIdJumpDown = 0x0103;
        public const int CommandIdJumpPrevious = 0x0104;
        public const int CommandIdJumpNext = 0x0105;

        public enum Mode
        {
            Horizontal,
            Vertical,
            Navigation
        }

    /// <summary>
    /// Command menu group (command set GUID).
    /// </summary>
    public static readonly Guid CommandSet = new Guid("67e760eb-74c3-46f2-9b85-c7af2d351428");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        /// <summary>
        /// Initializes a new instance of the <see cref="TabGroupJump"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private TabGroupJump(Package package)
        {
            if (package == null)
            {
                throw new ArgumentNullException("package");
            }

            this.package = package;

            OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                // Add the jump commands to the handler
                commandService.AddCommand(
                    new MenuCommand(MenuItemCallback_Horizontal, new CommandID(CommandSet, CommandIdJumpLeft)));
                commandService.AddCommand(
                    new MenuCommand(MenuItemCallback_Horizontal, new CommandID(CommandSet, CommandIdJumpRight)));
                commandService.AddCommand(
                    new MenuCommand(MenuItemCallback_Vertical, new CommandID(CommandSet, CommandIdJumpUp)));
                commandService.AddCommand(
                    new MenuCommand(MenuItemCallback_Vertical, new CommandID(CommandSet, CommandIdJumpDown)));
                commandService.AddCommand(
                    new MenuCommand(MenuItemCallback_PreviousNext, new CommandID(CommandSet, CommandIdJumpPrevious)));
                commandService.AddCommand(
                    new MenuCommand(MenuItemCallback_PreviousNext, new CommandID(CommandSet, CommandIdJumpNext)));
            }
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static TabGroupJump Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static void Initialize(Package package)
        {
            Instance = new TabGroupJump(package);
        }

        private IEnumerable<Window> GetValidDocuments(EnvDTE.Windows dteWindows)
        {
            return dteWindows.Cast<Window>()
                             .Where(w => w.Kind == "Document")
                             .Where(w => w.Top != 0 || w.Left != 0)
                             .OrderBy(w => w.Left)
                             .ThenBy(w => w.Top);
        }

        private int FindIndex(List<Window> windows, Document document)
        {
            var activeDoc = document;
            int activeIdx = 0;
            for (int i = 0; i < windows.Count; ++i)
            {
                if (windows[i].Document == activeDoc)
                {
                    activeIdx = i;
                    break;
                }
            }

            return activeIdx;
        }

        private void MenuItemCallback_Horizontal(object sender, EventArgs e)
        {
            DTE2 dte = (DTE2)this.ServiceProvider.GetService(typeof(DTE));
            int commandId = ((MenuCommand)sender).CommandID.ID;

            // Documents with a "left" or "top" = 0 are the non-focused ones in each group, so don't
            // collect those. 
            var topLevel = GetValidDocuments(dte.Windows)
                .OrderBy(w => w.Left)
                .ThenBy(w => w.Top)
                .Where(w => w.Document == dte.ActiveDocument 
                       || w.Left != dte.ActiveDocument.ActiveWindow.Left)
                .ToList();

            if (topLevel.Count == 0)
                return;

            // find the index of the active document
            var activeIdx = FindIndex(topLevel, dte.ActiveDocument);

            // set the new active document
            activeIdx += commandId == CommandIdJumpLeft ? -1 : 1;

            activeIdx = Clamp(activeIdx, topLevel.Count);
            topLevel[activeIdx].Activate();
        }

        private void MenuItemCallback_Vertical(object sender, EventArgs e)
        {
            DTE2 dte = (DTE2)this.ServiceProvider.GetService(typeof(DTE));
            int commandId = ((MenuCommand)sender).CommandID.ID;

            var topLevel = GetValidDocuments(dte.Windows)
                .Where(w => w.Document == dte.ActiveDocument
                            || w.Left == dte.ActiveDocument.ActiveWindow.Left)
                .OrderBy(w => w.Top)
                .ThenBy(w => w.Left)
                .ToList();

            if (topLevel.Count == 0)
                return;

            // find the index of the active document
            var activeIdx = FindIndex(topLevel, dte.ActiveDocument);

            // set the new active document
            activeIdx += commandId == CommandIdJumpDown ? -1 : 1;

            activeIdx = Clamp(activeIdx, topLevel.Count);
            topLevel[activeIdx].Activate();
        }

        private static int Clamp(int activeIdx, int count)
        {
            return (activeIdx < 0 ? activeIdx + count : activeIdx) % count;
        }

        private void MenuItemCallback_PreviousNext(object sender, EventArgs e)
        {
            DTE2 dte = (DTE2)this.ServiceProvider.GetService(typeof(DTE));
            int commandId = ((MenuCommand)sender).CommandID.ID;

            // Documents with a "left" or "top" = 0 are the non-focused ones in each group, so don't
            // collect those. 
            var topLevel = GetValidDocuments(dte.Windows)
                .OrderBy(w => w.Left)
                .ThenBy(w => w.Top)
                .ToList();

            if (topLevel.Count == 0)
                return;

            // find the index of the active document
            var activeIdx = FindIndex(topLevel, dte.ActiveDocument);

            // set the new active document
            activeIdx += commandId == CommandIdJumpPrevious ? -1 : 1;

            activeIdx = Clamp(activeIdx, topLevel.Count);
            topLevel[activeIdx].Activate();
        }
    }
}
