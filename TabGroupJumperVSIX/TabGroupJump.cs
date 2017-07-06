//------------------------------------------------------------------------------
// <copyright file="TabGroupJump.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Linq;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Windows = EnvDTE.Windows;

namespace TabGroupJumperVSIX
{
  /// <summary>
  /// Command handler
  /// </summary>
  internal sealed class TabGroupJump
  {
    public const int CommandIdJumpLeft = 0x0100;
    public const int CommandIdJumpRight = 0x0101;
    public const int CommandIdJumpUp = 0x0102;
    public const int CommandIdJumpDown = 0x0103;
    public const int CommandIdJumpPrevious = 0x0104;
    public const int CommandIdJumpNext = 0x0105;

    /// <summary>
    /// Command menu group (command set GUID).
    /// </summary>
    public static readonly Guid CommandSet = new Guid("67e760eb-74c3-46f2-9b85-c7af2d351428");

    /// <summary>
    /// VS Package that provides this command, not null.
    /// </summary>
    private readonly Package package;

    private readonly TabGroupMoverUpDown _tabGroupMoverUpDown;
    private readonly TabGroupMoverLeftRight _tabGroupMoverLeftRight;
    private readonly TabGroupMoverNextPrevious _tabGroupMoverNextPrevious;

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

      _tabGroupMoverLeftRight = new TabGroupMoverLeftRight();
      _tabGroupMoverUpDown = new TabGroupMoverUpDown();
      _tabGroupMoverNextPrevious = new TabGroupMoverNextPrevious();

      _tabGroupMoverLeftRight.Initialize(ServiceProvider);
      _tabGroupMoverUpDown.Initialize(ServiceProvider);
      _tabGroupMoverNextPrevious.Initialize(ServiceProvider);

      OleMenuCommandService commandService =
          ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
      if (commandService != null)
      {
        // Add the jump commands to the handler
        commandService.AddCommand(
            new MenuCommand(_tabGroupMoverLeftRight.MenuItemCallback,
                            new CommandID(CommandSet, CommandIdJumpLeft)));
        commandService.AddCommand(
            new MenuCommand(_tabGroupMoverLeftRight.MenuItemCallback,
                            new CommandID(CommandSet, CommandIdJumpRight)));

        commandService.AddCommand(
            new MenuCommand(_tabGroupMoverUpDown.MenuItemCallback,
                            new CommandID(CommandSet, CommandIdJumpUp)));
        commandService.AddCommand(
            new MenuCommand(_tabGroupMoverUpDown.MenuItemCallback,
                            new CommandID(CommandSet, CommandIdJumpDown)));

        commandService.AddCommand(
            new MenuCommand(_tabGroupMoverNextPrevious.MenuItemCallback,
                            new CommandID(CommandSet, CommandIdJumpPrevious)));
        commandService.AddCommand(
            new MenuCommand(_tabGroupMoverNextPrevious.MenuItemCallback,
                            new CommandID(CommandSet, CommandIdJumpNext)));
      }
    }

    /// <summary>
    /// Gets the instance of the command.
    /// </summary>
    public static TabGroupJump Instance { get; private set; }

    /// <summary>
    /// Gets the service provider from the owner package.
    /// </summary>
    private IServiceProvider ServiceProvider => package;

    /// <summary>
    /// Initializes the singleton instance of the command.
    /// </summary>
    /// <param name="package">Owner package, not null.</param>
    public static void Initialize(Package package)
    {
      Instance = new TabGroupJump(package);
    }

    /// <summary>
    ///  The implementation for moving left/right or up/down or next/previous is basically the same.
    ///  What's different in each case is which command determines if we're moving forward or backward
    ///  and how the tabs should be ordered when determining what is forward or backward. So those two
    ///  functions are abstract methods and the rest of the implementation is in
    ///  <see cref="MenuItemCallback"/>.
    /// </summary>
    private abstract class TabGroupMover
    {
      private IServiceProvider _serviceProvider;

      /// <summary>
      ///  We need a service provider and having an Initialize method makes for smaller derived classes.
      /// </summary>
      public void Initialize(IServiceProvider serviceProvider)
          => _serviceProvider = serviceProvider;

      /// <summary>
      ///  How the document-windows should ordered.  They should be ordered such that when
      ///  <see cref="ShouldMoveForward"/> returns true the Nth + 1 element is the logical group to jump
      ///  to.  For example, if you sort the groups vertically (smallest first), then ShouldMoveFoward()
      ///  should return true when moving down.
      /// </summary>
      /// <seealso cref="ShouldMoveForward"/>
      protected abstract IEnumerable<Window> FilterAndSort(IEnumerable<Window> windows, Document activeDocument);

      /// <summary>
      ///  True if the given command represents moving forward in the collection returned by
      ///  <see cref="FilterAndSort"/>.
      /// </summary>
      /// <seealso cref="FilterAndSort"/>
      protected abstract bool ShouldMoveForward(int commandId);

      /// <summary>
      ///  The callback that should be passed into the <see cref="MenuCommand"/> constructor.
      /// </summary>
      /// <param name="sender"> Source of the event. </param>
      /// <param name="e"> Event information. </param>
      public void MenuItemCallback(object sender, EventArgs e)
      {
        DTE2 dte = (DTE2)_serviceProvider.GetService(typeof(DTE));
        int commandId = ((MenuCommand)sender).CommandID.ID;

        var activeDocument = dte.ActiveDocument;

        var topLevel = FilterAndSort(GetValidDocuments(dte), activeDocument)
            .ToList();

        // for vertical tabs, t.Top might all be the same... not sure why.
        // Maybe this could help to get the actual positions? 
        // https://msdn.microsoft.com/en-us/library/microsoft.visualstudio.shell.interop.ivsuishelldocumentwindowmgr.savedocumentwindowpositions.aspx
        // See
        // https://github.com/eamodio/SaveAllTheTabs/blob/master/src/SaveAllTheTabs/DocumentManager.cs#L248
        // for example usage
        // 
        //var debug = topLevel.Select(w =>
        //                                new
        //                                {
        //                                    w.Left,
        //                                    w.Top,
        //                                    w.Width,
        //                                    w.Height,
        //                                    w.Document.Name,
        //                                }).ToList();

        if (topLevel.Count == 0)
          return;

        var indexOfCurrentTabGroup = GetIndexOfDocument(topLevel, activeDocument);

        // get the tab to activate
        var offset = ShouldMoveForward(commandId) ? 1 : -1;
        int nextIndex = Clamp(topLevel.Count, indexOfCurrentTabGroup + offset);

        // and activate it
        topLevel[nextIndex].Activate();
      }

      private static IEnumerable<Window> GetValidDocuments(DTE2 dte)
      {
        // documents that are not the focused document in a group will have Top == 0 && Left == 0
        return dte.Windows.Cast<Window>()
                  .Where(w => w.Kind == "Document")
                  .Where(w => w.Top != 0 || w.Left != 0);
      }

      private int GetIndexOfDocument(List<Window> windows, Document document)
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

      /// <summary> Clamp the given value to be between 0 and <paramref name="count"/>. </summary>
      private static int Clamp(int count, int number)
          => (number < 0 ? number + count : number) % count;
    }

    /// <summary> Command implementation for moving up/down. </summary>
    private class TabGroupMoverUpDown : TabGroupMover
    {
      /// <inheritdoc />
      protected override IEnumerable<Window> FilterAndSort(IEnumerable<Window> windows, Document activeDocument)
          => windows
              .Where(w => w.Document == activeDocument
                          || w.Left == activeDocument.ActiveWindow.Left)
              .OrderBy(w => w.Left)
              .ThenBy(w => w.Top);

      /// <inheritdoc />
      protected override bool ShouldMoveForward(int commandId)
          => commandId == CommandIdJumpDown;
    }

    /// <summary> Command implementation for moving left/right. </summary>
    private class TabGroupMoverLeftRight : TabGroupMover
    {
      /// <inheritdoc />
      protected override IEnumerable<Window> FilterAndSort(IEnumerable<Window> windows, Document activeDocument)
          => windows
              .Where(w => w.Document == activeDocument
                          || w.Left != activeDocument.ActiveWindow.Left)
              .OrderBy(w => w.Left)
              .ThenBy(w => w.Top);

      /// <inheritdoc />
      protected override bool ShouldMoveForward(int commandId)
          => commandId == CommandIdJumpRight;
    }

    /// <summary> Command implementation for moving next/previous. </summary>
    private class TabGroupMoverNextPrevious : TabGroupMover
    {
      /// <inheritdoc />
      protected override IEnumerable<Window> FilterAndSort(IEnumerable<Window> windows, Document activeDocument)
          => windows
              .OrderBy(w => w.Left)
              .ThenBy(w => w.Top);

      /// <inheritdoc />
      protected override bool ShouldMoveForward(int commandId)
          => commandId == CommandIdJumpNext;
    }
  }
}