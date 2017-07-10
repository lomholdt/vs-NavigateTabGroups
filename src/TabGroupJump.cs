//------------------------------------------------------------------------------
// <copyright file="TabGroupJump.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Linq;
using System.Runtime.InteropServices;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.TextManager.Interop;
using IServiceProvider = System.IServiceProvider;

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
    public static readonly Guid CommandSet = new Guid("06E4ACFA-6246-4E7F-B9EC-7843B718078C");

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

        var isMovingForward = ShouldMoveForward(commandId);
        var indexOfCurrentTabGroup = GetIndexOfDocument(topLevel, activeDocument, isMovingForward);

        // get the tab to activate
        var offset = isMovingForward ? 1 : -1;
        int nextIndex = Clamp(topLevel.Count, indexOfCurrentTabGroup + offset);

        // and activate it
        topLevel[nextIndex].Activate();
      }

      /// <summary>
      ///  Gets the IVsTextView associated with the given document so that it can be measured.
      /// </summary>
      private IVsTextView GetTextView(Document document)
      {
        uint itemID;
        IVsWindowFrame windowFrame;
        IVsUIHierarchy uiHierarchy;

        // TODO there must be a better way of getting the IVsTextView
        var isOpen = VsShellUtilities.IsDocumentOpen(_serviceProvider,
                                                     document.FullName,
                                                     Guid.Empty,
                                                     out uiHierarchy,
                                                     out itemID,
                                                     out windowFrame);

        object data;

        ErrorHandler.ThrowOnFailure(windowFrame.GetProperty(
                                      (int)__VSFPROPID.VSFPROPID_DocView,
                                      out data
                                    ));

        if (!(data is IVsTextView || data is IVsCodeWindow))
        {
        }

        return isOpen
          ? VsShellUtilities.GetTextView(windowFrame)
          : null;
      }

      private static IEnumerable<Window> GetValidDocuments(DTE2 dte)
      {
        // documents that are not the focused document in a group will have Top == 0 && Left == 0
        return dte.Windows.Cast<Window>()
                  .Where(w => w.Kind == "Document")
                  .Where(w => w.Top != 0 || w.Left != 0);
      }

      private int GetIndexOfDocument(List<Window> windows, Document activeDoc, bool isMovingForward)
      {
        int activeIdx = 0;

        // HACK: when we have a multi-pane editor (such as WPF design view + xml view), we'll have two
        // entries for the same document.  This causes problems when we're moving through the list:
        // Imagine that the multi-pane editor is at index N and N+1 (both the same document); when we
        // move from N to N+1, no problem, but when we then try to move forward again, if we start the
        // search for the "current" document at the beginning of the list, we'll think we're still at N
        // instead of N+1.  So, instead when we're moving forward through the list, we start the search
        // at the end of the list, and when we're moving backwards, we'll start the search at the
        // beginning.
        // 
        // Note that this is a hack; we should be looking at the current editor window, not the current
        // document; we can fix this if anyone ever complains.
        // 
        // (another way to fix it would be to ignore non-editor panes) 
        if (isMovingForward)
        {
          for (int i = windows.Count - 1; i >= 0; --i)
          {
            if (windows[i].Document == activeDoc)
            {
              activeIdx = i;
              break;
            }
          }
        }
        else
        {
          for (int i = 0; i < windows.Count; ++i)
          {
            if (windows[i].Document == activeDoc)
            {
              activeIdx = i;
              break;
            }
          }
        }

        return activeIdx;
      }

      /// <summary> Clamp the given value to be between 0 and <paramref name="count"/>. </summary>
      private static int Clamp(int count, int number)
        => (number < 0 ? number + count : number) % count;

      /// <summary> Measure the bounds of the given window </summary>
      internal RECT MeasureBounds(Window window)
      {
        var textView = GetTextView(window.Document);

        RECT rect;

        if (textView != null && GetWindowRect(textView.GetWindowHandle(), out rect))
        {
          return rect;
        }

        // fallback where Top is wrong for windows that are vertically split. 
        return new RECT
               {
                 left = window.Left,
                 right = window.Left + window.Width,
                 top = window.Top,
                 bottom = window.Top + window.Height
               };
      }

      [DllImport("user32.dll", SetLastError = true)]
      private static extern bool GetWindowRect(IntPtr hwnd, out RECT lpRect);
    }

    /// <summary> Command implementation for moving up/down. </summary>
    private class TabGroupMoverUpDown : TabGroupMover
    {
      /// <inheritdoc />
      protected override IEnumerable<Window> FilterAndSort(IEnumerable<Window> windows, Document activeDocument)
        => from w in windows
           where w.Document == activeDocument
                 || w.Left == activeDocument.ActiveWindow.Left
           let rect = MeasureBounds(w)
           orderby rect.left, rect.top
           select w;

      /// <inheritdoc />
      protected override bool ShouldMoveForward(int commandId)
        => commandId == CommandIdJumpDown;
    }

    /// <summary> Command implementation for moving left/right. </summary>
    private class TabGroupMoverLeftRight : TabGroupMover
    {
      /// <inheritdoc />
      protected override IEnumerable<Window> FilterAndSort(IEnumerable<Window> windows, Document activeDocument)
        => from w in windows
           where w.Document == activeDocument
                 || w.Left != activeDocument.ActiveWindow.Left
           let rect = MeasureBounds(w)
           orderby rect.left, rect.top
           select w;

      /// <inheritdoc />
      protected override bool ShouldMoveForward(int commandId)
        => commandId == CommandIdJumpRight;
    }

    /// <summary> Command implementation for moving next/previous. </summary>
    private class TabGroupMoverNextPrevious : TabGroupMover
    {
      /// <inheritdoc />
      protected override IEnumerable<Window> FilterAndSort(IEnumerable<Window> windows, Document activeDocument)
        => from w in windows
           let rect = MeasureBounds(w)
           orderby rect.left, rect.top
           select w;

      /// <inheritdoc />
      protected override bool ShouldMoveForward(int commandId)
        => commandId == CommandIdJumpNext;
    }
  }
}