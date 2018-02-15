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

      _tabGroupMoverLeftRight = new TabGroupMoverLeftRight(ServiceProvider);
      _tabGroupMoverUpDown = new TabGroupMoverUpDown(ServiceProvider);
      _tabGroupMoverNextPrevious = new TabGroupMoverNextPrevious(ServiceProvider);

      if (ServiceProvider.GetService(typeof(IMenuCommandService)) is OleMenuCommandService commandService)
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
    /// Gets the instance of the package.
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
      private readonly IServiceProvider _serviceProvider;
      private readonly IVsUIShell _uiShell;

      protected TabGroupMover(IServiceProvider serviceProvider)
      {
        _serviceProvider = serviceProvider;
        _uiShell = (IVsUIShell)serviceProvider.GetService(typeof(SVsUIShell));
      }

      /// <summary>
      ///  How the document-windows should ordered.  They should be ordered such that when
      ///  <see cref="ShouldMoveForward"/> returns true the Nth + 1 element is the logical group to jump
      ///  to.  For example, if you sort the groups vertically (smallest first), then ShouldMoveFoward()
      ///  should return true when moving down.
      /// </summary>
      /// <seealso cref="ShouldMoveForward"/>
      protected abstract IEnumerable<WindowData> FilterAndSort(IEnumerable<WindowData> windows,
                                                               WindowData activeDocument);

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

        var activeWindows = GetWindowData(dte).ToList();

        WindowData activeWindow = null;

        foreach (var windowData in activeWindows)
        {
          if (windowData.Window == dte.ActiveWindow)
          {
            activeWindow = windowData;
            break;
          }
        }

        if (activeWindow == null)
          return;

        var topLevel = FilterAndSort(activeWindows, activeWindow)
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
        var indexOfCurrentTabGroup = GetIndexofActiveWindow(topLevel, dte.ActiveWindow);

        // get the tab to activate
        var offset = isMovingForward ? 1 : -1;
        int nextIndex = Clamp(topLevel.Count, indexOfCurrentTabGroup + offset);

        // and activate it
        topLevel[nextIndex].Window.Activate();
      }

      /// <summary> Get all of the Windows that have an associated frame. </summary>
      private IEnumerable<WindowData> GetWindowData(DTE2 dte)
      {
        var existingWindows = new HashSet<Window>(GetActiveWindows(dte));

        var frames = GetFrames();

        foreach (var frame in frames)
        {
          if (frame.GetProperty((int)__VSFPROPID.VSFPROPID_ExtWindowObject, out var window) == 0
              && window is Window typedWindow
              && existingWindows.Contains(typedWindow)
            )
          {
            yield return new WindowData(typedWindow, frame);
          }
        }
      }

      /// <summary> Gets all of the windows that are currently positioned with a valid Top or Left. </summary>
      private IEnumerable<Window> GetActiveWindows(DTE2 dte)
      {
        // documents that are not the focused document in a group will have Top == 0 && Left == 0
        return from window in dte.Windows.Cast<Window>()
               where window.Kind == "Document"
               where window.Top != 0 || window.Left != 0
               select window;
      }

      /// <summary> Get all known <see cref="IVsWindowFrame"/>, lazily. </summary>
      private IEnumerable<IVsWindowFrame> GetFrames()
      {
        var array = new IVsWindowFrame[1];
        _uiShell.GetDocumentWindowEnum(out var frames);

        while (true)
        {
          var errorCode = frames.Next(1, array, out _);
          if (errorCode != VSConstants.S_OK)
            break;

          yield return array[0];
        }
      }

      private static int GetIndexofActiveWindow(List<WindowData> windows, Window activeWindow)
      {
        for (var i = 0; i < windows.Count; i++)
        {
          var data = windows[i];
          if (data.Window == activeWindow)
            return i;
        }

        return 0;
      }

      /// <summary> Clamp the given value to be between 0 and <paramref name="count"/>. </summary>
      private static int Clamp(int count, int number)
        => (number < 0 ? number + count : number) % count;
    }

    internal class WindowData
    {
      public Window Window { get; }
      public IVsWindowFrame AssociatedFrame { get; }

      private RECT? _measuredRect;

      public WindowData(Window window, IVsWindowFrame associatedFrame)
      {
        Window = window;
        AssociatedFrame = associatedFrame;
        Bounds = MeasureBounds();
      }

      public RECT Bounds { get; }

      /// <summary> Measure the bounds of the given window </summary>
      private RECT MeasureBounds()
      {
        var window = Window;
        var textView = VsShellUtilities.GetTextView(AssociatedFrame);

        if (textView != null && GetWindowRect(textView.GetWindowHandle(), out var rect))
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
      public TabGroupMoverUpDown(IServiceProvider serviceProvider)
        : base(serviceProvider)
      {
      }

      /// <inheritdoc />
      protected override IEnumerable<WindowData> FilterAndSort(IEnumerable<WindowData> windows,
                                                               WindowData activeWindow)
        => from w in windows
           where w == activeWindow
                  // only return those that aren't aligned vertically
                 || w.Bounds.left == activeWindow.Bounds.left
           orderby w.Bounds.left, w.Bounds.top
           select w;

      /// <inheritdoc />
      protected override bool ShouldMoveForward(int commandId)
        => commandId == CommandIdJumpDown;
    }

    /// <summary> Command implementation for moving left/right. </summary>
    private class TabGroupMoverLeftRight : TabGroupMover
    {
      public TabGroupMoverLeftRight(IServiceProvider serviceProvider)
        : base(serviceProvider)
      {
      }

      /// <inheritdoc />
      protected override IEnumerable<WindowData> FilterAndSort(IEnumerable<WindowData> windows,
                                                               WindowData activeWindow)
      {
        return from w in windows
               // only return those that aren't aligned vertically
               where w == activeWindow
                     || w.Bounds.left != activeWindow.Bounds.left
               orderby w.Bounds.left, w.Bounds.top
               select w;
      }

      /// <inheritdoc />
      protected override bool ShouldMoveForward(int commandId)
        => commandId == CommandIdJumpRight;
    }

    /// <summary> Command implementation for moving next/previous. </summary>
    private class TabGroupMoverNextPrevious : TabGroupMover
    {
      public TabGroupMoverNextPrevious(IServiceProvider serviceProvider)
        : base(serviceProvider)
      {
      }

      /// <inheritdoc />
      protected override IEnumerable<WindowData> FilterAndSort(IEnumerable<WindowData> windows,
                                                               WindowData activeWindow)
        => from w in windows
           orderby w.Bounds.left, w.Bounds.top
           select w;

      /// <inheritdoc />
      protected override bool ShouldMoveForward(int commandId)
        => commandId == CommandIdJumpNext;
    }
  }
}