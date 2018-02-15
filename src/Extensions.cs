using System;
using System.Collections.Generic;
using System.Linq;

namespace TabGroupJumperVSIX
{
  /// <summary> Useful extension methods. </summary>
  public static class Extensions
  {
    public static T GetService<T>(this IServiceProvider serviceProvider)
      => (T)serviceProvider.GetService(typeof(T));
  }
}