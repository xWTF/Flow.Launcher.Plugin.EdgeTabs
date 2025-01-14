using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Timers;
using System.Windows.Automation;

namespace Flow.Launcher.Plugin.EdgeTabs
{
    public class EdgeTabsPlugin : IPlugin
    {
        public static readonly PropertyCondition
            condClassChrome = new PropertyCondition(AutomationElement.ClassNameProperty, "Chrome_WidgetWin_1"),
            condClassBrowserRootView = new PropertyCondition(AutomationElement.ClassNameProperty, "BrowserRootView"),
            condClassNonClientView = new PropertyCondition(AutomationElement.ClassNameProperty, "NonClientView"),
            condClassBrowserFrameViewWin = new PropertyCondition(AutomationElement.ClassNameProperty, "BrowserFrameViewWin"),
            condClassBrowserView = new PropertyCondition(AutomationElement.ClassNameProperty, "BrowserView"),
            condClassTopContainerView = new PropertyCondition(AutomationElement.ClassNameProperty, "TopContainerView"),
            condClassEdgeVerticalTabContainerView = new PropertyCondition(AutomationElement.ClassNameProperty, "EdgeVerticalTabContainerView"),
            condClassContentsBackgroundView = new PropertyCondition(AutomationElement.ClassNameProperty, "ContentsBackgroundView"),
            condClassEdgeTabStripRegionView = new PropertyCondition(AutomationElement.ClassNameProperty, "EdgeTabStripRegionView"),
            condClassEdgeTabStrip = new PropertyCondition(AutomationElement.ClassNameProperty, "EdgeTabStrip"),
            condClassEdgeTabContainerImpl = new PropertyCondition(AutomationElement.ClassNameProperty, "EdgeTabContainerImpl"),
            condClassScrollView = new PropertyCondition(AutomationElement.ClassNameProperty, "ScrollView"),
            condClassScrollViewViewport = new PropertyCondition(AutomationElement.ClassNameProperty, "ScrollView::Viewport"),
            condClassView = new PropertyCondition(AutomationElement.ClassNameProperty, "View"),
            condTypeTabItem = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem);

        private PluginInitContext context;

        class hwndCacheEntry
        {
            public AutomationElement browserView, topContainer, sideContentsBackgroundView;
            public DateTime expires;
        }

        Dictionary<IntPtr, hwndCacheEntry> hwndCache = new Dictionary<IntPtr, hwndCacheEntry>();
        public Timer hwndCacheTimer = new Timer(60 * 1000);

        public List<Result> resultCache = new List<Result>();
        public DateTime resultCacheExpire = DateTime.MinValue;

        public EdgeTabsPlugin()
        {
            hwndCacheTimer.Elapsed += (sender, e) =>
            {
                lock (hwndCache)
                {
                    var now = DateTime.Now;
                    foreach (var entry in hwndCache)
                    {
                        if (entry.Value.expires < now)
                        {
                            hwndCache.Remove(entry.Key);
                        }
                    }
                }
            };
            hwndCacheTimer.Start();
        }

        public void Init(PluginInitContext context)
        {
            this.context = context;
        }

        AutomationElement findChildrenNested(AutomationElement element, params PropertyCondition[] conditions)
        {
            for (int i = 0; i < conditions.Length && element != null; i++)
            {
                element = element.FindFirst(TreeScope.Children, conditions[i]);
            }
            return element;
        }

        hwndCacheEntry createHwndCacheFor(IntPtr hwnd)
        {
            lock (hwndCache)
            {
                if (!hwndCache.TryGetValue(hwnd, out hwndCacheEntry cache))
                {
                    cache = new hwndCacheEntry
                    {
                        browserView = findChildrenNested(AutomationElement.FromHandle(hwnd), condClassBrowserRootView, condClassNonClientView, condClassBrowserFrameViewWin, condClassBrowserView),
                    };
                    if (cache.browserView != null)
                    {
                        cache.topContainer = findChildrenNested(cache.browserView, condClassTopContainerView);
                        cache.sideContentsBackgroundView = findChildrenNested(cache.browserView, condClassEdgeVerticalTabContainerView, condClassContentsBackgroundView);

                        if (cache.topContainer != null)
                        {
                            cache.expires = DateTime.Now.AddMinutes(10);
                        }
                        else
                        {
                            cache.browserView = cache.sideContentsBackgroundView = null;
                        }
                    }
                    if (cache.browserView == null)
                    {
                        cache.expires = DateTime.Now.AddMinutes(1);
                    }
                    hwndCache[hwnd] = cache;
                }
                return cache;
            }
        }

        public static List<IntPtr> GetEdgeWindows()
        {
            var windows = new List<IntPtr>();
            EnumWindows(delegate (IntPtr hWnd, int lParam)
            {
                if (!IsWindowVisible(hWnd)) return true;

                var sb = new StringBuilder(32);
                if (GetClassName(hWnd, sb, sb.Capacity) == 0 || sb.ToString() != "Chrome_WidgetWin_1") return true;

                int length = GetWindowTextLength(hWnd);
                if (length == 0) return true;

                sb = new StringBuilder(length);
                GetWindowText(hWnd, sb, length + 1);
                if (sb.ToString().EndsWith("- Microsoft​ Edge"))
                {
                    windows.Add(hWnd);
                }
                return true;
            }, 0);
            return windows;
        }

        private delegate bool EnumWindowsProc(IntPtr hWnd, int lParam);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc enumFunc, int lParam);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowTextLength(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        List<Result> getResults()
        {
            lock (resultCache)
            {
                if (resultCacheExpire > DateTime.Now)
                {
                    return resultCache;
                }

                var results = new List<Result>();
                void insertTabs(AutomationElement tabItem)
                {
                    results.Add(new Result
                    {
                        Title = tabItem.Current.Name,
                        SubTitle = "Edge Tab",
                        IcoPath = "icon.png",
                        Action = c =>
                        {
                            tabItem.SetFocus();
                            return true;
                        },
                        Score = 100
                    });
                }

                foreach (IntPtr wnd in GetEdgeWindows())
                {
                    var cache = createHwndCacheFor(wnd);
                    if (cache.browserView == null) continue;

                    var vertical = false;
                    var tabContainer = findChildrenNested(cache.topContainer, condClassEdgeTabStripRegionView, condClassEdgeTabStrip, condClassEdgeTabContainerImpl);
                    if (tabContainer == null)
                    {
                        // Sometimes UIA fails to find the EdgeVerticalTabContainerView
                        if (cache.sideContentsBackgroundView == null)
                        {
                            cache.sideContentsBackgroundView = findChildrenNested(cache.browserView, condClassEdgeVerticalTabContainerView, condClassContentsBackgroundView);
                        }
                        if (cache.sideContentsBackgroundView == null) continue;

                        vertical = true;
                        tabContainer = findChildrenNested(cache.sideContentsBackgroundView, condClassEdgeTabStrip, condClassEdgeTabContainerImpl);
                    }
                    if (tabContainer == null) continue;

                    // Horizental tabs, or pinned ones in vertical mode
                    foreach (AutomationElement tab in tabContainer.FindAll(TreeScope.Children, condTypeTabItem))
                    {
                        insertTabs(tab);
                    }

                    if (vertical)
                    {
                        // Other tabs in vertical mode
                        var view = findChildrenNested(tabContainer, condClassScrollView, condClassScrollViewViewport, condClassView);
                        if (view != null)
                        {
                            foreach (AutomationElement tab in view.FindAll(TreeScope.Children, condTypeTabItem))
                            {
                                insertTabs(tab);
                            }
                        }
                    }
                }

                resultCache = results;
                resultCacheExpire = DateTime.Now.AddSeconds(5);
                return results;
            }
        }

        public List<Result> Query(Query query)
        {
            var results = getResults();
            if (query.ActionKeyword != string.Empty && query.Search != string.Empty)
            {
                results = results.Where(r =>
                {
                    r.Score = context.API.FuzzySearch(query.Search, r.Title).Score;
                    return r.Score > 0;
                }).ToList();
            }
            return results;
        }
    }
}
