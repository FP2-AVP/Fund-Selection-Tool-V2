/**
 * Other Factors Patch Helper
 * This file provides instructions for integrating the Other Factors section into app.js
 *
 * USAGE:
 * 1. Copy and paste this entire file into the browser console on the Fund List Tool app
 * 2. It will display clear instructions on what needs to be added to app.js
 * 3. Follow the instructions to update your navigation menu
 */

(function() {
  console.clear();
  console.log('%c=== Other Factors Module Integration Instructions ===', 'color: #1a2744; font-size: 16px; font-weight: bold;');
  console.log('%c', '');

  console.log('%cSTEP 1: Add the Script Tag to index.html', 'color: #16a34a; font-size: 14px; font-weight: bold;');
  console.log('Add this line to your <head> section in index.html (after the other script imports):');
  console.log('%c<script src="js/other-factors.js"><\/script>', 'background: #f1f5f9; padding: 10px; display: block; white-space: pre; font-family: monospace;');
  console.log('%c', '');

  console.log('%cSTEP 2: Add Navigation Menu Item to app.js', 'color: #16a34a; font-size: 14px; font-weight: bold;');
  console.log('In app.js, find the navigation menu setup (look for the menu/nav section)');
  console.log('Add this new menu item after "ค่าธรรมเนียม 2":');
  console.log('%c', '');

  console.log('%cMenu Item Code to Add:', 'color: #1a2744; font-weight: bold;');
  const menuItemCode = `{
  label: 'ปัจจัยประกอบอื่นๆ',
  id: 'other-factors',
  render: function(container) {
    window.OtherFactors.render(container);
  }
}`;
  console.log('%c' + menuItemCode, 'background: #f1f5f9; padding: 10px; display: block; white-space: pre; font-family: monospace;');
  console.log('%c', '');

  console.log('%cSTEP 3: Verify Integration', 'color: #16a34a; font-size: 14px; font-weight: bold;');
  console.log('After adding the code:');
  console.log('  1. Refresh the page (Ctrl+R or Cmd+R)');
  console.log('  2. Look for "ปัจจัยประกอบอื่นๆ" in the navigation menu');
  console.log('  3. Click it to view the Other Factors comparison table');
  console.log('%c', '');

  console.log('%cFeatures Available:', 'color: #1a2744; font-weight: bold;');
  console.log('  ✓ Display metrics: Sharpe, Information, Sortino, Treynor Ratios');
  console.log('  ✓ Toggle between Annualized (YTD, 1Y, 3Y, 5Y, 10Y) and Calendar years (2016-2025)');
  console.log('  ✓ Color-coded values: Positive (green), Negative (red), Neutral (dark)');
  console.log('  ✓ Works with State.selected fund filtering');
  console.log('  ✓ Responsive design with hover effects');
  console.log('%c', '');

  console.log('%cData Format Requirements:', 'color: #1a2744; font-weight: bold;');
  console.log('  • The module reads from window.AppData.master');
  console.log('  • Column names should follow patterns like: "Metric|(Period)" or "Metric (Period)"');
  console.log('  • Supported periods: YTD, 1Y, 3Y, 5Y, 10Y (Annualized)');
  console.log('  • Supported periods: 2016-2025 (Calendar years)');
  console.log('%c', '');

  console.log('%cMetrics Included:', 'color: #1a2744; font-weight: bold;');
  console.log('  • Return(Cumulative)');
  console.log('  • Sharpe Ratio(Annualized)');
  console.log('  • Sharpe Ratio (arith)(Annualized)');
  console.log('  • Sharpe Ratio (geo)(Annualized)');
  console.log('  • Information Ratio (arith)(Annualized)');
  console.log('  • Information Ratio (geo)(Annualized)');
  console.log('  • Sortino Ratio(Annualized)');
  console.log('  • Sortino Ratio (arith)(Annualized)');
  console.log('  • Sortino Ratio (geo)(Annualized)');
  console.log('  • Treynor Ratio (arith)(Annualized)');
  console.log('  • Treynor Ratio (geo)(Annualized)');
  console.log('  • Max Drawdown');
  console.log('%c', '');

  console.log('%cTroubleshooting:', 'color: #1a2744; font-weight: bold;');
  console.log('  • If menu item doesn\'t appear: Check that other-factors.js is loaded (check Network tab)');
  console.log('  • If no data shows: Verify window.AppData.master is populated');
  console.log('  • If metrics are missing: Check column headers in your JSON match expected format');
  console.log('%c', '');

  console.log('%cNeed help? Check the browser console for any error messages.', 'color: #666; font-style: italic;');

  // Also provide a helper to verify setup
  window.OtherFactorsDebug = {
    checkSetup: function() {
      console.clear();
      console.log('%c=== Other Factors Setup Diagnostic ===', 'color: #1a2744; font-size: 16px; font-weight: bold;');

      const checks = {
        'Script Loaded': window.OtherFactors !== undefined,
        'AppData Available': window.AppData !== undefined,
        'Master Data Loaded': window.AppData && window.AppData.master !== undefined,
        'State Available': window.State !== undefined,
        'AppConfig Available': window.AppConfig !== undefined
      };

      Object.entries(checks).forEach(([label, status]) => {
        const icon = status ? '✓' : '✗';
        const color = status ? '#16a34a' : '#dc2626';
        console.log(`%c${icon} ${label}`, `color: ${color}; font-weight: bold;`);
      });

      if (window.AppData && window.AppData.master) {
        const rows = window.AppData.master.values ? window.AppData.master.values.length : 0;
        const cols = window.AppData.master.values ? window.AppData.master.values[0].length : 0;
        console.log(`\n%cMaster Data Shape: ${rows} rows × ${cols} columns`, 'color: #1a2744;');

        if (window.AppData.master.values && window.AppData.master.values[0]) {
          console.log('%cFirst few columns:', 'color: #1a2744; font-weight: bold;');
          const headers = window.AppData.master.values[0].slice(0, 5);
          headers.forEach((h, i) => console.log(`  [${i}] ${h}`));
        }
      }
    }
  };

  console.log('%c📋 Tip: Run %cwindow.OtherFactorsDebug.checkSetup()%c to verify your setup', 'color: #666; font-style: italic;', 'background: #f1f5f9; padding: 2px 6px; font-family: monospace;', 'color: #666; font-style: italic;');
})();
