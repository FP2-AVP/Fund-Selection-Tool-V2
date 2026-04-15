/**
 * Other Contributing Factors Module
 * Displays additional fund metrics including Sharpe, Information, Sortino, and Treynor ratios
 * for selected funds in both Annualized and Calendar view modes
 */

window.OtherFactors = {
  metrics: [
    'Return(Cumulative)',
    'Sharpe Ratio(Annualized)',
    'Sharpe Ratio (arith)(Annualized)',
    'Sharpe Ratio (geo)(Annualized)',
    'Information Ratio (arith)(Annualized)',
    'Information Ratio (geo)(Annualized)',
    'Sortino Ratio(Annualized)',
    'Sortino Ratio (arith)(Annualized)',
    'Sortino Ratio (geo)(Annualized)',
    'Treynor Ratio (arith)(Annualized)',
    'Treynor Ratio (geo)(Annualized)',
    'Max Drawdown'
  ],

  annualizedPeriods: ['YTD', '1Y', '3Y', '5Y', '10Y'],
  calendarYears: ['2016', '2017', '2018', '2019', '2020', '2021', '2022', '2023', '2024', '2025'],

  currentMode: 'annualized',
  currentPeriod: 'YTD',

  /**
   * Search for a column index matching a metric name with flexible matching
   */
  findColumnIndex(headers, metricName) {
    if (!headers || !metricName) return -1;

    const metric = metricName.trim().toLowerCase();

    for (let i = 0; i < headers.length; i++) {
      if (!headers[i]) continue;

      const headerLower = headers[i].toString().toLowerCase().trim();

      // Exact match
      if (headerLower === metric) return i;

      // Check if metric is contained in header
      if (headerLower.includes(metric)) return i;

      // Check header contains metric with flexible separators
      const headerNormalized = headerLower
        .replace(/[|\\/-]/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();

      const metricNormalized = metric
        .replace(/[|\\/-]/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();

      if (headerNormalized.includes(metricNormalized)) return i;
    }

    return -1;
  },

  /**
   * Find columns for a specific metric across all time periods
   */
  findMetricColumns(headers, metricName, periods) {
    const result = {};
    const baseMetric = metricName.split('(')[0].trim().toLowerCase();

    for (const period of periods) {
      // Try exact pattern: "Metric|(Period)" or "Metric (Period)"
      const patterns = [
        `${metricName}|(${period})`,
        `${metricName}|(${period})`,
        `${metricName}-(${period})`,
        `${metricName} (${period})`,
        `${metricName}|${period}`,
        `${metricName}-${period}`,
      ];

      let found = false;
      for (const pattern of patterns) {
        const idx = this.findColumnIndex(headers, pattern);
        if (idx !== -1) {
          result[period] = idx;
          found = true;
          break;
        }
      }

      // Fallback: search by metric + period anywhere in header
      if (!found) {
        for (let i = 0; i < headers.length; i++) {
          const header = headers[i].toString().toLowerCase();
          if (header.includes(baseMetric) && header.includes(period.toLowerCase())) {
            result[period] = i;
            break;
          }
        }
      }
    }

    return result;
  },

  /**
   * Format number for display with color coding
   */
  formatValue(value) {
    if (value === null || value === undefined || value === '') {
      return { text: '-', color: '#1e293b' };
    }

    const num = parseFloat(value);
    if (isNaN(num)) {
      return { text: String(value), color: '#1e293b' };
    }

    let text = num.toFixed(2);
    let color = '#1e293b'; // neutral dark

    if (num < 0) {
      color = '#dc2626'; // red for negative
    } else if (num > 0) {
      color = '#16a34a'; // dark green for positive
    }

    return { text, color };
  },

  /**
   * Get fund names and fund code columns from master data
   */
  getFundInfo(masterData) {
    if (!masterData || !masterData.values || masterData.values.length < 2) {
      return { fundNames: [], fundCodeIndices: [] };
    }

    const headers = masterData.values[0];
    const fundNames = [];
    const fundCodeIndices = [];

    // Find "Fund Code" column
    let fundCodeColIdx = -1;
    for (let i = 0; i < headers.length; i++) {
      const h = headers[i].toString().toLowerCase().trim();
      if (h === 'fund code' || h === 'fund_code' || h.includes('fund') && h.includes('code')) {
        fundCodeColIdx = i;
        break;
      }
    }

    if (fundCodeColIdx === -1) {
      return { fundNames: [], fundCodeIndices: [] };
    }

    // Collect unique fund codes
    const fundCodesSet = new Set();
    for (let row = 1; row < masterData.values.length; row++) {
      const fundCode = masterData.values[row][fundCodeColIdx];
      if (fundCode && fundCode.toString().trim()) {
        fundCodesSet.add(fundCode.toString().trim());
      }
    }

    return {
      fundNames: Array.from(fundCodesSet),
      fundCodeIndices: fundCodeColIdx,
      fundCodeColumn: fundCodeColIdx
    };
  },

  /**
   * Get data row for a specific fund
   */
  getFundRow(masterData, fundCode) {
    if (!masterData || !masterData.values || masterData.values.length < 2) {
      return null;
    }

    const { fundCodeColumn } = this.getFundInfo(masterData);
    if (fundCodeColumn === undefined) return null;

    for (let row = 1; row < masterData.values.length; row++) {
      if (masterData.values[row][fundCodeColumn] === fundCode) {
        return masterData.values[row];
      }
    }

    return null;
  },

  /**
   * Main render function
   */
  render(container) {
    if (!container) {
      console.error('OtherFactors: No container provided');
      return;
    }

    // Show loading state
    container.innerHTML = '<div style="padding: 20px; text-align: center; color: #666;">โหลดข้อมูล...</div>';

    // Load master data
    this.loadMasterData().then(masterData => {
      if (!masterData) {
        container.innerHTML = '<div style="padding: 20px; color: #dc2626;">ไม่สามารถโหลดข้อมูลได้</div>';
        return;
      }

      // Get selected funds
      const selectedFunds = (window.State && window.State.selected) ? Array.from(window.State.selected) : [];
      const { fundNames } = this.getFundInfo(masterData);

      const displayFunds = selectedFunds.length > 0
        ? selectedFunds.filter(f => fundNames.includes(f))
        : fundNames;

      if (displayFunds.length === 0) {
        container.innerHTML = '<div style="padding: 20px; color: #666; text-align: center;">กรุณาเลือกกองทุน</div>';
        return;
      }

      // Build HTML
      const html = this.buildHTML(masterData, displayFunds);
      container.innerHTML = html;

      // Attach event listeners
      this.attachEventListeners(container, masterData, displayFunds);
    }).catch(error => {
      console.error('OtherFactors: Error loading data', error);
      container.innerHTML = '<div style="padding: 20px; color: #dc2626;">เกิดข้อผิดพลาด: ' + error.message + '</div>';
    });
  },

  /**
   * Load master data from AppData or config URL
   */
  loadMasterData() {
    return new Promise((resolve, reject) => {
      // Try to use existing AppData
      if (window.AppData && window.AppData.master) {
        resolve(window.AppData.master);
        return;
      }

      // Try to fetch from config URL
      const url = (window.AppConfig && window.AppConfig.MASTER_URL) ||
                  (window.AppConfig && window.AppConfig.MASTER);

      if (!url) {
        reject(new Error('No master data source configured'));
        return;
      }

      fetch(url)
        .then(response => {
          if (!response.ok) throw new Error(`HTTP ${response.status}`);
          return response.json();
        })
        .then(data => {
          // Cache it
          if (!window.AppData) window.AppData = {};
          window.AppData.master = data;
          resolve(data);
        })
        .catch(error => reject(error));
    });
  },

  /**
   * Build the full HTML structure
   */
  buildHTML(masterData, displayFunds) {
    const styles = this.getStyles();
    const periods = this.currentMode === 'annualized' ? this.annualizedPeriods : this.calendarYears;
    const headers = masterData.values[0];

    // Build period selector HTML
    const periodButtonsHTML = periods.map(period => {
      const isActive = period === this.currentPeriod ? 'active' : '';
      return `<button class="period-btn ${isActive}" data-period="${period}">${period}</button>`;
    }).join('');

    // Build metric rows
    let tableRowsHTML = '';
    for (const metric of this.metrics) {
      const metricColumns = this.findMetricColumns(headers, metric, periods);
      const colIdx = metricColumns[this.currentPeriod];

      if (colIdx === undefined) continue;

      let rowHTML = `<tr><td class="metric-label">${metric}</td>`;

      for (const fund of displayFunds) {
        const fundRow = this.getFundRow(masterData, fund);
        if (!fundRow) {
          rowHTML += '<td class="data-cell">-</td>';
          continue;
        }

        const value = fundRow[colIdx];
        const { text, color } = this.formatValue(value);
        rowHTML += `<td class="data-cell" style="color: ${color};">${text}</td>`;
      }

      rowHTML += '</tr>';
      tableRowsHTML += rowHTML;
    }

    // Build fund header columns
    const fundHeadersHTML = displayFunds.map(fund => {
      return `<th class="fund-header">${fund}</th>`;
    }).join('');

    const html = `
      <div class="other-factors-container">
        ${styles}

        <div class="other-factors-header">
          <h2>ปัจจัยประกอบอื่นๆ</h2>

          <div class="controls">
            <div class="mode-selector">
              <button class="mode-btn active" data-mode="annualized">ประจำปี</button>
              <button class="mode-btn" data-mode="calendar">ปีปฏิทิน</button>
            </div>

            <div class="period-selector">
              ${periodButtonsHTML}
            </div>
          </div>
        </div>

        <div class="table-wrapper">
          <table class="data-table">
            <thead>
              <tr>
                <th class="metric-header">ตัวชี้วัด</th>
                ${fundHeadersHTML}
              </tr>
            </thead>
            <tbody>
              ${tableRowsHTML}
            </tbody>
          </table>
        </div>
      </div>
    `;

    return html;
  },

  /**
   * Get embedded CSS styles
   */
  getStyles() {
    return `
      <style id="other-factors-styles">
        .other-factors-container {
          background: #f8f9fa;
          border-radius: 8px;
          padding: 20px;
          font-family: 'Sarabun', 'THSarabunNew', sans-serif;
        }

        .other-factors-header {
          margin-bottom: 20px;
        }

        .other-factors-header h2 {
          margin: 0 0 15px 0;
          color: #1a2744;
          font-size: 20px;
          font-weight: 600;
        }

        .controls {
          display: flex;
          gap: 20px;
          flex-wrap: wrap;
          align-items: center;
        }

        .mode-selector,
        .period-selector {
          display: flex;
          gap: 8px;
        }

        .mode-btn,
        .period-btn {
          padding: 8px 16px;
          border: 1px solid #cbd5e1;
          background: white;
          color: #1e293b;
          border-radius: 20px;
          cursor: pointer;
          font-family: 'Sarabun', 'THSarabunNew', sans-serif;
          font-size: 13px;
          transition: all 0.2s ease;
        }

        .mode-btn:hover,
        .period-btn:hover {
          border-color: #94a3b8;
          background: #f1f5f9;
        }

        .mode-btn.active,
        .period-btn.active {
          background: #1a2744;
          color: white;
          border-color: #1a2744;
        }

        .table-wrapper {
          overflow-x: auto;
          background: white;
          border-radius: 6px;
          box-shadow: 0 1px 3px rgba(0,0,0,0.08);
        }

        .data-table {
          width: 100%;
          border-collapse: collapse;
          min-width: 600px;
        }

        .data-table thead {
          background: #1a2744;
          color: white;
        }

        .data-table th {
          padding: 12px;
          text-align: left;
          font-weight: 600;
          font-size: 13px;
          border-bottom: 2px solid #0f172a;
        }

        .metric-header {
          min-width: 200px;
          padding-left: 15px;
        }

        .fund-header {
          text-align: center;
          min-width: 120px;
        }

        .data-table tbody tr {
          border-bottom: 1px solid #e2e8f0;
          transition: background 0.15s ease;
        }

        .data-table tbody tr:hover {
          background: #f8fafc;
        }

        .data-table tbody tr:nth-child(even) {
          background: #fafbfc;
        }

        .metric-label {
          padding: 12px 15px;
          font-weight: 500;
          color: #1a2744;
          font-size: 13px;
          white-space: nowrap;
        }

        .data-cell {
          padding: 12px;
          text-align: center;
          font-size: 13px;
          font-weight: 500;
        }

        @media (max-width: 768px) {
          .controls {
            flex-direction: column;
            align-items: flex-start;
          }

          .mode-selector,
          .period-selector {
            flex-wrap: wrap;
          }

          .data-table {
            font-size: 12px;
          }

          .metric-label,
          .data-cell {
            padding: 8px;
          }
        }
      </style>
    `;
  },

  /**
   * Attach event listeners for mode and period switching
   */
  attachEventListeners(container, masterData, displayFunds) {
    const self = this;

    // Mode switcher
    const modeBtns = container.querySelectorAll('.mode-btn');
    modeBtns.forEach(btn => {
      btn.addEventListener('click', function() {
        const newMode = this.dataset.mode;
        if (newMode === self.currentMode) return;

        self.currentMode = newMode;
        self.currentPeriod = self.currentMode === 'annualized' ? 'YTD' : '2025';

        // Re-render
        modeBtns.forEach(b => b.classList.remove('active'));
        this.classList.add('active');

        const html = self.buildHTML(masterData, displayFunds);
        container.innerHTML = html;
        self.attachEventListeners(container, masterData, displayFunds);
      });
    });

    // Period switcher
    const periodBtns = container.querySelectorAll('.period-btn');
    periodBtns.forEach(btn => {
      btn.addEventListener('click', function() {
        const newPeriod = this.dataset.period;
        if (newPeriod === self.currentPeriod) return;

        self.currentPeriod = newPeriod;

        // Re-render
        periodBtns.forEach(b => b.classList.remove('active'));
        this.classList.add('active');

        const html = self.buildHTML(masterData, displayFunds);
        container.innerHTML = html;
        self.attachEventListeners(container, masterData, displayFunds);
      });
    });
  }
};
