// List of template lines
const templateLines = [
  "Current Assets",
  "Average Inventory",
  "Inventories",
  "Average Net Fixed Assets",
  "Average Total Assets",
  "Investment",
  "Total Assets",
  "Current Liabilities",
  "Average Accounts Payable",
  "Long Term Debt",
  "Total Liabilities",
  "Number of Outstanding Shares",
  "Shareholders Equity",
  "Net Sales",
  "COGS",
  "Purchases of COGS",
  "EBITDA",
  "Net Income",
  "Net Revenue",
  "Net Profit",
  "EBIT",
  "Interest",
  "Tax",
  "Average Working Capital",
  "Cash Flow From Operating Activities",
  "Capital Expenditures",
  "Principal",
  "Dividends",
  "Market Value per Share",
  "Earnings Per Share (EPS)",
  "Non-Operating Cash"
];

// Store loaded data
let loadedData = {};

// Utility elements
const statusDiv = document.getElementById('status');
const resultsDiv = document.getElementById('results');

// Buttons
const btnDownload = document.getElementById('downloadTemplate');
const btnLoad = document.getElementById('loadTemplate');
const btnRefresh = document.getElementById('refreshTemplate');
const fileInput = document.getElementById('fileInput');

// New category buttons
const btnShowGeneral = document.getElementById('showGeneral');
const btnShowOperative = document.getElementById('showOperative');
const btnShowCashFlow = document.getElementById('showCashFlow');
const btnShowReturns = document.getElementById('showReturns');

// Handle Download Template
btnDownload.addEventListener('click', () => {
  const wb = XLSX.utils.book_new();
  const ws_data = [['Parameter', 'Value']];
  templateLines.forEach(line => {
    ws_data.push([line, '']);
  });
  const ws = XLSX.utils.aoa_to_sheet(ws_data);
  XLSX.utils.book_append_sheet(wb, ws, 'Template');

  XLSX.writeFile(wb, 'financial_template.xlsx');
  updateStatus('Template downloaded as "financial_template.xlsx".');
});

// Handle Load Template
btnLoad.addEventListener('click', () => {
  fileInput.click();
});

fileInput.addEventListener('change', (e) => {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (evt) => {
    try {
      const data = new Uint8Array(evt.target.result);
      const workbook = XLSX.read(data, {type: 'array'});
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(sheet, {header:1});
      parseLoadedData(jsonData);
      updateStatus('Template loaded successfully.');
    } catch (err) {
      updateStatus('Error loading file: ' + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
});

// Handle Refresh KPI display
btnRefresh.addEventListener('click', () => {
  generateAndDisplayResults();
});

// Category buttons event listeners
btnShowGeneral.addEventListener('click', () => {
  displayCategory('general');
});
btnShowOperative.addEventListener('click', () => {
  displayCategory('operative');
});
btnShowCashFlow.addEventListener('click', () => {
  displayCategory('cashflow');
});
btnShowReturns.addEventListener('click', () => {
  displayCategory('returns');
});

// Parse loaded data from the sheet
function parseLoadedData(data) {
  loadedData = {}; // Reset previous data
  if (!data || data.length === 0) {
    updateStatus('Empty or invalid data.');
    return;
  }

  // Expect first row to be headers
  const headers = data[0];
  if (
    !headers ||
    headers.length < 2 ||
    headers[0].toString().toLowerCase() !== 'parameter' ||
    headers[1].toString().toLowerCase() !== 'value'
  ) {
    updateStatus('Invalid template format. Make sure the first row contains headers "Parameter" and "Value".');
    return;
  }

  // Parse rows
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row.length >= 2) {
      const param = row[0];
      const value = row[1];
      if (param) {
        loadedData[param] = value;
      }
    }
  }
}

// Generate and display all KPIs
function generateAndDisplayResults() {
  resultsDiv.innerHTML = '';
  displayAllKPIs();
}

// Display all KPIs
function displayAllKPIs() {
  const params = loadedData;

  const getVal = (param) => {
    const val = params[param];
    const num = parseFloat(val);
    return isNaN(num) ? 0 : num;
  };

  // Retrieve values
  const currentAssets = getVal('Current Assets');
  const inventories = getVal('Inventories');
  const currentLiabilities = getVal('Current Liabilities');
  const totalLiabilities = getVal('Total Liabilities');
  const shareholdersEquity = getVal('Shareholders Equity');
  const totalAssets = getVal('Total Assets');
  const outstandingShares = getVal('Number of Outstanding Shares');
  const cogs = getVal('COGS');
  const avgInventory = getVal('Average Inventory');
  const purchasesOfCogs = getVal('Purchases of COGS');
  const netSales = getVal('Net Sales');
  const avgTotalAssets = getVal('Average Total Assets');
  const avgNetFixedAssets = getVal('Average Net Fixed Assets');
  const netIncome = getVal('Net Income');
  const netRevenue = getVal('Net Revenue');
  const dividends = getVal('Dividends');
  const marketValuePerShare = getVal('Market Value per Share');
  const eps = getVal('Earnings Per Share (EPS)');
  const cashFlowOps = getVal('Cash Flow From Operating Activities');
  const capEx = getVal('Capital Expenditures');
  const principal = getVal('Principal');
  const interest = getVal('Interest');
  const EBITDA = getVal('EBITDA');
  const EBIT = getVal('EBIT');
  const tax = getVal('Tax');
  const netProfit = getVal('Net Profit');

  // Calculations
  const currentRatio = safeDiv(currentAssets, currentLiabilities);
  const quickRatio = safeDiv(currentAssets - inventories, currentLiabilities);
  const debtToEquity = safeDiv(totalLiabilities, shareholdersEquity);
  const debtRatio = safeDiv(totalLiabilities, totalAssets);
  const workingCapital = currentAssets - currentLiabilities;
  const equityRatio = safeDiv(shareholdersEquity, totalAssets);
  const bookValuePerShare = safeDiv(shareholdersEquity, outstandingShares);
  const inventoryTurnover = safeDiv(cogs, avgInventory);
  const accountsPayableTurnover = safeDiv(purchasesOfCogs, getVal('Average Accounts Payable') || 1);
  const assetsTurnoverRatio = safeDiv(netSales, avgTotalAssets);
  const fixedAssetTurnover = safeDiv(netSales, avgNetFixedAssets);
  const workingCapitalTurnover = safeDiv(netSales, safeDiv(workingCapital,1));
  const operCashFlowRatio = safeDiv(cashFlowOps, currentLiabilities);
  const freeCashFlow = cashFlowOps - capEx;
  const cashFlowToDebtRatio = safeDiv(cashFlowOps, totalLiabilities);
  const cashFlowMargin = safeDiv(cashFlowOps, netSales);
  const cashReturnOnAssets = safeDiv(cashFlowOps, totalAssets);
  const debtCoverageRatio = safeDiv(EBITDA, (principal + interest));
  const cashConversionRatio = safeDiv(cashFlowOps, netIncome);
  const ROA = safeDiv(netIncome, totalAssets);
  const ROE = safeDiv(netIncome, shareholdersEquity);
  const turnoverRatio = safeDiv(netRevenue, totalAssets);
  const dividendPayoutRatio = safeDiv(dividends, netIncome);
  const earningPerShare = safeDiv(netIncome, outstandingShares);
  const peRatio = safeDiv(marketValuePerShare, eps);
  const ROI = safeDiv(netProfit, getVal('Investment'));
  const ROIC = safeDiv(EBIT * (1 - (tax/100)), (getVal('Long Term Debt') + shareholdersEquity - getVal('Non-Operating Cash')));
  const ROE2 = safeDiv(EBIT - interest - tax, shareholdersEquity);
  const ROCE = safeDiv(EBIT, (getVal('Long Term Debt') + shareholdersEquity));
  const ROA2 = safeDiv(EBIT - interest - tax, totalAssets);

  // Full KPI list
  const lines = [
    // General
    `Current Ratio: ${currentRatio.toFixed(2)}`,
    `Quick Ratio: ${quickRatio.toFixed(2)}`,
    `Debt to Equity Ratio: ${debtToEquity.toFixed(2)}`,
    `Debt Ratio: ${debtRatio.toFixed(2)}`,
    `Shareholders Equity: ${shareholdersEquity}`,
    `Book Value per Share: ${bookValuePerShare.toFixed(2)}`,
    `Earnings Per Share (EPS): ${earningPerShare.toFixed(2)}`,
    `Price Earnings (P/E) Ratio: ${peRatio.toFixed(2)}`,
    // Operative
    `Working Capital: ${workingCapital}`,
    `Inventory Turnover: ${inventoryTurnover.toFixed(2)}`,
    `Accounts Payable Turnover: ${accountsPayableTurnover.toFixed(2)}`,
    `Assets Turnover Ratio: ${assetsTurnoverRatio.toFixed(2)}`,
    `Fixed Asset Turnover: ${fixedAssetTurnover.toFixed(2)}`,
    `Working Capital Turnover: ${workingCapitalTurnover.toFixed(2)}`,
    `Turnover Ratio: ${turnoverRatio.toFixed(2)}`,
    `Dividend Payout Ratio: ${dividendPayoutRatio.toFixed(2)}`,
    // Cash Flow
    `Operating Cash Flow Ratio: ${operCashFlowRatio.toFixed(2)}`,
    `Free Cash Flow: ${freeCashFlow}`,
    `Cash Flow to Debt Ratio: ${cashFlowToDebtRatio.toFixed(2)}`,
    `Cash Flow Margin: ${cashFlowMargin.toFixed(2)}`,
    `Cash Return on Assets (Cash ROA): ${cashReturnOnAssets.toFixed(2)}`,
    `Debt Coverage Ratio: ${debtCoverageRatio.toFixed(2)}`,
    `Cash Conversion Ratio: ${cashConversionRatio.toFixed(2)}`,
    `Returns On Assets (ROA): ${ROA.toFixed(2)}`,
    `Return On Equity (ROE): ${ROE.toFixed(2)}`,
    `ROI: ${ROI.toFixed(2)}`,
    `ROIC: ${ROIC.toFixed(2)}`,
    `ROE (alternative): ${ROE2.toFixed(2)}`,
    `ROCE: ${ROCE.toFixed(2)}`,
    `ROA (alternative): ${ROA2.toFixed(2)}`
  ];

  resultsDiv.innerHTML = lines.join('<br>');
}

// Functions to display specific categories
function displayCategory(category) {
  resultsDiv.innerHTML = '';
  const params = loadedData;

  const getVal = (param) => {
    const val = params[param];
    const num = parseFloat(val);
    return isNaN(num) ? 0 : num;
  };

  switch (category) {
    case 'general':
      resultsDiv.innerHTML = `
        Current Ratio: ${safeDiv(getVal('Current Assets'), getVal('Current Liabilities')).toFixed(2)}<br>
        Quick Ratio: ${safeDiv(getVal('Current Assets') - getVal('Inventories'), getVal('Current Liabilities')).toFixed(2)}<br>
        Debt to Equity Ratio: ${safeDiv(getVal('Total Liabilities'), getVal('Shareholders Equity')).toFixed(2)}<br>
        Debt Ratio: ${safeDiv(getVal('Total Liabilities'), getVal('Total Assets')).toFixed(2)}<br>
        Shareholders Equity: ${getVal('Shareholders Equity')}<br>
        Book Value per Share: ${safeDiv(getVal('Shareholders Equity'), getVal('Number of Outstanding Shares')).toFixed(2)}<br>
        Earnings Per Share (EPS): ${getVal('Earnings Per Share (EPS)').toFixed(2)}<br>
        Price Earnings (P/E) Ratio: ${safeDiv(getVal('Market Value per Share'), getVal('Earnings Per Share (EPS)')).toFixed(2)}
      `;
      break;
    case 'operative':
      resultsDiv.innerHTML = `
        Working Capital: ${getVal('Current Assets') - getVal('Current Liabilities')}<br>
        Inventory Turnover: ${safeDiv(getVal('COGS'), getVal('Average Inventory')).toFixed(2)}<br>
        Accounts Payable Turnover: ${safeDiv(getVal('Purchases of COGS'), getVal('Average Accounts Payable') || 1).toFixed(2)}<br>
        Assets Turnover Ratio: ${safeDiv(getVal('Net Sales'), getVal('Average Total Assets')).toFixed(2)}<br>
        Fixed Asset Turnover: ${safeDiv(getVal('Net Sales'), getVal('Average Net Fixed Assets')).toFixed(2)}<br>
        Working Capital Turnover: ${safeDiv(getVal('Net Sales'), getVal('Current Assets') - getVal('Current Liabilities')).toFixed(2)}<br>
        Turnover Ratio: ${safeDiv(getVal('Net Revenue'), getVal('Total Assets')).toFixed(2)}<br>
        Dividend Payout Ratio: ${safeDiv(getVal('Dividends'), getVal('Net Income')).toFixed(2)}
      `;
      break;
    case 'cashflow':
      resultsDiv.innerHTML = `
        Operating Cash Flow Ratio: ${safeDiv(getVal('Cash Flow From Operating Activities'), getVal('Current Liabilities')).toFixed(2)}<br>
        Free Cash Flow: ${getVal('Cash Flow From Operating Activities') - getVal('Capital Expenditures')}<br>
        Cash Flow to Debt Ratio: ${safeDiv(getVal('Cash Flow From Operating Activities'), getVal('Total Liabilities')).toFixed(2)}<br>
        Cash Flow Margin: ${safeDiv(getVal('Cash Flow From Operating Activities'), getVal('Net Sales')).toFixed(2)}<br>
        Cash Return on Assets (Cash ROA): ${safeDiv(getVal('Cash Flow From Operating Activities'), getVal('Total Assets')).toFixed(2)}<br>
        Debt Coverage Ratio: ${safeDiv(getVal('EBITDA'), getVal('Principal') + getVal('Interest')).toFixed(2)}<br>
        Cash Conversion Ratio: ${safeDiv(getVal('Cash Flow From Operating Activities'), getVal('Net Income')).toFixed(2)}
      `;
      break;
    case 'returns':
      resultsDiv.innerHTML = `
        Returns On Assets (ROA): ${safeDiv(getVal('Net Income'), getVal('Total Assets')).toFixed(2)}<br>
        Return On Equity (ROE): ${safeDiv(getVal('Net Income'), getVal('Shareholders Equity')).toFixed(2)}<br>
        ROI: ${safeDiv(getVal('Net Profit'), getVal('Investment')).toFixed(2)}<br>
        ROIC: ${safeDiv(
          getVal('EBIT') * (1 - getVal('Tax')/100),
          getVal('Long Term Debt') + getVal('Shareholders Equity') - getVal('Non-Operating Cash')
        ).toFixed(2)}<br>
        ROE (alternative): ${safeDiv(getVal('EBIT') - getVal('Interest') - getVal('Tax'), getVal('Shareholders Equity')).toFixed(2)}<br>
        ROCE: ${safeDiv(getVal('EBIT'), getVal('Long Term Debt') + getVal('Shareholders Equity')).toFixed(2)}<br>
        ROA (alternative): ${safeDiv(getVal('EBIT') - getVal('Interest') - getVal('Tax'), getVal('Total Assets')).toFixed(2)}
      `;
      break;
  }
}

// Helper functions
function safeDiv(a, b) {
  return b === 0 || isNaN(b) ? 0 : a / b;
}

function updateStatus(msg) {
  if (statusDiv) {
    statusDiv.innerText = msg;
  }
}
