async function downloadExcel() {
  const d = currentData;

  const wb = new ExcelJS.Workbook();

  // ═══════════════════════════════════════
  // Sheet 1: Financial
  // ═══════════════════════════════════════
  const ws = wb.addWorksheet("Financial", {
    views: [{ rightToLeft: false }] // 🔥 LEFT TO RIGHT
  });

  const numFmt = '#,##0;(#,##0);"-"';

  function setCell(r, c, val, style = {}) {
    const cell = ws.getCell(r, c);

    if (typeof val === "string" && val.startsWith("=")) {
      cell.value = { formula: val.substring(1) };
    } else {
      cell.value = val;
    }

    if (style.font) cell.font = style.font;
    if (style.fill) cell.fill = style.fill;
    if (style.border) cell.border = style.border;
    if (style.alignment) cell.alignment = style.alignment;
    if (style.numFmt) cell.numFmt = style.numFmt;

    return cell;
  }
  const borderStyle = {
    top:    { style: 'thin' },
    left:   { style: 'thin' },
    bottom: { style: 'thin' },
    right:  { style: 'thin' }
  };
  const fills = {
    subtotal: { type: "pattern", pattern: "solid", fgColor: { argb: "FFBDD7EE" } },
    total:    { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFE699" } },
    orange:   { type: "pattern", pattern: "solid", fgColor: { argb: "FFF7CAAC" } }
  };

  const redFont = { color: { argb: "000" } };
  const bold = { bold: true, underline: true  };

  let r = 1;

  function row(labelAr, labelEn, val1 = "", val2 = "", val3 = "", style = {}) {
    setCell(r, 1, labelAr, style);
    setCell(r, 2, labelEn, style);
    setCell(r, 3, val1, style);
    setCell(r, 4, val2, style);
    setCell(r, 5, val3, style);
    r++;
  }

  function emptyRow() {
    r++;
  }
  emptyRow(); 
  emptyRow(); 
  // ─── HEADER ───────────────────────────
  row("الفترة / Period", "", "", "", "", { font: bold });
  row("العملة / Currency", "", "جم / EGP", "جم / EGP", "جم / EGP", { font: bold });
  row("السنة / Year","" ,d.years?.[0] ?? null, d.years?.[1] ?? null, d.years?.[2] ?? null);


  emptyRow(); 
  emptyRow(); 
  row("", "","0","0","0",{ font: bold });
  row("", "","0","0","0",{ font: bold });
  emptyRow();

// ─── BALANCE SHEET ────────────────────
row("النقدية", "Cash", d.balance_sheet?.cash?.[0] ?? null, d.balance_sheet?.cash?.[1] ?? null, d.balance_sheet?.cash?.[2] ?? null);
row("المخزون", "Inventories", d.balance_sheet?.inventories?.[0] ?? null, d.balance_sheet?.inventories?.[1] ?? null, d.balance_sheet?.inventories?.[2] ?? null);
row("اوراق قبض ومدينون", "Accounts Receivable & Debtors", d.balance_sheet?.accounts_receivable_debtors?.[0] ?? null, d.balance_sheet?.accounts_receivable_debtors?.[1] ?? null, d.balance_sheet?.accounts_receivable_debtors?.[2] ?? null);
row("غطاءات نقدية وتأمينات", "Cash Collateral", d.balance_sheet?.cash_collateral?.[0] ?? null, d.balance_sheet?.cash_collateral?.[1] ?? null, d.balance_sheet?.cash_collateral?.[2] ?? null);
row("مصروفات مدفوعة مقدمة", "Prepaid Expenses", d.balance_sheet?.prepaid_expenses?.[0] ?? null, d.balance_sheet?.prepaid_expenses?.[1] ?? null, d.balance_sheet?.prepaid_expenses?.[2] ?? null);
row("أصول متداولة أخرى", "Others", d.balance_sheet?.others_current_assets?.[0] ?? null, d.balance_sheet?.others_current_assets?.[1] ?? null, d.balance_sheet?.others_current_assets?.[2] ?? null);



row("إجمالى الاصول المتداولة", "Total Current Assets",
  d.balance_sheet?.total_current_assets?.[0] ?? d.current_assets?.[0] ?? null,
  d.balance_sheet?.total_current_assets?.[1] ?? d.current_assets?.[1] ?? null,
  d.balance_sheet?.total_current_assets?.[2] ?? d.current_assets?.[2] ?? null,
  { font: redFont, fill: fills.subtotal, numFmt });



  
// Fixed Assets

// Fixed Assets
row("مبانى", "Buildings", d.balance_sheet?.buildings?.[0] ?? null, d.balance_sheet?.buildings?.[1] ?? null, d.balance_sheet?.buildings?.[2] ?? null);
row("الاصول الغير الملموسة", "Intangible assets", d.balance_sheet?.intangible_assets?.[0] ?? null, d.balance_sheet?.intangible_assets?.[1] ?? null, d.balance_sheet?.intangible_assets?.[2] ?? null);
row("مجمع الاهلاك", "Accumulated Depreciation", d.balance_sheet?.accumulated_depreciation?.[0] ?? null, d.balance_sheet?.accumulated_depreciation?.[1] ?? null, d.balance_sheet?.accumulated_depreciation?.[2] ?? null);
row("أراضي", "Lands", d.balance_sheet?.lands?.[0] ?? null, d.balance_sheet?.lands?.[1] ?? null, d.balance_sheet?.lands?.[2] ?? null);
row("الات و معدات", "Machinery & Equipment", d.balance_sheet?.machinery_equipment?.[0] ?? null, d.balance_sheet?.machinery_equipment?.[1] ?? null, d.balance_sheet?.machinery_equipment?.[2] ?? null);
row("سيارات", "Vehicles", d.balance_sheet?.vehicles?.[0] ?? null, d.balance_sheet?.vehicles?.[1] ?? null, d.balance_sheet?.vehicles?.[2] ?? null);
row("اثاث وتجهيزات", "Office Furniture", d.balance_sheet?.office_furniture?.[0] ?? null, d.balance_sheet?.office_furniture?.[1] ?? null, d.balance_sheet?.office_furniture?.[2] ?? null);
row("مشروعات تحت التنفيذ", "Properties Under Development", d.balance_sheet?.properties_under_development?.[0] ?? null, d.balance_sheet?.properties_under_development?.[1] ?? null, d.balance_sheet?.properties_under_development?.[2] ?? null);
row("أصول ثابتة أخرى", "Others", d.balance_sheet?.others_fixed_assets?.[0] ?? null, d.balance_sheet?.others_fixed_assets?.[1] ?? null, d.balance_sheet?.others_fixed_assets?.[2] ?? null);




row("إجمالى الاصول الثابتة", "Total Fixed Assets",
  d.balance_sheet?.total_fixed_assets?.[0] ?? null,
  d.balance_sheet?.total_fixed_assets?.[1] ?? null,
  d.balance_sheet?.total_fixed_assets?.[2] ?? null,
  { font: redFont, fill: fills.subtotal, numFmt });

row("إجمالى الأصول", "Total Assets",
  d.balance_sheet?.total_assets_bs?.[0] ?? d.total_assets?.[0] ?? null,
  d.balance_sheet?.total_assets_bs?.[1] ?? d.total_assets?.[1] ?? null,
  d.balance_sheet?.total_assets_bs?.[2] ?? d.total_assets?.[2] ?? null,
  { font: redFont, fill: fills.total, numFmt });


  emptyRow();

  // Liabilities
  row("الخصوم", "Liabilities", "", "", "", { font: bold });

  row("بنوك", "Bank Overdraft", d.balance_sheet?.bank_overdraft?.[0] ?? null, d.balance_sheet?.bank_overdraft?.[1] ?? null, d.balance_sheet?.bank_overdraft?.[2] ?? null);
row("اوراق دفع ودائنون قصيرة الأجل", "Accounts Payable", d.balance_sheet?.accounts_payable?.[0] ?? null, d.balance_sheet?.accounts_payable?.[1] ?? null, d.balance_sheet?.accounts_payable?.[2] ?? null);
row("اقساط القروض مستحقة خلال عام", "Current Portion of Loan", d.balance_sheet?.current_portion_term_loan?.[0] ?? null, d.balance_sheet?.current_portion_term_loan?.[1] ?? null, d.balance_sheet?.current_portion_term_loan?.[2] ?? null);
row("مصروفات مستحقة", "Accrued Expenses", d.balance_sheet?.accrued_expenses?.[0] ?? null, d.balance_sheet?.accrued_expenses?.[1] ?? null, d.balance_sheet?.accrued_expenses?.[2] ?? null);
row("دفعات مقدمة", "Advance Payments", d.balance_sheet?.advance_payments?.[0] ?? null, d.balance_sheet?.advance_payments?.[1] ?? null, d.balance_sheet?.advance_payments?.[2] ?? null);
row("خصوم متداولة أخرى", "Others", d.balance_sheet?.others_current_liabilities?.[0] ?? null, d.balance_sheet?.others_current_liabilities?.[1] ?? null, d.balance_sheet?.others_current_liabilities?.[2] ?? null);

row("إجمالى الخصوم المتداولة", "Total Current Liabilities",
  d.balance_sheet?.total_current_liabilities?.[0] ?? d.current_liabilities?.[0] ?? null,
  d.balance_sheet?.total_current_liabilities?.[1] ?? d.current_liabilities?.[1] ?? null,
  d.balance_sheet?.total_current_liabilities?.[2] ?? d.current_liabilities?.[2] ?? null,
  { font: redFont, fill: fills.subtotal, numFmt });


 row("قروض طويلة الاجل", "Long Term Loans", d.balance_sheet?.non_current_portion_term_loan?.[0] ?? null, d.balance_sheet?.non_current_portion_term_loan?.[1] ?? null, d.balance_sheet?.non_current_portion_term_loan?.[2] ?? null);
row("اوراق دفع طويلة الاجل", "Notes Payable", d.balance_sheet?.notes_payables?.[0] ?? null, d.balance_sheet?.notes_payables?.[1] ?? null, d.balance_sheet?.notes_payables?.[2] ?? null);
row("قروض الشركاء", "Partners Loans", d.balance_sheet?.partners_loans?.[0] ?? null, d.balance_sheet?.partners_loans?.[1] ?? null, d.balance_sheet?.partners_loans?.[2] ?? null);
row("خصوم أخرى طويلة الاجل", "Others", d.balance_sheet?.others_long_term?.[0] ?? null, d.balance_sheet?.others_long_term?.[1] ?? null, d.balance_sheet?.others_long_term?.[2] ?? null);

row("إجمالى الخصوم طويلة الاجل", "Total Long Term Liabilities",
  d.balance_sheet?.total_long_term_liabilities?.[0] ?? null,
  d.balance_sheet?.total_long_term_liabilities?.[1] ?? null,
  d.balance_sheet?.total_long_term_liabilities?.[2] ?? null,
  { font: redFont, fill: fills.subtotal, numFmt });

  // Equity
  row("حقوق الملكية", "Equity", "", "", "", { font: bold });

 row("جارى الشركاء / المساهمين", "Shareholders Current Account", d.balance_sheet?.shareholders_current_account?.[0] ?? null, d.balance_sheet?.shareholders_current_account?.[1] ?? null, d.balance_sheet?.shareholders_current_account?.[2] ?? null);
row("ارباح/خسائر العام", "Net Profit", d.balance_sheet?.current_year_net_profit_loss?.[0] ?? d.net_profit?.[0] ?? null, d.balance_sheet?.current_year_net_profit_loss?.[1] ?? d.net_profit?.[1] ?? null, d.balance_sheet?.current_year_net_profit_loss?.[2] ?? d.net_profit?.[2] ?? null);
row("ارباح/خسائر اعوام سابقة", "Retained Earnings", d.balance_sheet?.retained_earnings?.[0] ?? null, d.balance_sheet?.retained_earnings?.[1] ?? null, d.balance_sheet?.retained_earnings?.[2] ?? null);
row("احتياطيات", "Reserves", d.balance_sheet?.reserves?.[0] ?? null, d.balance_sheet?.reserves?.[1] ?? null, d.balance_sheet?.reserves?.[2] ?? null);
row("احتياطى قانوني", "Legal Reserve", d.balance_sheet?.legal_reserve?.[0] ?? null, d.balance_sheet?.legal_reserve?.[1] ?? null, d.balance_sheet?.legal_reserve?.[2] ?? null);
row("حقوق ملكية أخرى", "Other Equity", d.balance_sheet?.other_equity?.[0] ?? null, d.balance_sheet?.other_equity?.[1] ?? null, d.balance_sheet?.other_equity?.[2] ?? null);
row("راس المال المدفوع", "Paid-up Capital", d.balance_sheet?.paid_up_capital?.[0] ?? null, d.balance_sheet?.paid_up_capital?.[1] ?? null, d.balance_sheet?.paid_up_capital?.[2] ?? null);


 row("إجمالى حقوق الملكية", "Total Equity",
  d.balance_sheet?.total_equity_bs?.[0] ?? d.total_equity?.[0] ?? null,
  d.balance_sheet?.total_equity_bs?.[1] ?? d.total_equity?.[1] ?? null,
  d.balance_sheet?.total_equity_bs?.[2] ?? d.total_equity?.[2] ?? null,
  { font: redFont, fill: fills.total, numFmt });

row("إجمالى الخصوم وحقوق الملكية", "Total Liabilities & Equity",
  d.balance_sheet?.total_liabilities_equity?.[0] ?? ((d.total_liabilities?.[0] ?? 0) + (d.total_equity?.[0] ?? 0)),
  d.balance_sheet?.total_liabilities_equity?.[1] ?? ((d.total_liabilities?.[1] ?? 0) + (d.total_equity?.[1] ?? 0)),
  d.balance_sheet?.total_liabilities_equity?.[2] ?? ((d.total_liabilities?.[2] ?? 0) + (d.total_equity?.[2] ?? 0)),
  { font: redFont, fill: fills.total, numFmt });

row("الفرق", "Difference",
  d.balance_sheet?.difference?.[0] ?? null,
  d.balance_sheet?.difference?.[1] ?? null,
  d.balance_sheet?.difference?.[2] ?? null,
  { font: redFont });

  emptyRow();
   emptyRow();

   emptyRow();
   emptyRow();
   emptyRow();


  // ─── INCOME STATEMENT ─────────────────
 row("الفترة", "Period",
  d.income_statement?.period?.[0] ?? "",
  d.income_statement?.period?.[1] ?? "",
  d.income_statement?.period?.[2] ?? "",
  { font: bold });

row("السنة", "Year",
  d.income_statement?.year?.[0] ?? d.years?.[0] ?? "",
  d.income_statement?.year?.[1] ?? d.years?.[1] ?? "",
  d.income_statement?.year?.[2] ?? d.years?.[2] ?? "",
  { font: bold });

row("ايرادات النشاط", "Total Sales (Revenues)",
  d.income_statement?.total_sales_revenues?.[0] ?? d.sales?.[0] ?? null,
  d.income_statement?.total_sales_revenues?.[1] ?? d.sales?.[1] ?? null,
  d.income_statement?.total_sales_revenues?.[2] ?? d.sales?.[2] ?? null,
  { numFmt });

row(" مصاريف النشاط", "Cast of Good Sold (COGS)",
  d.income_statement?.cogs?.[0] ?? ((d.sales?.[0] ?? null) !== null && (d.gross_profit?.[0] ?? null) !== null ? d.sales?.[0] - d.gross_profit?.[0] : null),
  d.income_statement?.cogs?.[1] ?? ((d.sales?.[1] ?? null) !== null && (d.gross_profit?.[1] ?? null) !== null ? d.sales?.[1] - d.gross_profit?.[1] : null),
  d.income_statement?.cogs?.[2] ?? ((d.sales?.[2] ?? null) !== null && (d.gross_profit?.[2] ?? null) !== null ? d.sales?.[2] - d.gross_profit?.[2] : null),
  { numFmt });

row("مجمل الربح", "Gross Profit",
  d.income_statement?.gross_profit_is?.[0] ?? d.gross_profit?.[0] ?? null,
  d.income_statement?.gross_profit_is?.[1] ?? d.gross_profit?.[1] ?? null,
  d.income_statement?.gross_profit_is?.[2] ?? d.gross_profit?.[2] ?? null,
  { font: redFont, fill: fills.orange, numFmt });

row("مصاريف ادارية وعمومية", "Sales , Administrative & General Expenses",
  d.income_statement?.sga_expenses?.[0] ?? null,
  d.income_statement?.sga_expenses?.[1] ?? null,
  d.income_statement?.sga_expenses?.[2] ?? null);

row("اهلاك صناعى", " Depreciation Industrial ",
  d.income_statement?.depreciation_industrial?.[0] ?? null,
  d.income_statement?.depreciation_industrial?.[1] ?? null,
  d.income_statement?.depreciation_industrial?.[2] ?? null);

row("اهلاك ادارى ", " Depreciation Adminstration",
  d.income_statement?.depreciation_administration?.[0] ?? null,
  d.income_statement?.depreciation_administration?.[1] ?? null,
  d.income_statement?.depreciation_administration?.[2] ?? null);

row("الفوائد الدائنة", " Interest revenue",
  d.income_statement?.interest_revenue?.[0] ?? null,
  d.income_statement?.interest_revenue?.[1] ?? null,
  d.income_statement?.interest_revenue?.[2] ?? null);

row("فوائد القروض", "Intrest expenses",
  d.income_statement?.interest_expenses?.[0] ?? null,
  d.income_statement?.interest_expenses?.[1] ?? null,
  d.income_statement?.interest_expenses?.[2] ?? null);

row("مصروفات أخرى", "Other Expenses",
  d.income_statement?.other_expenses?.[0] ?? null,
  d.income_statement?.other_expenses?.[1] ?? null,
  d.income_statement?.other_expenses?.[2] ?? null);

row("الضرائب", "Taxes",
  d.income_statement?.taxes?.[0] ?? null,
  d.income_statement?.taxes?.[1] ?? null,
  d.income_statement?.taxes?.[2] ?? null);

row("ايرادات أخرى", "Other Revenues",
  d.income_statement?.other_revenues?.[0] ?? null,
  d.income_statement?.other_revenues?.[1] ?? null,
  d.income_statement?.other_revenues?.[2] ?? null);

row("المخصصات", "Provisions",
  d.income_statement?.provisions?.[0] ?? null,
  d.income_statement?.provisions?.[1] ?? null,
  d.income_statement?.provisions?.[2] ?? null);

row("اجمالي المصروفات", "Total Expenses",
  d.income_statement?.total_expenses?.[0] ?? ((d.gross_profit?.[0] ?? null) !== null && (d.net_profit?.[0] ?? null) !== null ? d.gross_profit?.[0] - d.net_profit?.[0] : null),
  d.income_statement?.total_expenses?.[1] ?? ((d.gross_profit?.[1] ?? null) !== null && (d.net_profit?.[1] ?? null) !== null ? d.gross_profit?.[1] - d.net_profit?.[1] : null),
  d.income_statement?.total_expenses?.[2] ?? ((d.gross_profit?.[2] ?? null) !== null && (d.net_profit?.[2] ?? null) !== null ? d.gross_profit?.[2] - d.net_profit?.[2] : null),
  { fill: fills.orange, numFmt });

row("صافى الربح", "Net Profit",
  d.income_statement?.net_profit_loss_is?.[0] ?? d.net_profit?.[0] ?? null,
  d.income_statement?.net_profit_loss_is?.[1] ?? d.net_profit?.[1] ?? null,
  d.income_statement?.net_profit_loss_is?.[2] ?? d.net_profit?.[2] ?? null,
  { font: redFont, fill: fills.total, numFmt });

// ─── Main Indicators ─────────────────
row("ROE", "ROE",
  d.roe?.[0] ?? null,
  d.roe?.[1] ?? null,
  d.roe?.[2] ?? null,
  { font: redFont, fill: fills.total, numFmt });

  row("ROA", "ROA",
  d.roa?.[0] ?? null,
  d.roa?.[1] ?? null,
  d.roa?.[2] ?? null,
  { font: redFont, fill: fills.total, numFmt });

  row("Leverage", "Leverage",
  d.Leverage?.[0] ?? null,
  d.Leverage?.[1] ?? null,
  d.Leverage?.[2] ?? null,
  { font: redFont, fill: fills.total, numFmt });

  row("Operating Cash Flow", "Operating Cash Flow",
  d.cash_flow_from_operations?.[0] ?? null,
  d.cash_flow_from_operations?.[1] ?? null,
  d.cash_flow_from_operations?.[2] ?? null,
  { font: redFont, fill: fills.total, numFmt });

  // Width
  ws.columns = [
    { width: 35 },
    { width: 35 },
    { width: 20 },
    { width: 20 },
    { width: 20 }
  ];

  // ═══════════════════════════════════════
  // Sheet 2: Main Indicators
  // ═══════════════════════════════════════
  const ws2 = wb.addWorksheet("Main Indicators", {
    views: [{ rightToLeft: false }]
  });

  let r2 = 1;

  function row2(a, b, c) {
    ws2.getCell(r2, 1).value = a;
    ws2.getCell(r2, 2).value = b;
    ws2.getCell(r2, 3).value = c;
    r2++;
  }

  row2("Main Indicators", "", d.financial_year);
  row2("", "", "");
  row2("ROE", "Return On Equity", d.roe);
  row2("ROA", "Return On Assets", d.roa);
  row2("Leverage", "Liabilities / Equity", d.Leverage);
  row2("Operating Cash Flow", "Cash Flow", d.cash_flow_from_operations);

  row2("Current Ratio", "",
    d.current_liabilities ? (d.current_assets / d.current_liabilities).toFixed(2) : "");

  row2("Gross Margin", "",
    d.sales ? ((d.gross_profit / d.sales) * 100).toFixed(1) + "%" : "");

  row2("Net Margin", "",
    d.sales ? ((d.net_profit / d.sales) * 100).toFixed(1) + "%" : "");

  ws2.columns = [
    { width: 35 },
    { width: 40 },
    { width: 20 }
  ];
  // Apply border to all cells
ws.eachRow((row) => {
  row.eachCell({ includeEmpty: true }, (cell) => {
    cell.border = borderStyle;
  });
});

ws.eachRow((row) => {
  row.eachCell({ includeEmpty: true }, (cell) => {
    cell.alignment = {
      horizontal: 'center', // في النص عرضي
      vertical: 'middle'    // في النص طولي
    };
  });
});
  // ── Download ──
  const buffer = await wb.xlsx.writeBuffer();

  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });

  const url = window.URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `financial_${d.company_name || "company"}.xlsx`;
  a.click();

  window.URL.revokeObjectURL(url);
}