// MIT License. Copyright (c) 2026 YAP WEI JUN.
import { parseLocalizedNumber } from '../src/values.js';

const quantityColumnsByExport = new WeakMap();

function quantityCellType(root, locale, element, text, context) {
    const cell = element.closest('td, th');
    const rowTable = cell?.closest('table.rowitem, table[id^="row"]');
    if (!rowTable || (!rowTable.classList.contains('rowitem') && !/^row\d+$/.test(rowTable.id))
        || parseLocalizedNumber(text, locale) === null) return;
    const bounds = cell.getBoundingClientRect();
    if (bounds.width <= 0 || bounds.height <= 0) return;
    // Separate legacy item tables share the header's horizontal column boundaries.
    // Match both edges so item identifiers and numeric-looking units stay text.
    let columns = quantityColumnsByExport.get(context);
    if (!columns) {
        columns = [];
        for (const header of root.querySelectorAll('td, th')) {
            if (!/^(qty\.?|quantity|數量|数量)$/i.test(header.innerText.trim())) continue;
            const headerTable = header.closest('table.rowheader, table.rowitem, table[id^="row"]');
            if (!headerTable || (!headerTable.matches('table.rowheader, table.rowitem')
                && headerTable.id !== 'rowheader' && !/^row\d+$/.test(headerTable.id))) continue;
            columns.push(header.getBoundingClientRect());
        }
        quantityColumnsByExport.set(context, columns);
    }
    for (const column of columns) {
        if (column.width > 0 && column.height > 0 && column.top <= bounds.top
            && Math.abs(column.left - bounds.left) <= 2 && Math.abs(column.right - bounds.right) <= 2) {
            return 'number';
        }
    }
}

// All legacy ERP selectors and compatibility defaults belong to this adapter.
export function installGlobe3(api, documentRoot = document) {
    const root = documentRoot.querySelector('.exp_to_excel_button_content');
    if (!root) return;
    const options = {
        root,
        currency: root.getAttribute('data-excel-currency') || 'USD',
        locale: root.getAttribute('data-excel-locale') || 'en-US',
        cellType: (element, text, context) => quantityCellType(root,
            root.getAttribute('data-excel-locale') || 'en-US', element, text, context),
        exclude: ['[data-xlsx-ignore]', 'table#mainapproval', 'table#ext_id', 'table#myMenu'],
        pageBreakSelector: '[data-xlsx-page-break], .pagebreak_bf_processed, .pagebreak_bf',
        // Keep editable cells while retaining the printform's physical text/image sizes.
        fontScale: 1,
        normalFont: { name: 'Arial', size: 9 },
        minColumnWidthPx: 1,
        imageWidthScale: 1,
        imagePlacement: 'precise',
        preserveSingleLine: true,
        preserveRichText: true,
        includeBackgrounds: false,
        pageSetup: {
            horizontalCentered: false,
            margins: { left: 0.25, right: 0.25, top: 0.4, bottom: 0.4, header: 0.2, footer: 0.2 }
        }
    };
    window.exportPrintFormToXlsx = () => api.exportXlsx(options);
    const button = documentRoot.querySelector('.html_to_excel_btn');
    if (button && !button.hasAttribute('data-xlsx-button')) {
        api.bindButton(button, options);
        button.addEventListener('xlsx:error', event => {
            const regions = event.detail?.conflicts ? ' | regions: ' + JSON.stringify(event.detail.conflicts) : '';
            console.error('XLSX export failed: ' + (event.detail?.message || 'Unknown export error') + regions);
            alert('Export failed. Check the export container, options and ExcelJS dependency.');
        });
    }
}
