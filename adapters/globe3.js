// MIT License. Copyright (c) 2026 YAP WEI JUN.
// All legacy ERP selectors and compatibility defaults belong to this adapter.
export function installGlobe3(api, documentRoot = document) {
    const root = documentRoot.querySelector('.exp_to_excel_button_content');
    if (!root) return;
    const options = {
        root,
        currency: root.getAttribute('data-excel-currency') || 'USD',
        locale: root.getAttribute('data-excel-locale') || 'en-US',
        exclude: ['[data-xlsx-ignore]', 'table#mainapproval', 'table#ext_id', 'table#myMenu'],
        pageBreakSelector: '.pagebreak_bf_processed'
    };
    window.exportPrintFormToXlsx = () => api.exportXlsx(options);
    const button = documentRoot.querySelector('.html_to_excel_btn');
    if (button && !button.hasAttribute('data-xlsx-button')) {
        api.bindButton(button, options);
        button.addEventListener('xlsx:error', () => {
            alert('Export failed. Check the export container, options and ExcelJS dependency.');
        });
    }
}
