// MIT License. Copyright (c) 2026 YAP WEI JUN.
import { exportXlsx, bindButton, downloadBuffer, version } from './export.js';
import { installGlobe3 } from '../adapters/globe3.js';

window.HtmlToXlsx = Object.freeze({ exportXlsx, bindButton, downloadBuffer, version });
const initialize = () => {
    for (const button of document.querySelectorAll('[data-xlsx-button]')) {
        bindButton(button, { root: button.getAttribute('data-xlsx-target') || '[data-xlsx-export]' });
    }
    // The adapter acts only when legacy selectors exist, keeping generic use independent of ERP markup.
    installGlobe3(window.HtmlToXlsx);
};
if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', initialize, { once: true });
else initialize();
