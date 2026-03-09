import { saveAs } from 'file-saver';

function getBridgeBaseUrl() {
    const fromEnv = (import.meta.env.EXCEL_BRIDGE_URL || import.meta.env.VITE_EXCEL_BRIDGE_URL || '').trim();
    if (fromEnv) return fromEnv.replace(/\/+$/, '');
    return 'http://127.0.0.1:8787';
}

export async function openExcelViaBridge(data, options = {}) {
    const baseUrl = getBridgeBaseUrl();
    const apiKey = (import.meta.env.EXCEL_BRIDGE_KEY || import.meta.env.VITE_EXCEL_BRIDGE_KEY || '').trim();
    const requestBody = {
        data,
        autoPrint: Boolean(options.autoPrint),
        openAfterPrint: options.openAfterPrint !== false,
        returnFile: Boolean(options.returnFile) // Ask for explicit binary download vs host-open
    };

    try {
        const response = await fetch(`${baseUrl}/open-excel`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'ngrok-skip-browser-warning': 'true',
                ...(apiKey ? { 'X-Bridge-Key': apiKey } : {})
            },
            body: JSON.stringify(requestBody)
        });

        if (!response.ok) {
            let details = '';
            try {
                // Determine if it returned an error JSON
                const contentType = response.headers.get('content-type');
                if (contentType && contentType.includes('application/json')) {
                    const payload = await response.json();
                    details = payload?.error || '';
                } else {
                    details = await response.text();
                }
            } catch {
                // Ignore parse errors.
            }
            return {
                success: false,
                error: details || `Bridge HTTP ${response.status}`
            };
        }

        // Check if the response is an actual file stream downloaded
        const contentType = response.headers.get('content-type');
        if (contentType && contentType.includes('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')) {
            const blob = await response.blob();

            // Extract filename from header if available
            const disposition = response.headers.get('content-disposition');
            let filename = `Form137_Export_${Date.now()}.xlsx`;
            if (disposition && disposition.indexOf('attachment') !== -1) {
                const filenameRegex = /filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/;
                const matches = filenameRegex.exec(disposition);
                if (matches != null && matches[1]) {
                    filename = matches[1].replace(/['"]/g, '');
                }
            }

            saveAs(blob, filename);
            return {
                success: true,
                filePath: filename,
                printed: false,
                warning: ''
            };
        }

        const payload = await response.json();
        if (!payload || !payload.success) {
            return {
                success: false,
                error: payload?.error || 'Bridge rejected request.'
            };
        }

        return {
            success: true,
            filePath: payload.filePath || '',
            printed: Boolean(payload.printed),
            warning: payload.warning || ''
        };
    } catch (err) {
        return {
            success: false,
            error: err.message
        };
    }
}
export function getBridgeDownloadUrl(type) {
    const baseUrl = getBridgeBaseUrl();
    const map = {
        installer: '/download/installer',
        exe: '/download/exe',
        template: '/download/template'
    };
    const path = map[type];
    if (!path) return null;

    const url = new URL(`${baseUrl}${path}`);
    url.searchParams.set('ngrok-skip-browser-warning', '1');
    return url.toString();
}

export async function checkBridgeHealth() {
    const baseUrl = getBridgeBaseUrl();
    const apiKey = (import.meta.env.EXCEL_BRIDGE_KEY || import.meta.env.VITE_EXCEL_BRIDGE_KEY || '').trim();

    try {
        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), 2000); // 2s timeout

        const response = await fetch(`${baseUrl}/health`, {
            method: 'GET',
            headers: {
                'ngrok-skip-browser-warning': 'true',
                ...(apiKey ? { 'X-Bridge-Key': apiKey } : {})
            },
            signal: controller.signal
        });

        clearTimeout(timeoutId);
        if (!response.ok) return { success: false, status: response.status };
        const data = await response.json();
        return {
            success: data?.success || false,
            host: data?.host,
            port: data?.port,
            service: data?.service
        };
    } catch (err) {
        return { success: false, error: err.message };
    }
}
