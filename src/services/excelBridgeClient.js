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
            // We use file-saver to force download on the browser client
            const { saveAs } = await import('file-saver');

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
