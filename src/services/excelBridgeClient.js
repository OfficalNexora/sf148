function getBridgeBaseUrl() {
    const fromEnv = (import.meta.env.VITE_EXCEL_BRIDGE_URL || '').trim();
    if (fromEnv) return fromEnv.replace(/\/+$/, '');
    return 'http://127.0.0.1:8787';
}

export async function openExcelViaBridge(data) {
    const baseUrl = getBridgeBaseUrl();
    const apiKey = (import.meta.env.VITE_EXCEL_BRIDGE_KEY || '').trim();

    try {
        const response = await fetch(`${baseUrl}/open-excel`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                ...(apiKey ? { 'X-Bridge-Key': apiKey } : {})
            },
            body: JSON.stringify({ data })
        });

        if (!response.ok) {
            let details = '';
            try {
                const payload = await response.json();
                details = payload?.error || '';
            } catch {
                // Ignore parse errors.
            }
            return {
                success: false,
                error: details || `Bridge HTTP ${response.status}`
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
            filePath: payload.filePath || ''
        };
    } catch (err) {
        return {
            success: false,
            error: err.message
        };
    }
}
