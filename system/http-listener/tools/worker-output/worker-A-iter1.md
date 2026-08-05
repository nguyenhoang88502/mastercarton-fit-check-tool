```html
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>WVN Master Carton Fit Tool</title>
    <style>
        :root {
            --bg: #121212;
            --surface: #1e1e1e;
            --surface2: #282828;
            --border: #3a3a3a;
            --text: #e0e0e0;
            --text-secondary: #a0a0a0;
            --text-muted: #707070;
            --accent: #f0a030;
            --accent-glow: #f0a03066;
            --success: #4caf50;
            --success-bg: #1b3a1b;
            --error: #e53935;
            --error-bg: #3a1b1b;
            --warning: #ff9800;
            --warning-bg: #3a2e1b;
            --info: #2196f3;
            --info-bg: #1b2a3a;
            --focus-ring: #f0a030;
            --font-mono: 'Consolas', 'Courier New', monospace;
            --radius: 6px;
            --radius-sm: 4px;
            --transition: 0.2s ease;
            --shadow: 0 2px 8px rgba(0, 0, 0, 0.4);
            --shadow-lg: 0 4px 24px rgba(0, 0, 0, 0.6);
        }

        *,
        *::before,
        *::after {
            box-sizing: border-box;
            margin: 0;
            padding: 0;
        }

        html {
            font-size: 15px;
            scroll-behavior: smooth;
        }
        body {
            font-family: 'Segoe UI', system-ui, -apple-system, sans-serif;
            background: var(--bg);
            color: var(--text);
            min-height: 100vh;
            line-height: 1.6;
            display: flex;
            flex-direction: column;
        }

        /* ââ HEADER ââââââââââââââââââââââââââââ */
        .app-header {
            background: var(--surface);
            border-bottom: 2px solid var(--accent);
            padding: 10px 16px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            flex-wrap: wrap;
            gap: 8px;
            position: sticky;
            top: 0;
            z-index: 100;
            box-shadow: var(--shadow);
        }
        .app-title {
            font-size: 1.15rem;
            font-weight: 700;
            letter-spacing: 0.03em;
            color: var(--accent);
            white-space: nowrap;
            display: flex;
            align-items: center;
            gap: 8px;
        }
        .app-title .icon {
            font-size: 1.4rem;
        }
        .queue-badge {
            display: inline-flex;
            align-items: center;
            gap: 4px;
            background: var(--warning-bg);
            color: var(--warning);
            border: 1px solid var(--warning);
            border-radius: 20px;
            padding: 4px 12px;
            font-size: 0.8rem;
            font-weight: 600;
            cursor: pointer;
            transition: var(--transition);
            white-space: nowrap;
        }
        .queue-badge:empty,
        .queue-badge[data-count="0"] {
            display: none;
        }
        .queue-badge:hover {
            background: var(--warning);
            color: #121212;
        }

        /* ââ TAB NAVIGATION ââââââââââââââââââââ */
        .tab-nav-wrapper {
            background: var(--surface2);
            border-bottom: 1px solid var(--border);
            overflow-x: auto;
            -webkit-overflow-scrolling: touch;
            scrollbar-width: thin;
            scrollbar-color: var(--border) transparent;
            flex-shrink: 0;
        }
        .tab-nav-wrapper::-webkit-scrollbar {
            height: 4px;
        }
        .tab-nav-wrapper::-webkit-scrollbar-thumb {
            background: var(--border);
            border-radius: 4px;
        }
        .tab-nav {
            display: flex;
            min-width: max-content;
            padding: 0 8px;
            gap: 2px;
        }
        .tab-btn {
            flex: 0 0 auto;
            padding: 11px 18px;
            background: transparent;
            border: none;
            color: var(--text-secondary);
            cursor: pointer;
            font-size: 0.9rem;
            font-weight: 500;
            border-bottom: 3px solid transparent;
            transition: var(--transition);
            white-space: nowrap;
            outline-offset: -2px;
            border-radius: var(--radius-sm) var(--radius-sm) 0 0;
            position: relative;
            top: 1px;
        }
        .tab-btn:hover {
            color: var(--text);
            background: rgba(255, 255, 255, 0.03);
        }
        .tab-btn:focus-visible {
            outline: 2px solid var(--focus-ring);
            outline-offset: 2px;
            border-radius: var(--radius-sm);
        }
        .tab-btn.active {
            color: var(--accent);
            border-bottom-color: var(--accent);
            background: var(--bg);
            font-weight: 600;
        }

        /* ââ MAIN CONTENT ââââââââââââââââââââââ */
        .main-content {
            flex: 1;
            padding: 16px;
            max-width: 960px;
            width: 100%;
            margin: 0 auto;
            display: flex;
            flex-direction: column;
            gap: 14px;
        }
        .tab-panel {
            display: none;
            flex-direction: column;
            gap: 14px;
            animation: fadeIn 0.25s ease;
        }
        .tab-panel.active {
            display: flex;
        }
        @keyframes fadeIn {
            from {
                opacity: 0;
                transform: translateY(6px);
            }
            to {
                opacity: 1;
                transform: translateY(0);
            }
        }

        /* ââ CARDS & FORMS âââââââââââââââââââââ */
        .card {
            background: var(--surface);
            border: 1px solid var(--border);
            border-radius: var(--radius);
            padding: 18px 16px;
            box-shadow: var(--shadow);
        }
        .card h3 {
            font-size: 1rem;
            font-weight: 600;
            margin-bottom: 14px;
            color: var(--text);
            letter-spacing: 0.02em;
            display: flex;
            align-items: center;
            gap: 8px;
        }
        .card h3 .badge {
            font-size: 0.7rem;
            background: var(--accent);
            color: #121212;
            padding: 2px 8px;
            border-radius: 10px;
            font-weight: 700;
        }
        .form-grid {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 12px;
        }
        .form-grid.single-col {
            grid-template-columns: 1fr;
        }
        .form-group {
            display: flex;
            flex-direction: column;
            gap: 4px;
        }
        .form-group.full-width {
            grid-column: 1 / -1;
        }
        .form-group label {
            font-size: 0.82rem;
            font-weight: 600;
            color: var(--text-secondary);
            letter-spacing: 0.02em;
        }
        .form-group label .required-star {
            color: var(--error);
        }
        .form-group input,
        .form-group select,
        .form-group textarea {
            background: var(--surface2);
            border: 1px solid var(--border);
            border-radius: var(--radius-sm);
            padding: 9px 11px;
            color: var(--text);
            font-size: 0.9rem;
            font-family: inherit;
            transition: var(--transition);
            width: 100%;
            min-width: 0;
        }
        .form-group input:focus,
        .form-group select:focus,
        .form-group textarea:focus {
            outline: none;
            border-color: var(--accent);
            box-shadow: 0 0 0 3px var(--accent-glow);
        }
        .form-group input.invalid,
        .form-group select.invalid,
        .form-group textarea.invalid {
            border-color: var(--error);
            box-shadow: 0 0 0 3px rgba(229, 57, 53, 0.25);
            background: #2a1515;
        }
        .form-group input.valid,
        .form-group select.valid,
        .form-group textarea.valid {
            border-color: var(--success);
            box-shadow: 0 0 0 3px rgba(76, 175, 80, 0.2);
        }
        .form-group .hint {
            font-size: 0.72rem;
            color: var(--text-muted);
            min-height: 1.1em;
            line-height: 1.3;
        }
        .form-group .error-msg {
            font-size: 0.75rem;
            color: var(--error);
            font-weight: 500;
            min-height: 1.2em;
            line-height: 1.3;
        }
        .btn {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            gap: 8px;
            padding: 10px 20px;
            border: none;
            border-radius: var(--radius-sm);
            cursor: pointer;
            font-size: 0.9rem;
            font-weight: 600;
            letter-spacing: 0.02em;
            transition: var(--transition);
            white-space: nowrap;
            min-width: 100px;
        }
        .btn:focus-visible {
            outline: 2px solid var(--focus-ring);
            outline-offset: 3px;
        }
        .btn-primary {
            background: var(--accent);
            color: #121212;
        }
        .btn-primary:hover:not(:disabled) {
            background: #f5b840;
            box-shadow: 0 0 12px var(--accent-glow);
        }
        .btn-primary:disabled {
            opacity: 0.5;
            cursor: not-allowed;
            filter: grayscale(30%);
        }
        .btn-secondary {
            background: var(--surface2);
            color: var(--text);
            border: 1px solid var(--border);
        }
        .btn-secondary:hover:not(:disabled) {
            background: #383838;
            border-color: #555;
        }
        .btn-sm {
            padding: 6px 12px;
            font-size: 0.8rem;
            min-width: auto;
        }

        /* ââ NOTIFICATION AREA âââââââââââââââââ */
        .notification-area {
            min-height: 20px;
            display: flex;
            flex-direction: column;
            gap: 8px;
        }
        .notif {
            padding: 10px 14px;
            border-radius: var(--radius-sm);
            font-size: 0.85rem;
            font-weight: 500;
            line-height: 1.5;
            animation: slideIn 0.3s ease;
            word-break: break-word;
        }
        @keyframes slideIn {
            from {
                opacity: 0;
                transform: translateY(-8px);
            }
            to {
                opacity: 1;
                transform: translateY(0);
            }
        }
        .notif.success {
            background: var(--success-bg);
            border: 1px solid var(--success);
            color: #81c784;
        }
        .notif.error {
            background: var(--error-bg);
            border: 1px solid var(--error);
            color: #ef9a9a;
        }
        .notif.warning {
            background: var(--warning-bg);
            border: 1px solid var(--warning);
            color: #ffcc80;
        }
        .notif.info {
            background: var(--info-bg);
            border: 1px solid var(--info);
            color: #90caf9;
        }
        .notif ul {
            margin: 6px 0 0 18px;
            padding: 0;
        }
        .notif li {
            margin-bottom: 2px;
        }

        /* ââ RAW RESPONSE DETAILS ââââââââââââââ */
        .raw-response {
            margin-top: 6px;
        }
        .raw-response>summary {
            cursor: pointer;
            font-size: 0.78rem;
            color: var(--text-muted);
            font-weight: 500;
            user-select: none;
            padding: 4px 0;
            transition: var(--transition);
        }
        .raw-response>summary:hover {
            color: var(--text-secondary);
        }
        .raw-response pre {
            background: #0a0a0a;
            border: 1px solid var(--border);
            border-radius: var(--radius-sm);
            padding: 12px;
            font-family: var(--font-mono);
            font-size: 0.75rem;
            color: #b0b0b0;
            overflow-x: auto;
            max-height: 300px;
            overflow-y: auto;
            line-height: 1.5;
            white-space: pre-wrap;
            word-break: break-all;
        }

        /* ââ SPINNER OVERLAY âââââââââââââââââââ */
        .spinner-overlay {
            display: none;
            position: fixed;
            inset: 0;
            background: rgba(0, 0, 0, 0.55);
            z-index: 999;
            align-items: center;
            justify-content: center;
            flex-direction: column;
            gap: 14px;
        }
        .spinner-overlay.active {
            display: flex;
        }
        .spinner {
            width: 44px;
            height: 44px;
            border: 3px solid rgba(255, 255, 255, 0.15);
            border-top-color: var(--accent);
            border-radius: 50%;
            animation: spin 0.75s linear infinite;
        }
        @keyframes spin {
            to {
                transform: rotate(360deg);
            }
        }
        .spinner-label {
            color: var(--text);
            font-size: 0.85rem;
            font-weight: 500;
        }

        /* ââ TABLE âââââââââââââââââââââââââââââ */
        .table-container {
            overflow-x: auto;
            max-height: 420px;
            overflow-y: auto;
            border: 1px solid var(--border);
            border-radius: var(--radius-sm);
            scrollbar-width: thin;
            scrollbar-color: var(--border) var(--surface2);
        }
        .table-container::-webkit-scrollbar {
            width: 6px;
            height: 6px;
        }
        .table-container::-webkit-scrollbar-thumb {
            background: var(--border);
            border-radius: 4px;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            font-size: 0.8rem;
            min-width: 500px;
        }
        table thead {
            position: sticky;
            top: 0;
            z-index: 2;
        }
        table th {
            background: var(--surface2);
            color: var(--text-secondary);
            font-weight: 600;
            padding: 9px 10px;
            text-align: left;
            border-bottom: 2px solid var(--border);
            white-space: nowrap;
            font-size: 0.75rem;
            letter-spacing: 0.03em;
            text-transform: uppercase;
        }
        table td {
            padding: 8px 10px;
            border-bottom: 1px solid var(--border);
            white-space: nowrap;
            color: var(--text);
        }
        table tbody tr:hover {
            background: rgba(255, 255, 255, 0.025);
        }
        table tbody tr.row-success {
            background: rgba(76, 175, 80, 0.08);
        }
        table tbody tr.row-error {
            background: rgba(229, 57, 53, 0.08);
        }
        .empty-state {
            text-align: center;
            padding: 32px 16px;
            color: var(--text-muted);
            font-size: 0.9rem;
        }
        .filter-row {
            display: flex;
            gap: 8px;
            align-items: center;
            flex-wrap: wrap;
        }
        .filter-row input {
            flex: 1;
            min-width: 150px;
            background: var(--surface2);
            border: 1px solid var(--border);
            border-radius: var(--radius-sm);
            padding: 8px 10px;
            color: var(--text);
            font-size: 0.85rem;
        }
        .filter-row input:focus {
            outline: none;
            border-color: var(--accent);
            box-shadow: 0 0 0 3px var(--accent-glow);
        }

        /* ââ BATCH TEXTAREA ââââââââââââââââââââ */
        .batch-textarea {
            width: 100%;
            min-height: 180px;
            font-family: var(--font-mono);
            font-size: 0.8rem;
            resize: vertical;
            background: var(--surface2);
            border: 1px solid var(--border);
            border-radius: var(--radius-sm);
            padding: 10px;
            color: var(--text);
            line-height: 1.6;
            tab-size: 2;
        }
        .batch-textarea:focus {
            outline: none;
            border-color: var(--accent);
            box-shadow: 0 0 0 3px var(--accent-glow);
        }
        .batch-options {
            display: flex;
            gap: 12px;
            align-items: center;
            flex-wrap: wrap;
        }
        .checkbox-label {
            display: flex;
            align-items: center;
            gap: 6px;
            cursor: pointer;
            font-size: 0.85rem;
            user-select: none;
            color: var(--text-secondary);
        }
        .checkbox-label input[type="checkbox"] {
            accent-color: var(--accent);
            width: 16px;
            height: 16px;
            cursor: pointer;
        }

        /* ââ RESPONSIVE ââââââââââââââââââââââââ */
        @media (max-width: 600px) {
            .form-grid {
                grid-template-columns: 1fr;
            }
            .form-grid.single-col {
                grid-template-columns: 1fr;
            }
            .app-header {
                padding: 8px 10px;
                gap: 4px;
            }
            .app-title {
                font-size: 1rem;
            }
            .tab-btn {
                padding: 9px 11px;
                font-size: 0.78rem;
            }
            .main-content {
                padding: 10px;
                gap: 10px;
            }
            .card {
                padding: 12px 10px;
            }
            .btn {
                padding: 8px 14px;
                font-size: 0.82rem;
                min-width: 80px;
            }
            table {
                font-size: 0.7rem;
            }
            table th,
            table td {
                padding: 6px 7px;
            }
            .batch-options {
                flex-direction: column;
                align-items: flex-start;
                gap: 6px;
            }
        }
        @media (max-width: 360px) {
            html {
                font-size: 13px;
            }
            .tab-btn {
                padding: 7px 8px;
                font-size: 0.7rem;
            }
            .btn {
                width: 100%;
            }
        }

        /* ââ ACCESSIBILITY âââââââââââââââââââââ */
        .sr-only {
            position: absolute;
            width: 1px;
            height: 1px;
            padding: 0;
            margin: -1px;
            overflow: hidden;
            clip: rect(0, 0, 0, 0);
            white-space: nowrap;
            border: 0;
        }
        *:focus-visible {
            outline: 2px solid var(--focus-ring);
            outline-offset: 2px;
        }
    </style>
</head>
<body>
    <!-- HEADER -->
    <header class="app-header">
        <div class="app-title">
            <span class="icon">&#9881;</span> WVN Master Carton Fit Tool
        </div>
        <span class="queue-badge" id="queueBadge" data-count="0" title="Pending offline items" role="status" aria-live="polite"></span>
    </header>

    <!-- TAB NAVIGATION -->
    <nav class="tab-nav-wrapper" aria-label="Main navigation tabs">
        <div class="tab-nav" role="tablist">
            <button class="tab-btn active" role="tab" data-tab="tab-fg" aria-selected="true" id="tab-btn-fg">1. New FG</button>
            <button class="tab-btn" role="tab" data-tab="tab-carton" aria-selected="false" id="tab-btn-carton">2. New Carton</button>
            <button class="tab-btn" role="tab" data-tab="tab-bom" aria-selected="false" id="tab-btn-bom">3. New BOM</button>
            <button class="tab-btn" role="tab" data-tab="tab-lookup" aria-selected="false" id="tab-btn-lookup">4. Lookup</button>
            <button class="tab-btn" role="tab" data-tab="tab-batch" aria-selected="false" id="tab-btn-batch">5. Batch</button>
        </div>
    </nav>

    <!-- MAIN CONTENT -->
    <main class="main-content" id="mainContent">
        <!-- TAB 1: New FG -->
        <section class="tab-panel active" id="tab-fg" role="tabpanel" aria-labelledby="tab-btn-fg">
            <div class="card">
                <h3><span class="badge">POST</span> New Finished Good Dimension</h3>
                <form id="formFg" novalidate autocomplete="off">
                    <div class="form-grid">
                        <div class="form-group">
                            <label for="fg_id">FG ID <span class="required-star">*</span></label>
                            <input type="text" id="fg_id" name="fg_id" required maxlength="50" pattern="[A-Za-z0-9_-]+" placeholder="e.g. FG-001234">
                            <span class="hint">Required. Max 50 chars. A-Z, a-z, 0-9, _, -</span>
                            <span class="error-msg" id="err-fg_id" aria-live="polite"></span>
                        </div>
                        <div class="form-group">
                            <label for="fg_length_mm">Length (mm) <span class="required-star">*</span></label>
                            <input type="number" id="fg_length_mm" name="length_mm" required min="10" max="1000" step="0.1" placeholder="150.5">
                            <span class="hint">Required. Range: 10â1000 mm</span>
                            <span class="error-msg" id="err-fg_length_mm" aria-live="polite"></span>
                        </div>
                        <div class="form-group">
                            <label for="fg_width_mm">Width (mm) <span class="required-star">*</span></label>
                            <input type="number" id="fg_width_mm" name="width_mm" required min="10" max="1000" step="0.1" placeholder="100.0">
                            <span class="hint">Required. Range: 10â1000 mm</span>
                            <span class="error-msg" id="err-fg_width_mm" aria-live="polite"></span>
                        </div>
                        <div class="form-group">
                            <label for="fg_height_mm">Height (mm) <span class="required-star">*</span></label>
                            <input type="number" id="fg_height_mm" name="height_mm" required min="10" max="1000" step="0.1" placeholder="80.0">
                            <span class="hint">Required. Range: 10â1000 mm</span>
                            <span class="error-msg" id="err-fg_height_mm" aria-live="polite"></span>
                        </div>
                        <div class="form-group">
                            <label for="fg_target_qty">Target Qty</label>
                            <input type="number" id="fg_target_qty" name="target_qty" min="1" max="10000" step="1" placeholder="48 (optional)">
                            <span class="hint">Optional. Integer, 1â10000</span>
                            <span class="error-msg" id="err-fg_target_qty" aria-live="polite"></span>
                        </div>
                        <div class="form-group">
                            <label for="fg_source">Source</label>
                            <select id="fg_source" name="source">
                                <option value="MANUAL_ENTRY" selected>MANUAL_ENTRY</option>
                                <option value="D365_IMPORT">D365_IMPORT</option>
                                <option value="MOBILE_APP">MOBILE_APP</option>
                                <option value="BATCH_IMPORT">BATCH_IMPORT</option>
                            </select>
                            <span class="hint">Default: MANUAL_ENTRY</span>
                            <span class="error-msg" id="err-fg_source" aria-live="polite"></span>
                        </div>
                    </div>
                    <div style="margin-top:14px;display:flex;gap:10px;flex-wrap:wrap;">
                        <button type="submit" class="btn btn-primary" id="btnFgSubmit">Submit FG</button>
                        <button type="button" class="btn btn-secondary btn-sm" id="btnFgClear">Clear</button>
                    </div>
                </form>
            </div>
            <div class="notification-area" id="notifFg" aria-live="polite"></div>
        </section>

        <!-- TAB 2: New Carton -->
        <section class="tab-panel" id="tab-carton" role="tabpanel" aria-labelledby="tab-btn-carton">
            <div class="card">
                <h3><span class="badge">POST</span> New Carton Dimension</h3>
                <form id="formCarton" novalidate autocomplete="off">
                    <div class="form-grid">
                        <div class="form-group">
                            <label for="carton_id">Carton ID <span class="required-star">*</span></label>
                            <input type="text" id="carton_id" name="carton_id" required maxlength="50" pattern="[A-Za-z0-9_-]+" placeholder="e.g. 1008609">
                            <span class="hint">Required. Max 50 chars. A-Z, a-z, 0-9, _, -</span>
                            <span class="error-msg" id="err-carton_id" aria-live="polite"></span>
                        </div>
                        <div class="form-group">
                            <label for="carton_length_mm">Length (mm) <span class="required-star">*</span></label>
                            <input type="number" id="carton_length_mm" name="length_mm" required min="10" max="1000" step="0.1" placeholder="200.0">
                            <span class="hint">Required. Range: 10â1000 mm</span>
                            <span class="error-msg" id="err-carton_length_mm" aria-live="polite"></span>
                        </div>
                        <div class="form-group">
                            <label for="carton_width_mm">Width (mm) <span class="required-star">*</span></label>
                            <input type="number" id="carton_width_mm" name="width_mm" required min="10" max="1000" step="0.1" placeholder="150.0">
                            <span class="hint">Required. Range: 10â1000 mm</span>
                            <span class="error-msg" id="err-carton_width_mm" aria-live="polite"></span>
                        </div>
                        <div class="form-group">
                            <label for="carton_height_mm">Height (mm) <span class="required-star">*</span></label>
                            <input type="number" id="carton_height_mm" name="height_mm" required min="10" max="1000" step="0.1" placeholder="120.0">
                            <span class="hint">Required. Range: 10â1000 mm</span>
                            <span class="error-msg" id="err-carton_height_mm" aria-live="polite"></span>
                        </div>
                        <div class="form-group full-width">
                            <label for="carton_dimension_source">Dimension Source</label>
                            <select id="carton_dimension_source" name="dimension_source">
                                <option value="BOM_LINKED_CA_AVERAGE" selected>BOM_LINKED_CA_AVERAGE</option>
                                <option value="MANUAL_MEASUREMENT">MANUAL_MEASUREMENT</option>
                                <option value="SUPPLIER_SPEC">SUPPLIER_SPEC</option>
                            </select>
                            <span class="hint">Select dimension source</span>
                            <span class="error-msg" id="err-carton_dimension_source" aria-live="polite"></span>
                        </div>
                    </div>
                    <div style="margin-top:14px;display:flex;gap:10px;flex-wrap:wrap;">
                        <button type="submit" class="btn btn-primary" id="btnCartonSubmit">Submit Carton</button>
                        <button type="button" class="btn btn-secondary btn-sm" id="btnCartonClear">Clear</button>
                    </div>
                </form>
            </div>
            <div class="notification-area" id="notifCarton" aria-live="polite"></div>
        </section>

        <!-- TAB 3: New BOM -->
        <section class="tab-panel" id="tab-bom" role="tabpanel" aria-labelledby="tab-btn-bom">
            <div class="card">
                <h3><span class="badge">POST</span> New BOM Assignment</h3>
                <form id="formBom" novalidate autocomplete="off">
                    <div class="form-grid single-col">
                        <div class="form-group">
                            <label for="bom_fg_id">FG ID <span class="required-star">*</span></label>
                            <input type="text" id="bom_fg_id" name="fg_id" required maxlength="50" pattern="[A-Za-z0-9_-]+" placeholder="e.g. FG-001234">
                            <span class="hint">Required. Max 50 chars. A-Z, a-z, 0-9, _, -</span>
                            <span class="error-msg" id="err-bom_fg_id" aria-live="polite"></span>
                        </div>
                        <div class="form-group">
                            <label for="bom_carton_id">Carton ID <span class="required-star">*</span></label>
                            <input type="text" id="bom_carton_id" name="carton_id" required maxlength="50" pattern="[A-Za-z0-9_-]+" placeholder="e.g. 1008609">
                            <span class="hint">Required. Max 50 chars. A-Z, a-z, 0-9, _, -</span>
                            <span class="error-msg" id="err-bom_carton_id" aria-live="polite"></span>
                        </div>
                        <div class="form-group">
                            <label for="bom_source_bom_date">Source BOM Date <span class="required-star">*</span></label>
                            <input type="date" id="bom_source_bom_date" name="source_bom_date" required placeholder="2026-08-04">
                            <span class="hint">Required. ISO format: YYYY-MM-DD</span>
                            <span class="error-msg" id="err-bom_source_bom_date" aria-live="polite"></span>
                        </div>
                        <div class="form-group">
                            <label for="bom_confidence">Confidence</label>
                            <select id="bom_confidence" name="confidence">
                                <option value="MEDIUM" selected>MEDIUM</option>
                                <option value="HIGH">HIGH</option>
                                <option value="LOW">LOW</option>
                            </select>
                            <span class="hint">Default: MEDIUM</span>
                            <span class="error-msg" id="err-bom_confidence" aria-live="polite"></span>
                        </div>
                    </div>
                    <div style="margin-top:14px;display:flex;gap:10px;flex-wrap:wrap;">
                        <button type="submit" class="btn btn-primary" id="btnBomSubmit">Submit BOM</button>
                        <button type="button" class="btn btn-secondary btn-sm" id="btnBomClear">Clear</button>
                    </div>
                </form>
            </div>
            <div class="notification-area" id="notifBom" aria-live="polite"></div>
        </section>

        <!-- TAB 4: Lookup -->
        <section class="tab-panel" id="tab-lookup" role="tabpanel" aria-labelledby="tab-btn-lookup">
            <div class="card">
                <h3>&#128269; Lookup Records</h3>
                <div class="filter-row" style="margin-bottom:10px;">
                    <select id="lookupDataset" aria-label="Select dataset to query">
                        <option value="fg-dimensions">FG Dimensions</option>
                        <option value="carton-dimensions">Carton Dimensions</option>
                        <option value="bom-assignments">BOM Assignments</option>
                    </select>
                    <button type="button" class="btn btn-primary btn-sm" id="btnLookupRefresh">Refresh</button>
                    <input type="text" id="lookupFilter" placeholder="Filter resultsâ¦" aria-label="Text filter for results">
                </div>
                <div class="table-container" id="lookupTableContainer">
                    <p class="empty-state">Select a dataset and click Refresh to load records.</p>
                </div>
                <div class="notification-area" id="notifLookup" aria-live="polite" style="margin-top:8px;"></div>
            </div>
        </section>

        <!-- TAB 5: Batch -->
        <section class="tab-panel" id="tab-batch" role="tabpanel" aria-labelledby="tab-btn-batch">
            <div class="card">
                <h3><span class="badge">BATCH</span> Bulk Import</h3>
                <div class="batch-options" style="margin-bottom:10px;">
                    <select id="batchDataset" aria-label="Target dataset for batch import">
                        <option value="fg-dimensions">FG Dimensions</option>
                        <option value="carton-dimensions">Carton Dimensions</option>
                        <option value="bom-assignments">BOM Assignments</option>
                    </select>
                    <label class="checkbox-label">
                        <input type="checkbox" id="batchValidateOnly">
                        Validate only (no server submit)
                    </label>
                </div>
                <textarea class="batch-textarea" id="batchTextarea" placeholder="Paste JSON array or CSV (with header row) hereâ¦&#10;&#10;JSON example:&#10;[{&quot;fg_id&quot;:&quot;FG-001&quot;,&quot;length_mm&quot;:150,&quot;width_mm&quot;:100,&quot;height_mm&quot;:80,&quot;target_qty&quot;:48,&quot;source&quot;:&quot;MANUAL_ENTRY&quot;}]&#10;&#10;CSV example:&#10;fg_id,length_mm,width_mm,height_mm,target_qty,source&#10;FG-001,150,100,80,48,MANUAL_ENTRY" aria-label="Batch data input"></textarea>
                <div style="margin-top:10px;display:flex;gap:10px;flex-wrap:wrap;">
                    <button type="button" class="btn btn-primary" id="btnBatchSubmit">Submit Batch</button>
                    <button type="button" class="btn btn-secondary btn-sm" id="btnBatchClear">Clear</button>
                </div>
                <div class="notification-area" id="notifBatch" aria-live="polite" style="margin-top:8px;"></div>
                <div class="table-container" id="batchResultsContainer" style="margin-top:10px;display:none;"></div>
            </div>
        </section>
    </main>

    <!-- SPINNER OVERLAY -->
    <div class="spinner-overlay" id="spinnerOverlay" aria-hidden="true">
        <div class="spinner"></div>
        <span class="spinner-label">Processingâ¦</span>
    </div>

    <script>
        (function() {
            'use strict';

            // âââ DOM REFERENCES ââââââââââââââââââââââââââââââ
            const $ = (sel) => document.querySelector(sel);
            const $$ = (sel) => document.querySelectorAll(sel);

            const spinnerOverlay = $('#spinnerOverlay');
            const queueBadge = $('#queueBadge');
            const mainContent = $('#mainContent');

            // Forms
            const formFg = $('#formFg');
            const formCarton = $('#formCarton');
            const formBom = $('#formBom');

            // Notification areas
            const notifFg = $('#notifFg');
            const notifCarton = $('#notifCarton');
            const notifBom = $('#notifBom');
            const notifLookup = $('#notifLookup');
            const notifBatch = $('#notifBatch');

            // âââ TAB MANAGEMENT âââââââââââââââââââââââââââââ
            const tabBtns = $$('.tab-btn');
            const tabPanels = $$('.tab-panel');

            function switchTab(tabId) {
                tabBtns.forEach(b => {
                    b.classList.remove('active');
                    b.setAttribute('aria-selected', 'false');
                });
                tabPanels.forEach(p => p.classList.remove('active'));
                const btn = document.querySelector(`[data-tab="${tabId}"]`);
                const panel = document.getElementById(tabId);
                if (btn) {
                    btn.classList.add('active');
                    btn.setAttribute('aria-selected', 'true');
                }
                if (panel) panel.classList.add('active');
            }

            tabBtns.forEach(btn => {
                btn.addEventListener('click', () => {
                    const tabId = btn.getAttribute('data-tab');
                    switchTab(tabId);
                });
                btn.addEventListener('keydown', (e) => {
                    if (e.key === 'Enter' || e.key === ' ') {
                        e.preventDefault();
                        btn.click();
                    }
                });
            });

            // âââ SPINNER ââââââââââââââââââââââââââââââââââââ
            function showSpinner() { spinnerOverlay.classList.add('active');
                spinnerOverlay.setAttribute('aria-hidden', 'false'); }

            function hideSpinner() { spinnerOverlay.classList.remove('active');
                spinnerOverlay.setAttribute('aria-hidden', 'true'); }

            // âââ NOTIFICATION HELPERS âââââââââââââââââââââââ
            function clearNotifications(container) { container.innerHTML = ''; }

            function showNotification(container, type, message, rawResponse) {
                const div = document.createElement('div');
                div.className = 'notif ' + type;
                if (typeof message === 'string') {
                    div.textContent = message;
                } else if (Array.isArray(message)) {
                    const ul = document.createElement('ul');
                    message.forEach(m => { const li = document.createElement('li');
                        li.textContent = m;
                        ul.appendChild(li); });
                    div.appendChild(document.createTextNode('Validation errors:'));
                    div.appendChild(ul);
                }
                container.appendChild(div);

                if (rawResponse !== undefined && rawResponse !== null) {
                    const details = document.createElement('details');
                    details.className = 'raw-response';
                    const summary = document.createElement('summary');
                    summary.textContent = 'Raw response';
                    details.appendChild(summary);
                    const pre = document.createElement('pre');
                    try {
                        pre.textContent = JSON.stringify(rawResponse, null, 2);
                    } catch (e) {
                        pre.textContent = String(rawResponse);
                    }
                    details.appendChild(pre);
                    container.appendChild(details);
                }
                // Auto-scroll to notification
                div.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
            }

            function showSuccess(container, msg, raw) { showNotification(container, 'success', msg,
                raw); }

            function showError(container, msg, raw) { showNotification(container, 'error', msg, raw); }

            function showWarning(container, msg, raw) { showNotification(container, 'warning', msg, raw); }

            function showInfo(container, msg, raw) { showNotification(container, 'info', msg, raw); }

            // âââ VALIDATION FUNCTIONS ââââââââââââââââââââââ
            const ID_REGEX = /^[A-Za-z0-9_-]+$/;
            const SOURCE_ENUM = ['MANUAL_ENTRY', 'D365_IMPORT', 'MOBILE_APP', 'BATCH_IMPORT'];
            const DIM_SOURCE_ENUM = ['BOM_LINKED_CA_AVERAGE', 'MANUAL_MEASUREMENT', 'SUPPLIER_SPEC'];
            const CONFIDENCE_ENUM = ['HIGH', 'MEDIUM', 'LOW'];

            function validateId(value, fieldName) {
                if (!value || value.trim() === '') return { valid: false, msg: `${fieldName} is required.` };
                if (value.length > 50) return { valid: false, msg: `${fieldName} must be 50 characters or fewer.` };
                if (!ID_REGEX.test(value)) return { valid: false,
                    msg: `${fieldName} contains invalid characters. Allowed: A-Z, a-z, 0-9, _, -.` };
                return { valid: true, msg: '' };
            }

            function validateDimension(value, fieldName) {
                if (value === '' || value === null || value === undefined) return { valid: false,
                    msg: `${fieldName} is required.` };
                const num = Number(value);
                if (isNaN(num)) return { valid: false, msg: `${fieldName} must be a number.` };
                if (num < 10 || num > 1000) return { valid: false,
                msg: `${fieldName} must be between 10 and 1000 mm.` };
                return { valid: true, msg: '' };
            }

            function validateTargetQty(value) {
                if (value === '' || value === null || value === undefined) return { valid: true, msg: '' }; // optional
                const num = Number(value);
                if (isNaN(num) || !Number.isInteger(num)) return { valid: false,
                    msg: 'Target Qty must be an integer.' };
                if (num < 1 || num > 10000) return { valid: false, msg: 'Target Qty must be between 1 and 10000.' };
                return { valid: true, msg: '' };
            }

            function validateEnum(value, allowedValues, fieldName) {
                if (!value || !allowedValues.includes(value)) return { valid: false,
                    msg: `${fieldName} must be one of: ${allowedValues.join(', ')}.` };
                return { valid: true, msg: '' };
            }

            function validateDateISO(value, fieldName) {
                if (!value || value.trim() === '') return { valid: false, msg: `${fieldName} is required.` };
                const isoRegex = /^\d{4}-\d{2}-\d{2}$/;
                if (!isoRegex.test(value)) return { valid: false,
                    msg: `${fieldName} must be in YYYY-MM-DD format.` };
                const parsed = new Date(value + 'T00:00:00');
                if (isNaN(parsed.getTime())) return { valid: false, msg: `${fieldName} is not a valid date.` };
                return { valid: true, msg: '' };
            }

            function setFieldState(inputEl, errEl, validationResult) {
                if (!inputEl || !errEl) return;
                inputEl.classList.remove('valid', 'invalid');
                errEl.textContent = '';
                if (validationResult.valid && inputEl.value.trim() !== '') {
                    inputEl.classList.add('valid');
                } else if (!validationResult.valid) {
                    inputEl.classList.add('invalid');
                    errEl.textContent = validationResult.msg;
                }
            }

            // âââ LIVE VALIDATION BINDING ââââââââââââââââââââ
            function bindLiveValidation(formEl, validationMap) {
                Object.keys(validationMap).forEach(fieldId => {
                    const inputEl = document.getElementById(fieldId);
                    const errEl = document.getElementById('err-' + fieldId);
                    if (!inputEl || !errEl) return;
                    inputEl.addEventListener('input', () => {
                        const result = validationMap[fieldId](inputEl.value);
                        setFieldState(inputEl, errEl, result);
                    });
                    inputEl.addEventListener('blur', () => {
                        const result = validationMap[fieldId](inputEl.value);
                        setFieldState(inputEl, errEl, result);
                    });
                });
            }

            // FG live validation
            bindLiveValidation(formFg, {
                'fg_id': (v) => validateId(v, 'FG ID'),
                'fg_length_mm': (v) => validateDimension(v, 'Length'),
                'fg_width_mm': (v) => validateDimension(v, 'Width'),
                'fg_height_mm': (v) => validateDimension(v, 'Height'),
                'fg_target_qty': (v) => validateTargetQty(v),
                'fg_source': (v) => validateEnum(v, SOURCE_ENUM, 'Source'),
            });

            // Carton live validation
            bindLiveValidation(formCarton, {
                'carton_id': (v) => validateId(v, 'Carton ID'),
                'carton_length_mm': (v) => validateDimension(v, 'Length'),
                'carton_width_mm': (v) => validateDimension(v, 'Width'),
                'carton_height_mm': (v) => validateDimension(v, 'Height'),
                'carton_dimension_source': (v) => validateEnum(v, DIM_SOURCE_ENUM, 'Dimension Source'),
            });

            // BOM live validation
            bindLiveValidation(formBom, {
                'bom_fg_id': (v) => validateId(v, 'FG ID'),
                'bom_carton_id': (v) => validateId(v, 'Carton ID'),
                'bom_source_bom_date': (v) => validateDateISO(v, 'Source BOM Date'),
                'bom_confidence': (v) => validateEnum(v, CONFIDENCE_ENUM, 'Confidence'),
            });

            // âââ FULL FORM VALIDATION ââââââââââââââââââââââ
            function validateFgForm() {
                const errors = [];
                const fgIdVal = $('#fg_id').value;
                const lenVal = $('#fg_length_mm').value;
                const widVal = $('#fg_width_mm').value;
                const hgtVal = $('#fg_height_mm').value;
                const tgtVal = $('#fg_target_qty').value;
                const srcVal = $('#fg_source').value;

                const r1 = validateId(fgIdVal, 'FG ID');
                if (!r1.valid) errors.push({ field: 'fg_id', message: r1.msg });
                const r2 = validateDimension(lenVal, 'Length');
                if (!r2.valid) errors.push({ field: 'length_mm', message: r2.msg });
                const r3 = validateDimension(widVal, 'Width');
                if (!r3.valid) errors.push({ field: 'width_mm', message: r3.msg });
                const r4 = validateDimension(hgtVal, 'Height');
                if (!r4.valid) errors.push({ field: 'height_mm', message:
