
const SETTINGS = {
    mapStart: [43.919, -80.094],
    mapZoom: 18,
    taxRate: 0.13,
    defaultDiscountRate: 0.05,
    pricingTierBands: [
        { endSqFt: 8000, incrementSqFt: 500 },
        { endSqFt: null, incrementSqFt: 1000 }
    ],
    pricingWorkbookUrl: '2026 PRICE TABLE - 5%.xlsx',
    pricingWorksheetName: 'Sheet1',
    installmentCount: 5,
    sqftPerSqMeter: 10.76391,
    sqftPerAcre: 43560,
    geocoder: {
        url: 'https://geocode.arcgis.com/arcgis/rest/services/World/GeocodeServer/findAddressCandidates',
        ontarioExtent: '-95.2,41.6,-74.3,56.9',
        minScore: 90,
        ambiguousScoreGap: 2,
        ambiguousDistanceMeters: 75,
        acceptedTypes: new Set(['PointAddress', 'Subaddress', 'StreetAddress'])
    }
};

const PRICING_CACHE_STORAGE_KEY = 'dufferin-lawn-life-spreadsheet-pricing-cache-v1';

const PACKAGES = [
    {
        key: 'golfGreen',
        name: 'GolfGreen Program',
        pricingRules: []
    },
    {
        key: 'ecolawn',
        name: 'EcoLawn Program',
        pricingRules: []
    },
    {
        key: 'greenscape',
        name: 'GreenScape Program',
        pricingRules: []
    }
];


const state = {
    verified: false,
    verifiedLabel: '',
    verifiedPoint: null,
    verifiedAttrs: null,
    verifiedScore: null,
    addressMarker: null,
    selectedDiscountRate: SETTINGS.defaultDiscountRate,
    pricingLoaded: false,
    pricingError: '',
    quotePdfBytes: null,
    lastTierSqFt: 0,
    lastMeasuredSqFeet: 0,
    lastAdditionalSqFeet: 0,
    lastTotalSqFeet: 0
};
const $ = id => document.getElementById(id);

const ui = {
    address: $('address'),
    searchBtn: $('searchBtn'),
    showAddressMarker: $('showAddressMarker'),
    status: $('status'),
    area: $('area'),
    calculatedArea: $('calculatedArea'),
    pricingTier: $('pricingTier'),
    pricing: $('pricing'),
    importPricingFile: $('importPricingFile'),
    additionalSqFt: $('additionalSqFt'),
    quotePdfFile: $('quotePdfFile'),
    quoteName: $('quoteName'),
    quotePhone: $('quotePhone'),
    quoteEmail: $('quoteEmail'),
    quoteProgram: $('quoteProgram'),
    exportQuotePdfBtn: $('exportQuotePdfBtn'),

    setStatus(message, type = '') {
        this.status.className = type;
        this.status.textContent = message;
    },

    setBusy(isBusy) {
        this.searchBtn.disabled = isBusy;
        this.searchBtn.textContent = isBusy ? 'Checking...' : 'Find Property';
    },

    resetEstimate() {
        this.area.textContent = '0';
        if (this.calculatedArea) this.calculatedArea.textContent = '0';
        this.pricingTier.textContent = '0';
        this.pricing.innerHTML = 'Find a property and draw the service area.';
        state.lastTierSqFt = 0;
        state.lastMeasuredSqFeet = 0;
        state.lastAdditionalSqFeet = 0;
        state.lastTotalSqFeet = 0;
    }
};

const map = createMap();
const drawnItems = L.featureGroup().addTo(map);
addDrawingTools();
wireEvents();
initializePricingFromSpreadsheet();



async function initializePricingFromSpreadsheet() {
    try {
        state.pricingLoaded = false;
        state.pricingError = '';
        await ensureXlsxLibrary();

        if (loadCachedPricingPackages()) {
            state.pricingLoaded = true;
            state.pricingError = '';
            calculateEstimate();
            return;
        }

        if (window.location.protocol === 'file:') {
            throw new Error('Browser security blocked automatic spreadsheet loading from a local folder.');
        }

        const response = await fetch(encodeURI(SETTINGS.pricingWorkbookUrl), { cache: 'no-store' });
        if (!response.ok) {
            throw new Error(`Could not load ${SETTINGS.pricingWorkbookUrl}. HTTP ${response.status}`);
        }

        const arrayBuffer = await response.arrayBuffer();
        loadPricingPackagesFromWorkbook(arrayBuffer, { saveToCache: true });
        state.pricingLoaded = true;
        state.pricingError = '';
        calculateEstimate();
    } catch (error) {
        console.error(error);
        state.pricingLoaded = false;
        state.pricingError = error.message || 'Pricing spreadsheet could not be loaded.';
        ui.setStatus(
            'Pricing spreadsheet could not be loaded automatically. Select the spreadsheet file below.',
            'error'
        );
        calculateEstimate();
    }
}

async function ensureXlsxLibrary() {
    if (window.XLSX) return;

    await new Promise((resolve, reject) => {
        const script = document.createElement('script');
        script.src = 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js';
        script.onload = resolve;
        script.onerror = () => reject(new Error('Could not load the XLSX spreadsheet parser library.'));
        document.head.appendChild(script);
    });
}

function handlePricingWorkbookImport(event) {
    const file = event.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();

    reader.onload = () => {
        try {
            loadPricingPackagesFromWorkbook(reader.result, { saveToCache: true });
            state.pricingLoaded = true;
            state.pricingError = '';
            ui.setStatus('Pricing spreadsheet loaded and saved in this browser.', 'ok');
            calculateEstimate();
        } catch (error) {
            console.error(error);
            state.pricingLoaded = false;
            state.pricingError = error.message || 'Could not read pricing spreadsheet.';
            ui.setStatus('Could not read that pricing spreadsheet. Check the column layout.', 'error');
            calculateEstimate();
        } finally {
            event.target.value = '';
        }
    };

    reader.onerror = () => {
        state.pricingLoaded = false;
        state.pricingError = 'Could not read selected pricing file.';
        ui.setStatus('Could not read selected pricing file.', 'error');
    };

    reader.readAsArrayBuffer(file);
}

function loadPricingPackagesFromWorkbook(arrayBuffer, options = {}) {
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
    const sheetName = workbook.Sheets[SETTINGS.pricingWorksheetName]
        ? SETTINGS.pricingWorksheetName
        : workbook.SheetNames[0];

    const sheet = workbook.Sheets[sheetName];
    if (!sheet) throw new Error('Pricing workbook does not contain a readable worksheet.');

    const rows = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        raw: true,
        defval: null
    });

    const pricePoints = parsePricingRows(rows);

    const packages = [
        makeSpreadsheetPackage('golfGreen', 'GolfGreen Program', pricePoints.golfGreen),
        makeSpreadsheetPackage('ecolawn', 'EcoLawn Program', pricePoints.ecolawn),
        makeSpreadsheetPackage('greenscape', 'GreenScape Program', pricePoints.greenscape)
    ];

    packages.forEach(pkg => {
        if (!pkg.pricingRules[0].pricePoints.length) {
            throw new Error(`${pkg.name} pricing was not found in the spreadsheet.`);
        }
    });

    PACKAGES.splice(0, PACKAGES.length, ...packages);

    if (options.saveToCache) {
        savePricingPackagesToCache(packages);
    }

    console.log('Pricing spreadsheet loaded:', PACKAGES.map(pkg => ({
        program: pkg.name,
        pricePoints: pkg.pricingRules[0].pricePoints.length
    })));
}

function savePricingPackagesToCache(packages) {
    try {
        localStorage.setItem(PRICING_CACHE_STORAGE_KEY, JSON.stringify({
            savedAt: new Date().toISOString(),
            packages
        }));
    } catch (error) {
        console.warn('Pricing cache could not be saved.', error);
    }
}

function loadCachedPricingPackages() {
    try {
        const cached = localStorage.getItem(PRICING_CACHE_STORAGE_KEY);
        if (!cached) return false;

        const parsed = JSON.parse(cached);
        if (!Array.isArray(parsed.packages)) return false;

        PACKAGES.splice(0, PACKAGES.length, ...parsed.packages);
        ui.setStatus('Pricing loaded from saved spreadsheet cache.', 'ok');
        return true;
    } catch (error) {
        console.warn('Pricing cache could not be loaded.', error);
        return false;
    }
}

function parsePricingRows(rows) {
    const output = {
        golfGreen: [],
        ecolawn: [],
        greenscape: []
    };

    rows.forEach(row => {
        const tierSqFt = normalizeSpreadsheetSqFt(row[0]);
        if (!tierSqFt) return;

        // Spreadsheet layout:
        // A = Sq. Ft.
        // B/C = GolfGreen prepaid/installments
        // F/H = EcoLawn prepaid/installments
        // J/K = GreenScape prepaid/installments
        //
        // Important: EcoLawn also has an extra X 5 helper column in G,
        // so the real EcoLawn installment price is column H.
        addPricePoint(output.golfGreen, tierSqFt, row[1], row[2]);
        addPricePoint(output.ecolawn, tierSqFt, row[5], row[7]);
        addPricePoint(output.greenscape, tierSqFt, row[9], row[10]);
    });

    Object.values(output).forEach(points => {
        points.sort((a, b) => Number(a.tierSqFt) - Number(b.tierSqFt));
    });

    return output;
}

function addPricePoint(points, tierSqFt, prepaidValue, installmentValue) {
    const prepaid = parseSpreadsheetNumber(prepaidValue);
    const installmentPayment = parseSpreadsheetNumber(installmentValue);

    if (!isFilledNumber(prepaid) || !isFilledNumber(installmentPayment)) return;

    points.push({
        tierSqFt,
        prepaid: roundMoney(prepaid),
        installmentPayment: roundMoney(installmentPayment)
    });
}

function makeSpreadsheetPackage(key, name, points) {
    const firstPoint = points[0];
    const lastPoint = points[points.length - 1];

    return {
        key,
        name,
        pricingRules: [
            {
                startSqFt: firstPoint.tierSqFt,
                endSqFt: lastPoint.tierSqFt,
                pricePoints: points
            }
        ]
    };
}

function normalizeSpreadsheetSqFt(value) {
    const number = parseSpreadsheetNumber(value);
    if (!isFilledNumber(number) || number <= 0) return null;

    // The spreadsheet uses 9, 10, 11, etc. to mean 9,000, 10,000, 11,000 sq ft.
    return number < 500 ? Math.round(number * 1000) : Math.round(number);
}

function parseSpreadsheetNumber(value) {
    if (value === null || value === undefined || value === '') return null;
    if (typeof value === 'number') return Number.isFinite(value) ? value : null;

    const cleaned = String(value).replace(/[$,\s]/g, '');
    const number = Number(cleaned);

    return Number.isFinite(number) ? number : null;
}


function createMap() {
    const mapInstance = L.map('map').setView(SETTINGS.mapStart, SETTINGS.mapZoom);

    const satellite = L.tileLayer(
        'https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}',
        { attribution: 'Tiles (c) Esri', maxZoom: 22 }
    ).addTo(mapInstance);

    const labels = L.tileLayer(
        'https://server.arcgisonline.com/ArcGIS/rest/services/Reference/World_Boundaries_and_Places/MapServer/tile/{z}/{y}/{x}',
        { attribution: 'Labels (c) Esri', maxZoom: 22 }
    ).addTo(mapInstance);

    const propertyLines = L.tileLayer.wms(
        'https://www.wpsgn.ca/s22/services/Parcels/Ontario_Parcel/MapServer/WMSServer',
        { layers: '0', format: 'image/png', transparent: true }
    ).addTo(mapInstance);

    L.control.layers(
        { Satellite: satellite },
        { Labels: labels, 'Property Lines': propertyLines }
    ).addTo(mapInstance);

    return mapInstance;
}

function addDrawingTools() {
    map.addControl(new L.Control.Draw({
        draw: {
            polygon: true,
            rectangle: true,
            polyline: false,
            circle: false,
            circlemarker: false,
            marker: false
        },
        edit: {
            featureGroup: drawnItems,
            remove: true
        }
    }));
}

function wireEvents() {
    map.on(L.Draw.Event.CREATED, handleShapeCreated);
    map.on(L.Draw.Event.EDITED, calculateEstimate);
    map.on(L.Draw.Event.DELETED, calculateEstimate);

    ui.searchBtn.addEventListener('click', searchAddress);
    ui.address.addEventListener('input', handleAddressInput);
    ui.address.addEventListener('keydown', event => {
        if (event.key === 'Enter') searchAddress();
    });

    document.querySelectorAll('input[name="discountRate"]').forEach(input => {
        input.addEventListener('change', handleDiscountChange);
    });

    if (ui.importPricingFile) {
        ui.importPricingFile.addEventListener('change', handlePricingWorkbookImport);
    }

    if (ui.showAddressMarker) {
        ui.showAddressMarker.addEventListener('change', updateAddressMarker);
    }

    if (ui.additionalSqFt) {
        ui.additionalSqFt.addEventListener('input', calculateEstimate);
    }

    if (ui.quotePdfFile) {
        ui.quotePdfFile.addEventListener('change', handleQuotePdfImport);
    }

    if (ui.exportQuotePdfBtn) {
        ui.exportQuotePdfBtn.addEventListener('click', exportFilledQuotePdf);
    }
}

function handleDiscountChange(event) {
    state.selectedDiscountRate = Number(event.target.value);
    calculateEstimate();
}

function handleShapeCreated(event) {
    if(!state.verified) {
        alert('Find a verified address first. Estimate blocked to prevent measuring the wrong property.');
        return;
    }

    const layer = event.layer;

    setMeasuredAreaStyle(layer);
    addMeasuredAreaDeleteEvents(layer);

    drawnItems.addLayer(layer);
    calculateEstimate();
}

function setMeasuredAreaStyle(layer) {
    layer.setStyle({
            color: '#006b2d',
            fillColor: '#006b2d',
            fillOpacity: 0.35
    });
}

function addMeasuredAreaDeleteEvents(layer) {
    layer.on('mouseover', function () {
        layer.setStyle({
            color: '#ff0000',
            fillColor: '#ff6666',
            fillOpacity: 0.45
        });

        layer.bindTooltip('Right-click to delete this measured area', {
            sticky: true
        }).openTooltip();
    });

    layer.on('mouseout', function () {
        setMeasuredAreaStyle(layer);
        layer.closeTooltip();
    });

    layer.on('contextmenu', function (event) {
        event.originalEvent.preventDefault();

        const confirmed = confirm('Delete this measured area?');

        if (confirmed) {
            drawnItems.removeLayer(layer);
            calculateEstimate();
        }
    });
}

function handleAddressInput() {
    clearProperty({ removeMarker: true, clearDrawings: true });
    ui.setStatus('Address changed. Click Find Property before drawing.', 'warn');
}

function clearProperty(options = {}) {
    state.verified = false;
    ui.resetEstimate();

    if (options.clearDrawings) drawnItems.clearLayers();

    if (options.removeMarker && state.addressMarker) {
        map.removeLayer(state.addressMarker);
        state.addressMarker = null;
    }
}

function calculateEstimate() {
    if (!state.verified) {
        ui.resetEstimate();
        return;
    }

    let totalSqMeters = 0;
    drawnItems.eachLayer(layer => {
        totalSqMeters += turf.area(layer.toGeoJSON());
    });

    const measuredSqFeet = totalSqMeters * SETTINGS.sqftPerSqMeter;
    const additionalSqFeet = Math.max(0, Number(ui.additionalSqFt?.value || 0));
    const totalSqFeet = measuredSqFeet + additionalSqFeet;

    state.lastMeasuredSqFeet = measuredSqFeet;
    state.lastAdditionalSqFeet = additionalSqFeet;
    state.lastTotalSqFeet = totalSqFeet;

    // Keep measured area as the map-drawn area only.
    // The added square footage is shown separately and only included once in calculated area/pricing.
    ui.area.textContent = formatSqFt(measuredSqFeet);
    if (ui.calculatedArea) ui.calculatedArea.textContent = formatSqFt(totalSqFeet);

    renderPricing(totalSqFeet);
}

function renderPricing(totalSqFeet) {
    if (totalSqFeet <= 0) {
        ui.pricingTier.textContent = '0';
        ui.pricing.innerHTML = 'Draw the service area to show package pricing.';
        return;
    }

    if (!state.pricingLoaded) {
        ui.pricingTier.textContent = formatSqFt(pricingTierFor(totalSqFeet));
        ui.pricing.innerHTML = state.pricingError
            ? `
                <div class="price-band">
                    Pricing unavailable: ${state.pricingError}<br>
                    <strong>Select your pricing spreadsheet:</strong><br>
                    <input id="pricingSpreadsheetPicker" type="file" accept=".xlsx,.xls">
                    <br><small>If you opened this page directly from a folder, your browser may block automatic spreadsheet loading.</small>
                </div>
            `
            : 'Loading pricing spreadsheet...';

        const picker = document.getElementById('pricingSpreadsheetPicker');
        if (picker) picker.addEventListener('change', handlePricingWorkbookImport);
        return;
    }

    const tierSqFt = pricingTierFor(totalSqFeet);
    state.lastTierSqFt = tierSqFt;
    const acres = roundMoney(totalSqFeet / SETTINGS.sqftPerAcre).toFixed(2);

    ui.pricingTier.textContent = formatSqFt(tierSqFt);
    ui.pricing.innerHTML = `
        <div class="price-band">
            Measured area: ${formatSqFt(state.lastMeasuredSqFeet)} sq ft<br>
            Additional sq ft: ${formatSqFt(state.lastAdditionalSqFeet)} sq ft<br>
            Calculated area: ${formatSqFt(totalSqFeet)} sq ft (${acres} acres)<br>
            Rounded up to: ${formatSqFt(tierSqFt)} sq ft pricing<br>
            Prepaid discount: ${formatPercent(state.selectedDiscountRate)}<br>
            Installments: no discount applied
        </div>
        <table class="pricing-table">
            <thead>
                <tr>
                    <th>Package</th>
                    <th>Prepaid</th>
                    <th>Prepaid + HST</th>
                    <th>Installments</th>
                    <th>Installments + HST</th>
                </tr>
            </thead>
            <tbody>
                ${PACKAGES.map(pkg => pricingRow(pkg, tierSqFt)).join('')}
            </tbody>
        </table>
    `;
}

function pricingRow(pkg, tierSqFt) {
    const price = calculatePackagePrice(pkg, tierSqFt);

    if (!price) {
        return packageRow(
            pkg.name,
            missingPrice(),
            missingPrice(),
            missingPrice(),
            missingPrice()
        );
    }

    if (price.manualQuote) {
        return packageRow(
            pkg.name,
            'Manual quote',
            'Manual quote',
            'Manual quote',
            'Manual quote'
        );
    }

    const adjustedPrice = applySelectedDiscount(price);

    const prepaidWithTax = roundMoney(adjustedPrice.prepaid * (1 + SETTINGS.taxRate));
    const installmentPaymentWithTax = roundMoney(adjustedPrice.installmentPayment * (1 + SETTINGS.taxRate));
    const installmentTotalWithTax = roundMoney(adjustedPrice.installmentTotal * (1 + SETTINGS.taxRate));

    const installments = `
        ${SETTINGS.installmentCount} x ${formatMoney(adjustedPrice.installmentPayment)}
        <br><small>Total: ${formatMoney(adjustedPrice.installmentTotal)}</small>
        ${adjustedPrice.savings > 0 ? `<span class="discount">Prepaid saves ${formatMoney(adjustedPrice.savings)}</span>` : ''}
    `;

    const installmentsWithTax = `
        ${SETTINGS.installmentCount} x ${formatMoney(installmentPaymentWithTax)}
        <br><small>Total: ${formatMoney(installmentTotalWithTax)}</small>
    `;

    return packageRow(
        pkg.name,
        formatMoney(adjustedPrice.prepaid),
        formatMoney(prepaidWithTax),
        installments,
        installmentsWithTax
    );
}

function packageRow(packageName, prepaidHtml, prepaidWithTaxHtml, installmentHtml, installmentWithTaxHtml) {
    return `
        <tr>
            <td>${packageName}</td>
            <td>${prepaidHtml}</td>
            <td>${prepaidWithTaxHtml}</td>
            <td>${installmentHtml}</td>
            <td>${installmentWithTaxHtml}</td>
        </tr>
    `;
}

function applySelectedDiscount(price) {
    const selectedDiscountRate = Number(state.selectedDiscountRate);

    // This applies to ALL programs: GolfGreen, EcoLawn, and GreenScape.
    //
    // The spreadsheet prepaid columns are already the 5% prepaid prices.
    // To show a different prepaid discount, first rebuild the full regular price,
    // then apply the selected prepaid discount.
    //
    // Installment prices are NEVER discounted. The installment payment always stays
    // exactly as listed in the spreadsheet, regardless of 5%, 7%, or 10% selection.
    const prepaidBase = Number(price.prepaid) / (1 - SETTINGS.defaultDiscountRate);
    const prepaid = roundMoney(prepaidBase * (1 - selectedDiscountRate));

    const installmentPayment = roundMoney(Number(price.installmentPayment));
    const installmentTotal = roundMoney(installmentPayment * SETTINGS.installmentCount);

    return {
        prepaid,
        installmentPayment,
        installmentTotal,
        savings: roundMoney(installmentTotal - prepaid)
    };
}

function calculatePackagePrice(pkg, tierSqFt) {
    if (!hasPricingRules(pkg)) return null;

    const rule = matchingRule(pkg, tierSqFt);
    if (!rule) return { manualQuote: true };

    if (Array.isArray(rule.pricePoints)) {
        return calculateExactPricePoint(rule, tierSqFt);
    }

    return { manualQuote: true };
}

function calculateExactPricePoint(rule, tierSqFt) {
    const points = [...rule.pricePoints].sort((a, b) =>
        Number(a.tierSqFt) - Number(b.tierSqFt)
    );

    const exactPoint = points.find(item =>
        Number(item.tierSqFt) === Number(tierSqFt)
    );

    let prepaid;
    let installmentPayment;

    if (exactPoint) {
        prepaid = Number(exactPoint.prepaid);
        installmentPayment = Number(exactPoint.installmentPayment);
    } else {
        const lowerPoint = [...points].reverse().find(item =>
            Number(item.tierSqFt) < Number(tierSqFt)
        );
        const upperPoint = points.find(item =>
            Number(item.tierSqFt) > Number(tierSqFt)
        );

        if (!lowerPoint || !upperPoint) return { manualQuote: true };

        const lowerSqFt = Number(lowerPoint.tierSqFt);
        const upperSqFt = Number(upperPoint.tierSqFt);
        const ratio = (Number(tierSqFt) - lowerSqFt) / (upperSqFt - lowerSqFt);

        prepaid = interpolateMoney(lowerPoint.prepaid, upperPoint.prepaid, ratio);
        installmentPayment = interpolateMoney(lowerPoint.installmentPayment, upperPoint.installmentPayment, ratio);
    }

    const installmentTotal = roundMoney(installmentPayment * SETTINGS.installmentCount);

    return {
        prepaid: roundMoney(prepaid),
        installmentPayment: roundMoney(installmentPayment),
        installmentTotal,
        savings: roundMoney(installmentTotal - prepaid)
    };
}

function matchingRule(pkg, tierSqFt) {
    return pkg.pricingRules.find(rule => {
        const starts = tierSqFt >= Number(rule.startSqFt);
        const ends = rule.endSqFt == null || tierSqFt <= Number(rule.endSqFt);
        return starts && ends;
    }) || null;
}

function hasPricingRules(pkg) {
    return Array.isArray(pkg.pricingRules) && pkg.pricingRules.length > 0;
}

function pricingTierFor(sqFeet) {
    if (sqFeet <= 0) return 0;

    const band = SETTINGS.pricingTierBands.find(band =>
        band.endSqFt == null || sqFeet <= Number(band.endSqFt)
    );

    if (!band) return Math.ceil(sqFeet / 1000) * 1000;

    const incrementSqFt = Number(band.incrementSqFt);
    return Math.ceil(sqFeet / incrementSqFt) * incrementSqFt;
}

async function searchAddress() {
    clearProperty({ removeMarker: true, clearDrawings: true });

    const rawAddress = ui.address.value.trim();
    const houseNumber = getInputHouseNumber(rawAddress);

    if (!rawAddress) {
        return blockEstimate('Enter an address first.', 'Enter an address');
    }

    if (!houseNumber) {
        return blockEstimate(
            'A civic house number is required. Example: 33 Maplewood Drive.',
            'Please enter the full civic address, including the house number.'
        );
    }

    try {
        ui.setBusy(true);
        ui.setStatus('Checking address accuracy...');

        const candidates = await fetchAddressCandidates(rawAddress);

        if (candidates.length === 0) {
            return blockEstimate('Address not found in Ontario.', 'Address not found in Ontario.');
        }

        const validCandidates = candidates
            .filter(candidate => isValidAddressCandidate(candidate, houseNumber))
            .sort((a, b) => Number(b.score || 0) - Number(a.score || 0));

        if (validCandidates.length === 0) {
            const message = 'No rooftop/point-address level Ontario match was found. Estimate blocked. ' +
                formatCandidateDebug(candidates[0]);

            return blockEstimate(
                message,
                'This address could not be verified accurately enough for an automatic estimate. Try adding the town/city or postal code.'
            );
        }

        const best = validCandidates[0];

        if (hasAmbiguousMatch(best, validCandidates)) {
            return blockEstimate(
                'More than one Ontario property closely matches this address. Estimate blocked. Add the town/city or postal code to avoid the wrong house.',
                'More than one Ontario property closely matches this address. Please add the town/city or postal code.'
            );
        }

        verifyProperty(best);
    } catch (error) {
        console.error(error);
        blockEstimate(
            'Address verification service failed. Estimate blocked.',
            'Address verification failed. Estimate blocked to prevent wrong-property pricing.'
        );
    } finally {
        ui.setBusy(false);
    }
}

async function fetchAddressCandidates(address) {
    const params = new URLSearchParams({
        f: 'pjson',
        SingleLine: normalizeAddressForOntario(address),
        sourceCountry: 'CAN',
        countryCode: 'CAN',
        category: 'Address',
        outFields: 'Match_addr,LongLabel,ShortLabel,Addr_type,Score,DisplayX,DisplayY,City,Region,RegionAbbr,Postal,Country,AddNum,StName',
        maxLocations: '10',
        forStorage: 'false',
        locationType: 'rooftop',
        matchOutOfRange: 'false',
        outSR: '4326',
        searchExtent: SETTINGS.geocoder.ontarioExtent
    });

    const response = await fetch(`${SETTINGS.geocoder.url}?${params}`);
    if (!response.ok) throw new Error('Geocoder HTTP error ' + response.status);

    const data = await response.json();
    if (data.error) throw new Error(data.error.message || 'Geocoder error');

    return Array.isArray(data.candidates) ? data.candidates : [];
}

function isValidAddressCandidate(candidate, inputHouseNumber) {
    const attrs = candidate.attributes || {};
    const score = Number(candidate.score || 0);

    return isOntarioCandidate(candidate) &&
        SETTINGS.geocoder.acceptedTypes.has(attrs.Addr_type) &&
        score >= SETTINGS.geocoder.minScore &&
        getCandidateHouseNumber(candidate) === inputHouseNumber &&
        Boolean(getCandidatePoint(candidate));
}

function hasAmbiguousMatch(best, candidates) {
    const bestPoint = getCandidatePoint(best);
    if (!bestPoint) return true;

    return candidates.some(candidate => {
        if (candidate === best) return false;

        const scoreGap = Math.abs(Number(candidate.score || 0) - Number(best.score || 0));
        if (scoreGap > SETTINGS.geocoder.ambiguousScoreGap) return false;

        const point = getCandidatePoint(candidate);
        if (!point) return false;

        return distanceMeters(bestPoint.lat, bestPoint.lon, point.lat, point.lon) >
            SETTINGS.geocoder.ambiguousDistanceMeters;
    });
}

function verifyProperty(candidate) {
    const attrs = candidate.attributes || {};
    const point = getCandidatePoint(candidate);
    const label = attrs.LongLabel || attrs.Match_addr || candidate.address;

    state.verified = true;
    state.verifiedLabel = label;
    state.verifiedPoint = point;
    state.verifiedAttrs = attrs;
    state.verifiedScore = candidate.score;

    map.setView([point.lat, point.lon], 21);

    updateAddressMarker();

    ui.setStatus(`Verified address match: ${label}`, 'ok');
    calculateEstimate();
}

function updateAddressMarker() {
    if (state.addressMarker) {
        map.removeLayer(state.addressMarker);
        state.addressMarker = null;
    }

    if (!ui.showAddressMarker?.checked) return;
    if (!state.verifiedPoint) return;

    state.addressMarker = L.marker([state.verifiedPoint.lat, state.verifiedPoint.lon])
    .addTo(map)
    .bindPopup(`
        <strong>Verified Property</strong><br>
        ${state.verifiedLabel}<br>
        <small>${state.verifiedAttrs?.Addr_type || ''}, score ${state.verifiedScore || ''}</small>
    `);
}

function blockEstimate(statusMessage, alertMessage) {
    ui.setStatus(statusMessage, 'error');
    alert(alertMessage);
}

function normalizeAddressForOntario(address) {
    const parts = [String(address).trim()];
    const lower = parts[0].toLowerCase();

    if (!lower.includes('ontario') && !/\bon\b/.test(lower)) parts.push('Ontario');
    if (!lower.includes('canada')) parts.push('Canada');

    return parts.join(', ');
}

function isOntarioCandidate(candidate) {
    const attrs = candidate.attributes || {};
    const region = String(attrs.Region || '').toLowerCase();
    const regionAbbr = String(attrs.RegionAbbr || '').toLowerCase();
    const country = String(attrs.Country || '').toLowerCase();
    const label = candidateSearchText(candidate).toLowerCase();

    const inOntario = region === 'ontario' ||
        regionAbbr === 'on' ||
        label.includes(' ontario') ||
        label.includes(', on,') ||
        label.includes(', on ') ||
        label.includes(' on, canada');

    const inCanada = !country ||
        country === 'can' ||
        country === 'ca' ||
        country === 'canada' ||
        label.includes('canada');

    return inOntario && inCanada;
}

function candidateSearchText(candidate) {
    const attrs = candidate.attributes || {};
    return [
        candidate.address,
        attrs.Match_addr,
        attrs.LongLabel,
        attrs.ShortLabel
    ].filter(Boolean).join(' ');
}

function getCandidatePoint(candidate) {
    const attrs = candidate.attributes || {};
    const lon = parseFloat(firstFilled(attrs.DisplayX, candidate.location?.x));
    const lat = parseFloat(firstFilled(attrs.DisplayY, candidate.location?.y));

    return Number.isFinite(lat) && Number.isFinite(lon) ? { lat, lon } : null;
}

function getInputHouseNumber(address) {
    const match = String(address).match(/\b\d+[a-zA-Z]?\b/);
    return match ? normalizeHouseNumber(match[0]) : '';
}

function getCandidateHouseNumber(candidate) {
    const attrs = candidate.attributes || {};
    const labelNumber = String(candidateSearchText(candidate)).match(/\b\d+[a-zA-Z]?\b/);
    return normalizeHouseNumber(attrs.AddNum || (labelNumber ? labelNumber[0] : ''));
}

function normalizeHouseNumber(value) {
    return String(value || '').trim().toLowerCase().replace(/[^0-9a-z]/g, '');
}

function formatCandidateDebug(candidate) {
    if (!candidate) return 'No candidate returned.';

    const attrs = candidate.attributes || {};
    const type = attrs.Addr_type || 'unknown type';
    const score = candidate.score ?? 'unknown score';
    const label = attrs.LongLabel || attrs.Match_addr || candidate.address || 'unknown address';

    return `Best returned match was ${type}, score ${score}: ${label}`;
}

function distanceMeters(lat1, lon1, lat2, lon2) {
    const radiusMeters = 6371000;
    const toRad = value => value * Math.PI / 180;
    const dLat = toRad(lat2 - lat1);
    const dLon = toRad(lon2 - lon1);

    const a = Math.sin(dLat / 2) ** 2 +
        Math.cos(toRad(lat1)) *
        Math.cos(toRad(lat2)) *
        Math.sin(dLon / 2) ** 2;

    return radiusMeters * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}


function handleQuotePdfImport(event) {
    const file = event.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();

    reader.onload = () => {
        state.quotePdfBytes = reader.result;
        ui.setStatus('Quote PDF loaded. You can now export a filled quote.', 'ok');
    };

    reader.onerror = () => {
        state.quotePdfBytes = null;
        ui.setStatus('Could not read the quote PDF.', 'error');
    };

    reader.readAsArrayBuffer(file);
}

async function exportFilledQuotePdf() {
    if (!window.PDFLib) {
        alert('PDF library has not loaded yet. Refresh the page and try again.');
        return;
    }

    if (!state.quotePdfBytes) {
        alert('Import the quote PDF first.');
        return;
    }

    if (!state.verified) {
        alert('Find a verified property first.');
        return;
    }

    if (!state.lastTierSqFt) {
        alert('Draw the service area first so pricing can be calculated.');
        return;
    }

    try {
        const pdfDoc = await PDFLib.PDFDocument.load(state.quotePdfBytes);
        const form = pdfDoc.getForm();
        const quotePrices = getQuotePricesForPdf();

        setPdfText(form, 'NAME', ui.quoteName?.value || '');
        setPdfText(form, 'ADDRESS', state.verifiedLabel || ui.address.value);
        setPdfText(form, 'PHONE', ui.quotePhone?.value || '');
        setPdfText(form, 'EMAIL', ui.quoteEmail?.value || '');
        setPdfText(form, 'Date', new Date().toLocaleDateString('en-CA'));
        setCenteredPdfText(form, 'Size', formatSqFt(state.lastTierSqFt), 9);
        setPdfText(form, 'Notes', `${Math.round(state.selectedDiscountRate * 100)}% Prepaid Discount`);
        setLawnAreaCheckboxes(form);

        // PDF package price fields:
        // PRICE 1 = GolfGreen prepaid before HST
        // PRICE 2 = EcoLawn prepaid before HST
        // PRICE 3 = GreenScape prepaid before HST
        setPdfText(form, 'PRICE 1', quotePrices.golfGreen);
        setPdfText(form, 'PRICE 2', quotePrices.ecolawn);
        setPdfText(form, 'PRICE 3', quotePrices.greenscape);

        form.updateFieldAppearances();

        const filledPdfBytes = await pdfDoc.save();
        downloadPdf(filledPdfBytes, makeQuoteFileName());
        ui.setStatus('Filled quote PDF exported.', 'ok');
    } catch (error) {
        console.error(error);
        ui.setStatus('Could not export the filled quote PDF.', 'error');
        alert('Could not export the filled quote PDF. Check the console for details.');
    }
}


function getQuotePricesForPdf() {
    const prices = {
        golfGreen: '',
        ecolawn: '',
        greenscape: ''
    };

    if (!state.pricingLoaded || !state.lastTierSqFt) return prices;

    PACKAGES.forEach(pkg => {
        const price = calculatePackagePrice(pkg, state.lastTierSqFt);
        if (!price || price.manualQuote) return;

        const adjustedPrice = applySelectedDiscount(price);

        // Use prepaid price BEFORE HST because the PDF already prints "+ HST" beside the line.
        prices[pkg.key] = Number(adjustedPrice.prepaid).toFixed(2);
    });

    return prices;
}

function setPdfText(form, fieldName, value) {
    try {
        const field = form.getTextField(fieldName);
        field.setText(String(value || ''));
    } catch (error) {
        console.warn(`PDF field not found: ${fieldName}`);
    }
}

function setCenteredPdfText(form, fieldName, value, fontSize = 9) {
    try {
        const field = form.getTextField(fieldName);
        field.setText(String(value || ''));
        field.setAlignment(PDFLib.TextAlignment.Center);
        field.setFontSize(fontSize);
    } catch (error) {
        console.warn(`PDF field could not be formatted: ${fieldName}`);
        setPdfText(form, fieldName, value);
    }
}

function setLawnAreaCheckboxes(form) {
    const selectedLawnArea = document.querySelector('input[name="quoteLawnArea"]:checked')?.value || 'whole';

    const lawnAreaFields = {
        whole: 'Check Box8',
        front: 'Check Box9.0',
        back: 'Check Box9.1'
    };

    Object.values(lawnAreaFields).forEach(fieldName => {
        try {
            form.getCheckBox(fieldName).uncheck();
        } catch (error) {
            console.warn(`PDF checkbox not found: ${fieldName}`);
        }
    });

    try {
        form.getCheckBox(lawnAreaFields[selectedLawnArea]).check();
    } catch (error) {
        console.warn(`Could not check lawn area field: ${selectedLawnArea}`);
    }
}

function downloadPdf(bytes, fileName) {
    const blob = new Blob([bytes], { type: 'application/pdf' });
    const url = URL.createObjectURL(blob);

    const link = document.createElement('a');
    link.href = url;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    link.remove();

    URL.revokeObjectURL(url);
}

function makeQuoteFileName() {

    const fullAddress = state.verifiedLabel || ui.address.value || 'property';
    const streetAddress = fullAddress.split(',')[0];
    const cleanAddress = streetAddress
        .trim()
        .replace(/^-|-$/g, '');         

    return `${cleanAddress || 'property'} quote.pdf`;
}

function firstFilled(...values) {
    return values.find(value => value !== null && value !== undefined && value !== '');
}

function isFilledNumber(value) {
    return value !== null &&
        value !== undefined &&
        value !== '' &&
        Number.isFinite(Number(value));
}

function interpolateMoney(startValue, endValue, ratio) {
    return Number(startValue) + (Number(endValue) - Number(startValue)) * Number(ratio);
}

function roundMoney(value) {
    return Math.round((Number(value) + 1e-9) * 100) / 100;
}

function formatSqFt(value) {
    return Math.round(Number(value) || 0).toString();
}

function formatPercent(value) {
    return Number(value).toLocaleString('en-CA', {
        style: 'percent',
        minimumFractionDigits: 0,
        maximumFractionDigits: 0
    });
}

function formatMoney(value) {
    if (!isFilledNumber(value)) return missingPrice();

    return '$' + Number(value).toLocaleString('en-CA', {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
    });
}

function missingPrice() {
    return '<span class="missing-price">Add formula</span>';
}