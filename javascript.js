// ==========================================
// DATASET FILE PATHS
// ==========================================

const capadUrl =
    "data/Collaborative_Australian_Protected_Areas_Database_(CAPAD)_2024_-_Terrestrial__.csv";

const pestUrl =
    "data/Date_EexoticEnvironmentalPests_v1.0.0.xlsx";


// ==========================================
// LOAD CSV AND EXCEL DATA
// ==========================================

async function readDataset(url) {
    const response = await fetch(url);

    if (!response.ok) {
        throw new Error(
            "Could not load " + url + ". HTTP status: " + response.status
        );
    }

    const buffer = await response.arrayBuffer();

    const workbook = XLSX.read(buffer, {
        type: "array"
    });

    const sheet = workbook.Sheets[workbook.SheetNames[0]];

    return XLSX.utils.sheet_to_json(sheet, {
        defval: ""
    });
}


// ==========================================
// HELPER FUNCTIONS
// ==========================================

function toNumber(value) {
    if (typeof value === "number") {
        return value;
    }

    if (value === null || value === undefined || value === "") {
        return NaN;
    }

    return Number(String(value).replace(/,/g, "").trim());
}


async function showChart(element, specification) {
    await vegaEmbed(element, specification, {
        actions: false,
        renderer: "canvas"
    });
}


// ==========================================
// CHART 1: PROTECTED AREAS BY STATE
// ==========================================

async function loadChart1(capadRows) {
    const counts = {};

    capadRows.forEach(row => {
        const state = String(row.STATE || "").trim();

        if (state !== "") {
            counts[state] = (counts[state] || 0) + 1;
        }
    });

    const rows = Object.keys(counts).map(state => ({
        state: state,
        count: counts[state]
    }));

    if (rows.length === 0) {
        throw new Error(
            "Chart 1: No STATE values found in the CAPAD dataset."
        );
    }

    await showChart("#chart1", {
        $schema: "https://vega.github.io/schema/vega-lite/v6.json",
        width: "container",
        height: 300,
        data: {
            values: rows
        },
        mark: {
            type: "bar",
            color: "#47785B",
            cornerRadiusEnd: 4,
            tooltip: true
        },
        encoding: {
            x: {
                field: "state",
                type: "nominal",
                title: "State or territory",
                sort: "-y",
                axis: {
                    labelAngle: 0
                }
            },
            y: {
                field: "count",
                type: "quantitative",
                title: "Number of protected-area records"
            },
            tooltip: [
                {
                    field: "state",
                    title: "State or territory"
                },
                {
                    field: "count",
                    title: "Number of records"
                }
            ]
        }
    });
}


// ==========================================
// CHART 2: LARGEST PROTECTED AREAS
// ==========================================

async function loadChart2(capadRows) {
    const rows = capadRows
        .map(row => ({
            name: String(row.NAME || "Unknown"),
            state: String(row.STATE || "Unknown"),
            area: toNumber(row.GIS_AREA)
        }))
        .filter(row =>
            Number.isFinite(row.area) && row.area > 0
        )
        .sort((a, b) => b.area - a.area)
        .slice(0, 15);

    if (rows.length === 0) {
        throw new Error(
            "Chart 2: No valid GIS_AREA values found."
        );
    }

    await showChart("#chart2", {
        $schema: "https://vega.github.io/schema/vega-lite/v6.json",
        width: "container",
        height: 400,
        data: {
            values: rows
        },
        mark: {
            type: "bar",
            color: "#C27A43",
            cornerRadiusEnd: 4,
            tooltip: true
        },
        encoding: {
            y: {
                field: "name",
                type: "nominal",
                sort: "-x",
                title: "Protected area",
                axis: {
                    labelLimit: 240
                }
            },
            x: {
                field: "area",
                type: "quantitative",
                title: "GIS area (dataset units)"
            },
            tooltip: [
                {
                    field: "name",
                    title: "Protected area"
                },
                {
                    field: "state",
                    title: "State"
                },
                {
                    field: "area",
                    title: "GIS area",
                    format: ",.3f"
                }
            ]
        }
    });
}


// ==========================================
// CHART 3: PROTECTED AREA SIZE VS SHAPE AREA
// ==========================================

async function loadChart3(capadRows) {
    const rows = capadRows
        .map(row => ({
            name: String(row.NAME || "Unknown"),
            state: String(row.STATE || "Unknown"),
            area: toNumber(row.GIS_AREA),
            shapeArea: toNumber(row.Shape__Area)
        }))
        .filter(row =>
            Number.isFinite(row.area) &&
            Number.isFinite(row.shapeArea) &&
            row.area > 0 &&
            row.shapeArea > 0
        );

    if (rows.length === 0) {
        throw new Error(
            "Chart 3: No valid GIS_AREA and Shape__Area values found."
        );
    }

    await showChart("#chart3", {
        $schema: "https://vega.github.io/schema/vega-lite/v6.json",
        width: "container",
        height: 340,
        data: {
            values: rows
        },
        mark: {
            type: "circle",
            size: 70,
            opacity: 0.65,
            color: "#47785B",
            tooltip: true
        },
        encoding: {
            x: {
                field: "area",
                type: "quantitative",
                scale: {
                    type: "log"
                },
                title: "GIS area (log scale)"
            },
            y: {
                field: "shapeArea",
                type: "quantitative",
                scale: {
                    type: "log"
                },
                title: "Shape area (log scale)"
            },
            tooltip: [
                {
                    field: "name",
                    title: "Protected area"
                },
                {
                    field: "state",
                    title: "State"
                },
                {
                    field: "area",
                    title: "GIS area",
                    format: ",.4f"
                },
                {
                    field: "shapeArea",
                    title: "Shape area",
                    format: ",.8f"
                }
            ]
        }
    });
}


// ==========================================
// CHART 4: EXOTIC PEST RISK HEATMAP
// ==========================================

async function loadChart4(pestRows) {
    const riskFields = [
        "entry_likelihood",
        "establishment_likelihood",
        "spread_likelihood"
    ];

    const rows = [];

    pestRows.forEach(row => {
        const pestName = String(
            row.common_name ||
            row.species_name ||
            "Unnamed pest"
        ).trim();

        riskFields.forEach(field => {
            const value = String(row[field] ?? "").trim();

            if (value !== "") {
                rows.push({
                    pest: pestName,
                    riskType: field.replace("_likelihood", "")
                        .replace(/_/g, " "),
                    likelihood: value
                });
            }
        });
    });

    if (rows.length === 0) {
        throw new Error(
            "Chart 4: No likelihood values found. Check the pest Excel column names."
        );
    }

    await showChart("#chart4", {
        $schema: "https://vega.github.io/schema/vega-lite/v6.json",
        width: "container",
        height: {
            step: 32
        },
        data: {
            values: rows
        },
        mark: {
            type: "rect",
            stroke: "white",
            strokeWidth: 1,
            tooltip: true
        },
        encoding: {
            x: {
                field: "riskType",
                type: "nominal",
                title: "Risk stage",
                axis: {
                    labelAngle: 0
                }
            },
            y: {
                field: "pest",
                type: "nominal",
                title: "Pest or disease",
                sort: "ascending",
                axis: {
                    labelLimit: 260
                }
            },
            color: {
                field: "likelihood",
                type: "nominal",
                title: "Recorded likelihood",
                scale: {
                    scheme: "oranges"
                }
            },
            tooltip: [
                {
                    field: "pest",
                    title: "Pest or disease"
                },
                {
                    field: "riskType",
                    title: "Risk stage"
                },
                {
                    field: "likelihood",
                    title: "Recorded likelihood"
                }
            ]
        }
    });
}


// ==========================================
// CHART 5: ENVIRONMENTAL IMPACT CATEGORIES
// ==========================================

async function loadChart5(pestRows) {
    const counts = {};

    pestRows.forEach(row => {
        const category = String(
            row.environmental_impact || ""
        ).trim();

        if (category !== "") {
            counts[category] = (counts[category] || 0) + 1;
        }
    });

    const rows = Object.keys(counts).map(category => ({
        impact: category,
        count: counts[category]
    }));

    if (rows.length === 0) {
        throw new Error(
            "Chart 5: No environmental_impact values found."
        );
    }

    await showChart("#chart5", {
        $schema: "https://vega.github.io/schema/vega-lite/v6.json",
        width: "container",
        height: 300,
        data: {
            values: rows
        },
        mark: {
            type: "bar",
            color: "#B66B4D",
            cornerRadiusEnd: 4,
            tooltip: true
        },
        encoding: {
            y: {
                field: "impact",
                type: "nominal",
                sort: "-x",
                title: "Environmental impact category",
                axis: {
                    labelLimit: 260
                }
            },
            x: {
                field: "count",
                type: "quantitative",
                title: "Number of pest records"
            },
            tooltip: [
                {
                    field: "impact",
                    title: "Environmental impact"
                },
                {
                    field: "count",
                    title: "Number of records"
                }
            ]
        }
    });
}


// ==========================================
// LOAD DATA AND DISPLAY ALL CHARTS
// ==========================================

async function main() {
    const status = document.getElementById("status");

    try {
        status.textContent = "Loading datasets...";

        const capadRows = await readDataset(capadUrl);
        const pestRows = await readDataset(pestUrl);

        // Debugging information for the browser console.
        console.log("CAPAD records:", capadRows.length);
        console.log(
            "CAPAD columns:",
            Object.keys(capadRows[0] || {})
        );
        console.log("Pest records:", pestRows.length);
        console.log(
            "Pest columns:",
            Object.keys(pestRows[0] || {})
        );

        status.textContent = "Building charts...";

        await loadChart1(capadRows);
        await loadChart2(capadRows);
        await loadChart3(capadRows);
        await loadChart4(pestRows);
        await loadChart5(pestRows);

        status.textContent = "All charts loaded successfully.";
    } catch (error) {
        console.error("Chart loading error:", error);

        status.textContent =
            "Error loading charts: " + error.message +
            " Check the browser Console (F12) for details.";
    }
}

main();
