function parseVendorName(name) {

    name = name.trim();

    // 306_022_0650_FG01_V01
    let match = name.match(
        /^(.+?)_([a-zA-Z]+\d+)_((?:v|V)\d+)$/
    );

    if (match) {
        return {
            shot: match[1],
            element: match[2],
            version: match[3],
            name: name
        };
    }


    // 306_022_0650_FG01
    match = name.match(
        /^(.+?)_([a-zA-Z]+\d+)$/
    );

    if (match) {
        return {
            shot: match[1],
            element: match[2],
            version: null,
            name: name
        };
    }


    return null;
}


function normalizeShotName(value) {
    if (value === null || value === undefined) {
        return null;
    }

    let name = String(value).trim();

    if (!name) {
        return null;
    }

    // Accept standard shot names such as:
    // 306_001_0020
    // 306_022_0650
    let match = name.match(
        /^\d{3}_\d{3}_\d{4}$/
    );

    if (match) {
        return name;
    }

    // Also allow a shot name embedded in a larger value.
    match = name.match(
        /(\d{3}_\d{3}_\d{4})/
    );

    if (match) {
        return match[1];
    }

    return null;
}


function readExcelRequest(file) {

    return file.arrayBuffer().then(
        function (buffer) {

            const workbook = XLSX.read(
                buffer,
                {
                    type: "array"
                }
            );

            if (!workbook.SheetNames.length) {
                throw new Error(
                    "The Excel file contains no worksheets."
                );
            }

            const sheet =
                workbook.Sheets[
                    workbook.SheetNames[0]
                ];

            const rows =
                XLSX.utils.sheet_to_json(
                    sheet,
                    {
                        header: 1,
                        defval: ""
                    }
                );

            if (!rows.length) {
                throw new Error(
                    "The Excel request file is empty."
                );
            }

            const candidates = [];

            const columnCount =
                Math.max(
                    ...rows.map(
                        row => row.length
                    )
                );


            for (
                let column = 0;
                column < columnCount;
                column++
            ) {

                const values = [];

                for (
                    let row = 0;
                    row < rows.length;
                    row++
                ) {

                    const value =
                        rows[row][column];

                    const shot =
                        normalizeShotName(value);

                    if (shot) {
                        values.push(shot);
                    }
                }


                const uniqueValues =
                    [...new Set(values)];


                if (uniqueValues.length > 0) {

                    candidates.push({
                        column: column,
                        score: uniqueValues.length,
                        values: uniqueValues
                    });
                }
            }


            if (!candidates.length) {
                throw new Error(
                    "No valid shot names were found in the Excel request."
                );
            }


            candidates.sort(
                (a, b) => b.score - a.score
            );


            const highestScore =
                candidates[0].score;


            const bestCandidates =
                candidates.filter(
                    candidate =>
                        candidate.score === highestScore
                );


            if (bestCandidates.length > 1) {
                return {
                    ambiguous: true,
                    candidates: bestCandidates
                };
            }


            return {
                ambiguous: false,
                shots:
                    bestCandidates[0].values
            };
        }
    );
}

function readDIExcelRequest(file) {

    return file.arrayBuffer().then(
        function (buffer) {

            if (typeof XLSX === "undefined") {
                throw new Error(
                    "Excel reader library is not available."
                );
            }

            const workbook =
                XLSX.read(
                    buffer,
                    { type: "array" }
                );

            if (!workbook.SheetNames.length) {
                throw new Error(
                    "The DI Excel file contains no worksheets."
                );
            }

            const sheet =
                workbook.Sheets[
                    workbook.SheetNames[0]
                ];

            const rows =
                XLSX.utils.sheet_to_json(
                    sheet,
                    {
                        header: 1,
                        defval: ""
                    }
                );

            if (!rows.length) {
                throw new Error(
                    "The DI Excel file is empty."
                );
            }

            /*
             * The first row is the header row.
             * Do not treat any header value as a
             * requested DI name.
             */

            const header =
                rows[0].map(
                    function (value) {
                        return String(value)
                            .trim()
                            .toLowerCase();
                    }
                );

            const preferredNames = [
                "shot name",
                "shot",
                "shot code",
                "name",
                "filename",
                "file name"
            ];

            let nameColumn = -1;

            for (
                let i = 0;
                i < preferredNames.length;
                i++
            ) {

                const index =
                    header.indexOf(
                        preferredNames[i]
                    );

                if (index !== -1) {
                    nameColumn = index;
                    break;
                }
            }

            /*
             * If no recognizable header exists,
             * choose the column containing the
             * largest number of non-empty values
             * after the header row.
             */

            if (nameColumn === -1) {

                let bestColumn = -1;
                let bestCount = 0;

                const columnCount =
                    Math.max.apply(
                        null,
                        rows.map(
                            function (row) {
                                return row.length;
                            }
                        )
                    );

                for (
                    let column = 0;
                    column < columnCount;
                    column++
                ) {

                    let count = 0;

                    for (
                        let row = 1;
                        row < rows.length;
                        row++
                    ) {

                        const value =
                            String(
                                rows[row][column] || ""
                            ).trim();

                        if (value) {
                            count++;
                        }
                    }

                    if (count > bestCount) {
                        bestCount = count;
                        bestColumn = column;
                    }
                }

                nameColumn = bestColumn;
            }

            if (nameColumn === -1) {
                throw new Error(
                    "Could not determine the DI shot/name column."
                );
            }

            const names = [];

            /*
             * DI request files may be either:
             *
             * 1. Header-based Excel:
             *      Shot Name
             *      301_020_0240_comp_geo001_avid
             *
             * 2. Headerless single-column lists:
             *      301_020_0240_comp_geo001_avid
             *      301_020_0380_comp_geo001_avid
             *
             * If the first row itself looks like a DI name,
             * include row 0 instead of treating it as a header.
             */

            let firstDataRow = 1;

            const firstRowValue =
                rows[0] && rows[0][nameColumn] !== undefined
                    ? String(rows[0][nameColumn]).trim()
                    : "";

            const looksLikeDIName =
                /^\d{3}_\d{3}_\d{4}_.+$/i.test(
                    firstRowValue
                );

            if (looksLikeDIName) {
                firstDataRow = 0;
            }

            for (
                let row = firstDataRow;
                row < rows.length;
                row++
            ) {

                const value =
                    String(
                        rows[row][nameColumn] || ""
                    ).trim();

                if (value) {
                    names.push(value);
                }
            }

            if (!names.length) {
                throw new Error(
                    "The DI Excel file contains no shot names."
                );
            }

            return names;
        }
    );
}


function checkVendorTurnover(requestedShots, actualNames) {

    let sentByShot = {};

    let counts = {};

    actualNames.forEach(name => {

        counts[name] = (counts[name] || 0) + 1;

        let parsed = parseVendorName(name);

        if (!parsed) {
            return;
        }

        if (!sentByShot[parsed.shot]) {
            sentByShot[parsed.shot] = [];
        }

        sentByShot[parsed.shot].push(parsed);

    });


    let requested = new Set(requestedShots);

    let sentShots = new Set(
        Object.keys(sentByShot)
    );


    let missing = [...requested]
        .filter(
            shot => !sentShots.has(shot)
        );


    let unrequested = [...sentShots]
        .filter(
            shot => !requested.has(shot)
        );


    let duplicates = Object.keys(counts)
        .filter(
            name => counts[name] > 1
        );


    let output = "";

    output += "==============================\n";
    output += "       VFX TURNOVER QC\n";
    output += "==============================\n\n";


    output += `Requested shots: ${requested.size}\n`;
    output += `Shots found:     ${requested.size - missing.length}\n`;
    output += `Missing shots:   ${missing.length}\n`;
    output += `Unrequested:     ${unrequested.length}\n`;
    output += `Duplicates:      ${duplicates.length}\n`;


    output += "\n------------------------------\n";
    output += "SHOT DETAILS\n";
    output += "------------------------------\n";


    requestedShots.forEach(shot => {

        output += `\n${shot}\n`;

        if (!sentByShot[shot]) {

            output += "   MISSING\n";

            return;
        }


        sentByShot[shot].forEach(item => {

            if (item.version) {

                output +=
                    `   OK  ${item.element}_${item.version}\n`;

            } else {

                output +=
                    `   OK  ${item.element}\n`;
            }

        });

    });


    if (unrequested.length > 0) {

        output += "\n------------------------------\n";
        output += "UNREQUESTED SHOTS\n";
        output += "------------------------------\n";

        unrequested.forEach(shot => {

            output +=
                `   UNREQUESTED  ${shot}\n`;

        });

    }


    output += "\n==============================\n";


    window.lastVendorQCData = {
        requested: requested.size,
        found: requested.size - missing.length,
        missing: missing.slice(),
        unrequested: unrequested.slice(),
        duplicates: duplicates.slice(),
        sentByShot: sentByShot
    };

    return output;
}



function checkDISend(requestedNames, actualNames) {

    /*
     * DI comparison normalization.
     *
     * The same DI name may appear as:
     *
     *   308_002_0010_comp_mrz002
     *
     * or:
     *
     *   308_002_0010_COMP_MRZ002_AVID
     *
     * Treat those as the same name.
     *
     * Vendor comparison is completely separate.
     */

    function normalizeDIName(name) {

        return String(name || "")
            .trim()
            .replace(/_AVID$/i, "")
            .toLowerCase();
    }


    const requestedCounts = {};
    const actualCounts = {};
    const actualDisplayNames = {};

    requestedNames.forEach(function(name) {

        const key =
            normalizeDIName(name);

        requestedCounts[key] =
            (requestedCounts[key] || 0) + 1;
    });


    actualNames.forEach(function(name) {

        const key =
            normalizeDIName(name);

        actualCounts[key] =
            (actualCounts[key] || 0) + 1;

        if (!actualDisplayNames[key]) {
            actualDisplayNames[key] = name;
        }
    });


    const requestedUnique =
        Object.keys(requestedCounts);

    const actualUnique =
        Object.keys(actualCounts);


    const missing = [];
    const extra = [];
    const countMismatch = [];


    requestedUnique.forEach(function(key) {

        const requestedCount =
            requestedCounts[key];

        const actualCount =
            actualCounts[key] || 0;

        if (actualCount === 0) {

            missing.push({
                name: key,
                requested: requestedCount
            });

        } else if (
            actualCount !== requestedCount
        ) {

            countMismatch.push({
                name: key,
                requested: requestedCount,
                actual: actualCount
            });
        }
    });


    actualUnique.forEach(function(key) {

        if (!requestedCounts[key]) {

            extra.push({
                name: actualDisplayNames[key] || key,
                actual: actualCounts[key]
            });
        }
    });


    const passed =
        missing.length === 0 &&
        extra.length === 0 &&
        countMismatch.length === 0;


    let output = "";

    output += "==============================\n";
    output += "          DI SEND QC\n";
    output += "==============================\n\n";

    output +=
        `Requested entries: ${requestedNames.length}\n`;

    output +=
        `Actual entries:    ${actualNames.length}\n`;

    output +=
        `Missing names:     ${missing.length}\n`;

    output +=
        `Extra names:       ${extra.length}\n`;

    output +=
        `Count mismatches:  ${countMismatch.length}\n`;

    output +=
        `Status:            ${passed ? "PASSED" : "ISSUES FOUND"}\n`;


    output += "\n------------------------------\n";
    output += "REQUESTED DI SEND NAMES\n";
    output += "------------------------------\n";


    requestedNames.forEach(function(name) {

        const key =
            normalizeDIName(name);

        const requestedCount =
            requestedCounts[key];

        const actualCount =
            actualCounts[key] || 0;

        output += `\n${name}\n`;

        if (actualCount === requestedCount) {

            output +=
                `   MATCH (${actualCount})\n`;

        } else if (actualCount === 0) {

            output +=
                "   MISSING\n";

        } else {

            output +=
                `   COUNT MISMATCH — REQUESTED ${requestedCount}, ACTUAL ${actualCount}\n`;
        }
    });


    if (countMismatch.length > 0) {

        output += "\n------------------------------\n";
        output += "COUNT MISMATCHES\n";
        output += "------------------------------\n";

        countMismatch.forEach(function(item) {

            output +=
                `\n${item.name}\n`;

            output +=
                `   REQUESTED: ${item.requested}\n`;

            output +=
                `   ACTUAL:    ${item.actual}\n`;
        });
    }


    if (missing.length > 0) {

        output += "\n------------------------------\n";
        output += "MISSING NAMES\n";
        output += "------------------------------\n";

        missing.forEach(function(item) {

            output +=
                `\n${item.name}\n`;
        });
    }


    if (extra.length > 0) {

        output += "\n------------------------------\n";
        output += "EXTRA NAMES\n";
        output += "------------------------------\n";

        extra.forEach(function(item) {

            output +=
                `\n${item.name}\n`;

            output +=
                `   EXTRA (${item.actual})\n`;
        });
    }


    output += "\n==============================\n";


    window.lastDIQCData = {
        requested: requestedNames.length,
        actual: actualNames.length,
        missing: missing.slice(),
        extra: extra.slice(),
        countMismatch: countMismatch.slice(),
        passed: passed,
        requestedCounts: {},
        actualCounts: {}
    };


    requestedNames.forEach(function(name) {

        const key =
            normalizeDIName(name);

        window.lastDIQCData.requestedCounts[key] =
            (window.lastDIQCData.requestedCounts[key] || 0) + 1;
    });


    actualNames.forEach(function(name) {

        const key =
            normalizeDIName(name);

        window.lastDIQCData.actualCounts[key] =
            (window.lastDIQCData.actualCounts[key] || 0) + 1;
    });


    return output;
}

function readActualFile(file) {

    return file.text().then(
        function (text) {

            const lines =
                text.split(/\r\n|\n|\r/);

            const extension =
                file.name
                    .toLowerCase()
                    .split(".")
                    .pop();

            // -----------------------------------------
            // EDL
            // -----------------------------------------

            if (extension === "edl") {

                const names = [];
                let locLinesFound = 0;

                // -----------------------------------------
                // PRIMARY: *LOC: marker names
                // -----------------------------------------

                lines.forEach(function(line) {

                    const trimmed =
                        line.trim();

                    if (
                        !trimmed
                            .toUpperCase()
                            .startsWith("*LOC:")
                    ) {
                        return;
                    }

                    locLinesFound++;

                    const parts =
                        trimmed.split(/\s+/);

                    if (parts.length >= 4) {

                        const name =
                            parts[parts.length - 1]
                                .trim();

                        /*
                         * Only treat the LOC value as a DI
                         * name if it has the expected shot/name
                         * structure.
                         *
                         * This prevents unrelated editorial
                         * LOC notes such as:
                         *
                         * WHITE DAILIES SCALING DISCREPANCY
                         *
                         * from blocking the FROM CLIP NAME
                         * fallback.
                         */

                        if (
                            /^[0-9]+_[0-9]+_[0-9]+_[A-Za-z0-9]+(?:_[A-Za-z0-9]+)*$/i
                                .test(name)
                        ) {
                            names.push(name);
                        }
                    }
                });

                // If usable marker names were found,
                // use them and do not inspect clip names.
                if (names.length > 0) {
                    return names;
                }

                // -----------------------------------------
                // FALLBACK: *FROM CLIP NAME:
                // -----------------------------------------

                const clipNames = [];

                lines.forEach(function(line) {

                    const trimmed =
                        line.trim();

                    if (
                        !trimmed
                            .toUpperCase()
                            .startsWith("*FROM CLIP NAME:")
                    ) {
                        return;
                    }

                    let name =
                        trimmed
                            .slice(
                                trimmed.indexOf(":") + 1
                            )
                            .trim();

                    if (name) {
                        clipNames.push(name);
                    }
                });

                if (clipNames.length > 0) {
                    return clipNames;
                }

                // -----------------------------------------
                // NOTHING USABLE FOUND
                // -----------------------------------------

                if (locLinesFound > 0) {
                    throw new Error(
                        "*LOC: markers were found, but no marker names " +
                        "could be read from the EDL. No usable *FROM CLIP NAME: " +
                        "fallback names were found either."
                    );
                }

                throw new Error(
                    "Could not find DI SEND names in this EDL. " +
                    "Expected either *LOC: marker names or " +
                    "*FROM CLIP NAME: entries."
                );
            }

            // -----------------------------------------
            // ALE / TAB
            // -----------------------------------------

            /*
             * ALE files have a Column section followed by
             * the actual column header row, then a Data section.
             *
             * We specifically locate the Name column and then
             * read that column from the Data rows.
             */

            function normalizeColumnName(value) {

                if (
                    value === null ||
                    value === undefined
                ) {
                    return "";
                }

                return String(value)
                    .trim()
                    .toLowerCase()
                    .replace(/[_-]+/g, " ")
                    .replace(/\s+/g, " ");
            }


            let headerIndex = -1;
            let dataIndex = -1;

            /*
             * Find the actual "Data" marker.
             */

            for (
                let index = 0;
                index < lines.length;
                index++
            ) {

                if (
                    lines[index]
                        .replace(/^\uFEFF/, "")
                        .trim()
                        .toLowerCase() === "data"
                ) {
                    dataIndex = index;
                    break;
                }
            }


            /*
             * Find the header row immediately before Data.
             */

            if (dataIndex > 0) {

                for (
                    let index = dataIndex - 1;
                    index >= 0;
                    index--
                ) {

                    const candidate =
                        lines[index]
                            .split("\t");

                    const normalized =
                        candidate.map(
                            normalizeColumnName
                        );

                    if (
                        normalized.some(
                            value => value === "name"
                        )
                    ) {
                        headerIndex = index;
                        break;
                    }
                }
            }


            if (
                headerIndex === -1 ||
                dataIndex === -1
            ) {

                /*
                 * VENDOR ALE FALLBACK
                 *
                 * Some ALE variants do not expose the
                 * Name/Data structure exactly as expected.
                 * Scan the tab-delimited rows for a column
                 * containing recognizable Vendor names.
                 *
                 * Do NOT modify the names. The existing
                 * Vendor QC parser handles them afterward.
                 */

                const fallbackNames = [];

                for (
                    let row = 0;
                    row < lines.length;
                    row++
                ) {

                    const columns =
                        lines[row].split("\t");

                    if (columns.length < 2) {
                        continue;
                    }

                    for (
                        let column = 0;
                        column < columns.length;
                        column++
                    ) {

                        const value =
                            String(
                                columns[column] || ""
                            )
                            .replace(/^\uFEFF/, "")
                            .trim();

                        if (
                            /^\d{3}_\d{3}_\d{4}_[A-Za-z0-9]+(?:_[A-Za-z0-9]+)*$/
                                .test(value)
                        ) {
                            fallbackNames.push(value);
                        }
                    }
                }

                if (fallbackNames.length > 0) {
                    return fallbackNames;
                }

                throw new Error(
                    "Could not find Vendor names in the ALE/TAB file."
                );
            }


            const headerColumns =
                lines[headerIndex]
                    .split("\t");

            const normalizedHeaders =
                headerColumns.map(
                    normalizeColumnName
                );

            const nameIndex =
                normalizedHeaders.indexOf("name");


            if (nameIndex === -1) {

                throw new Error(
                    "The ALE/TAB file contains a header row, " +
                    "but no Name column could be identified."
                );
            }


            const names = [];


            /*
             * Read ONLY rows after the Data marker.
             */

            for (
                let index = dataIndex + 1;
                index < lines.length;
                index++
            ) {

                const line =
                    lines[index];

                if (!line.trim()) {
                    continue;
                }

                const values =
                    line.split("\t");

                if (
                    values.length <= nameIndex
                ) {
                    continue;
                }

                const name =
                    values[nameIndex]
                        .trim();

                /*
                 * Only accept recognizable VFX
                 * turnover names.
                 *
                 * Examples:
                 * 305_008_0110_FG01
                 * 305_008_0070_FG01
                 * 305_005_0680_FG01
                 */

                if (
                    /^[A-Za-z0-9]+(?:_[A-Za-z0-9]+)+_[A-Za-z]+\d+(?:_[A-Za-z0-9]+)*$/
                        .test(name)
                ) {
                    names.push(name);
                }
            }


            if (names.length === 0) {

                throw new Error(
                    "The ALE/TAB file contains a Name column and Data section, " +
                    "but no recognizable turnover names were found."
                );
            }


            return names;
        }
    );
}


/* DI EXACT CLIP NAME FALLBACK - 2026-10-01 */

/*
 * DI EDL fallback:
 * If an EDL has no usable *LOC: marker names,
 * read the exact *FROM CLIP NAME: value instead.
 *
 * IMPORTANT:
 * Do NOT modify the name.
 * The DI request must match the DI send name 1:1.
 */

const originalReadActualFile_DI =
    readActualFile;

readActualFile =
    function(file) {

        const extension =
            file.name
                .toLowerCase()
                .split(".")
                .pop();

        if (extension !== "edl") {
            return originalReadActualFile_DI(file);
        }

        return file.text().then(
            function(text) {

                const lines =
                    text.split(/\r\n|\n|\r/);

                let hasUsableLOC =
                    false;

                lines.forEach(function(line) {

                    const trimmed =
                        line.trim();

                    if (
                        trimmed
                            .toUpperCase()
                            .startsWith("*LOC:")
                    ) {

                        const value =
                            trimmed
                                .slice(
                                    trimmed.indexOf(":") + 1
                                )
                                .trim();

                        if (value) {
                            hasUsableLOC = true;
                        }
                    }
                });

                /*
                 * If the EDL has usable LOC markers,
                 * let the existing parser handle it.
                 */
                if (hasUsableLOC) {
                    return originalReadActualFile_DI(file);
                }

                /*
                 * No LOC markers.
                 * Read exact FROM CLIP NAME values.
                 */
                const clipNames = [];

                lines.forEach(function(line) {

                    const trimmed =
                        line.trim();

                    if (
                        !trimmed
                            .toUpperCase()
                            .startsWith("*FROM CLIP NAME:")
                    ) {
                        return;
                    }

                    const name =
                        trimmed
                            .slice(
                                trimmed.indexOf(":") + 1
                            )
                            .trim();

                    if (name) {
                        clipNames.push(name);
                    }
                });

                if (clipNames.length > 0) {
                    return clipNames;
                }

                /*
                 * Neither naming method worked.
                 * Let the existing error handling report it.
                 */
                return originalReadActualFile_DI(file);
            }
        );
    };

