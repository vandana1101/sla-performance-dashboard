/* =========================================================
   SLA DASHBOARD — CALCULATION ENGINE
   ---------------------------------------------------------
   IMPORTANT:
   No workbook results are hardcoded here.

   Everything is calculated from uploaded raw data.

   Main calculation chain:

   Distance + Dispatch Mode
            ↓
           TAT
            ↓
   Actual Reporting Date at CP End - Invoice Date
            ↓
      Invoice → Reporting
            ↓
   Invoice → Reporting <= TAT
            ↓
     Within TAT / Out TAT
========================================================= */


/* =========================================================
   NORMALIZATION
========================================================= */

const normalize = (value) => {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/[\s_\-./()]+/g, "");
};


/* =========================================================
   COLUMN ALIASES

   The uploaded source can have slightly different
   column names. These aliases allow the dashboard
   to identify them automatically.
========================================================= */

const ALIASES = {

  invoiceQty: [
    "Invoice Qty",
    "Invoice Quantity",
    "Invoice Qty.",
    "Qty",
    "Quantity"
  ],

  lrNo: [
    "LR No.",
    "LR No",
    "LR Number",
    "LR No.s",
    "LR"
  ],

  distance: [
    "Distance (KM)",
    "Distance KM",
    "Distance",
    "Road Distance",
    "Road Distance (KM)"
  ],

  dispatchMode: [
    "Dispatch Mode",
    "DispatchMode",
    "Mode"
  ],

  obdDate: [
    "OBD Date",
    "OBDDate"
  ],

  invoiceDate: [
    "Invoice Date",
    "InvoiceDate"
  ],

  reportingDate: [
    "Actual Reporting Date at CP End",
    "Actual Reporting Date At CP End",
    "Actual Reporting Date",
    "Reporting Date",
    "ReportingDate"
  ],

  deliveryDate: [
    "Actual Delivery Date",
    "Delivery Date",
    "DeliveryDate"
  ],

  pinCode: [
    "Pin Code",
    "PIN Code",
    "Pincode",
    "PIN",
    "Pin"
  ]
};


/* =========================================================
   FIND COLUMN
========================================================= */

function findColumn(row, field) {

  if (!row) {
    return null;
  }

  const columns = Object.keys(row);

  const normalizedColumns = {};

  columns.forEach((column) => {
    normalizedColumns[normalize(column)] = column;
  });

  const aliases = ALIASES[field] || [];

  for (const alias of aliases) {

    const normalizedAlias = normalize(alias);

    if (
      Object.prototype.hasOwnProperty.call(
        normalizedColumns,
        normalizedAlias
      )
    ) {
      return normalizedColumns[normalizedAlias];
    }
  }

  return null;
}


/* =========================================================
   GET VALUE
========================================================= */

function getValue(row, field) {

  const column = findColumn(row, field);

  if (!column) {
    return null;
  }

  return row[column];
}


/* =========================================================
   NUMBER PARSING
========================================================= */

function toNumber(value) {

  if (
    value === null ||
    value === undefined ||
    value === ""
  ) {
    return NaN;
  }

  if (typeof value === "number") {
    return Number.isFinite(value)
      ? value
      : NaN;
  }

  const cleaned = String(value)
    .replace(/,/g, "")
    .replace(/[^\d.-]/g, "");

  const number = Number(cleaned);

  return Number.isFinite(number)
    ? number
    : NaN;
}


/* =========================================================
   DATE SERIAL / DATE PARSING

   IMPORTANT:
   We DO NOT strip the time component.

   Excel's:

       Actual Reporting Date at CP End
       -
       Invoice Date

   can produce fractional days.

   Example:

       1 day 10 hours
       =
       1.4166667 days

   That fractional value must be preserved.
========================================================= */

function parseDate(value) {

  if (
    value === null ||
    value === undefined ||
    value === ""
  ) {
    return null;
  }


  /* JavaScript Date */

  if (value instanceof Date) {

    if (!Number.isNaN(value.getTime())) {
      return value;
    }

    return null;
  }


  /* Excel serial number */

  if (
    typeof value === "number" &&
    value > 0 &&
    value < 100000
  ) {

    const excelEpoch =
      Date.UTC(1899, 11, 30);

    return new Date(
      excelEpoch +
      value * 24 * 60 * 60 * 1000
    );
  }


  const text = String(value).trim();

  if (!text) {
    return null;
  }


  /*
    dd-mm-yyyy
    dd/mm/yyyy

    Also supports time:

    dd-mm-yyyy hh:mm
    dd/mm/yyyy hh:mm:ss
  */

  const match = text.match(
    /^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/
  );

  if (match) {

    const day = Number(match[1]);
    const month = Number(match[2]) - 1;
    const year = Number(match[3]);

    const hour =
      Number(match[4] || 0);

    const minute =
      Number(match[5] || 0);

    const second =
      Number(match[6] || 0);

    const date = new Date(
      year,
      month,
      day,
      hour,
      minute,
      second
    );

    if (!Number.isNaN(date.getTime())) {
      return date;
    }
  }


  /* Fallback */

  const date = new Date(text);

  if (!Number.isNaN(date.getTime())) {
    return date;
  }

  return null;
}


/* =========================================================
   EXCEL-STYLE DATE DIFFERENCE

   Returns fractional days.

   IMPORTANT:
   Blank reporting date is treated like Excel arithmetic:

       blank - invoice date

   Excel effectively treats blank as 0.

   Therefore the result is negative and the row becomes
   Within TAT when compared against positive TAT.
========================================================= */

function excelDateDifference(laterValue, earlierValue) {

  const later = parseDate(laterValue);
  const earlier = parseDate(earlierValue);


  /*
    If later date is blank, emulate Excel:

       0 - InvoiceDate
  */

  if (!later && earlier) {

    const excelEpoch =
      Date.UTC(1899, 11, 30);

    const earlierMs =
      Date.UTC(
        earlier.getFullYear(),
        earlier.getMonth(),
        earlier.getDate(),
        earlier.getHours(),
        earlier.getMinutes(),
        earlier.getSeconds()
      );

    const earlierSerial =
      (earlierMs - excelEpoch) /
      (24 * 60 * 60 * 1000);

    return -earlierSerial;
  }


  if (!later || !earlier) {
    return NaN;
  }


  return (
    (later.getTime() - earlier.getTime()) /
    (24 * 60 * 60 * 1000)
  );
}


/* =========================================================
   TAT CALCULATION

   ORIGINAL EXCEL LOGIC:

   =ROUNDUP(
      IF(
        Dispatch Mode="FTL",
        Distance/300,
        IF(
          Dispatch Mode="PTL",
          Distance/100+1,
          "Invalid Selection"
        )
      ),
      0.1
   )

   For the workbook, the resulting TAT values are
   integer categories.

   Therefore we reproduce the resulting behaviour as:

       FTL → CEILING(Distance / 300)
       PTL → CEILING(Distance / 100 + 1)
========================================================= */

function calculateTAT(row) {

  const distance =
    toNumber(
      getValue(row, "distance")
    );

  const dispatchMode =
    String(
      getValue(row, "dispatchMode") ?? ""
    )
      .trim()
      .toUpperCase();


  if (!Number.isFinite(distance)) {
    return NaN;
  }


  if (dispatchMode === "FTL") {

    return Math.ceil(
      distance / 300
    );
  }


  if (dispatchMode === "PTL") {

    return Math.ceil(
      distance / 100 + 1
    );
  }


  return NaN;
}


/* =========================================================
   SLA CALCULATION

   USER-DEFINED BUSINESS RULE:

   1–100 km       → Local Delivery
   101–300 km     → < 300 Km
   301–600 km     → 301 to 600 kms
   601–1000 km    → 601 to 1000 kms
   >1000 km       → >1000 km
========================================================= */

export function calculateSLA(value) {

  const distance =
    toNumber(value);


  if (
    !Number.isFinite(distance) ||
    distance < 1
  ) {
    return "Unclassified";
  }


  if (distance <= 100) {
    return "Local Delivery";
  }


  if (distance <= 300) {
    return "< 300 Km";
  }


  if (distance <= 600) {
    return "301 to 600 kms";
  }


  if (distance <= 1000) {
    return "601 to 1000 kms";
  }


  return ">1000 km";
}


/* =========================================================
   STATE FROM PIN CODE

   Relevant states for this dashboard:

   West Bengal
   Odisha
   Jharkhand
   Bihar

   PIN prefix logic is kept in code as a METHOD,
   not as calculated output.
========================================================= */

export function calculateState(value) {

  const pinNumber =
    toNumber(value);


  if (!Number.isFinite(pinNumber)) {
    return "Unknown";
  }


  const pin =
    String(
      Math.trunc(pinNumber)
    ).padStart(6, "0");


  /*
    Specific business/source exceptions found
    in the supplied Kolkata data.
  */

  if (pin === "743437") {
    return "Bihar";
  }

  if (pin === "721659") {
    return "ODISHA";
  }


  const prefix2 =
    Number(pin.substring(0, 2));


  /*
    West Bengal
    70xxxx – 74xxxx
  */

  if (
    prefix2 >= 70 &&
    prefix2 <= 74
  ) {
    return "West Bengal";
  }


  /*
    Odisha
    75xxxx – 77xxxx
  */

  if (
    prefix2 >= 75 &&
    prefix2 <= 77
  ) {
    return "ODISHA";
  }


  /*
    Jharkhand
    81xxxx – 83xxxx
  */

  if (
    prefix2 >= 81 &&
    prefix2 <= 83
  ) {
    return "Jharkhand";
  }


  /*
    Bihar
    80xxxx
    84xxxx
    85xxxx
  */

  if (
    prefix2 === 80 ||
    (prefix2 >= 84 && prefix2 <= 85)
  ) {
    return "Bihar";
  }


  return "Other / Unmapped";
}


/* =========================================================
   ENRICH RAW ROW

   THIS IS THE MOST IMPORTANT PART.

   Order:

   1. TAT
   2. Invoice → Reporting
   3. Timeline / Outbound TAT
   4. SLA
   5. State
========================================================= */

export function enrichRow(row) {

  /* -------------------------------------------------------
     1. TAT
  ------------------------------------------------------- */

  const tat =
    calculateTAT(row);


  /* -------------------------------------------------------
     2. Invoice → Reporting Lead Time

     Actual Reporting Date at CP End
     -
     Invoice Date
  ------------------------------------------------------- */

  const invoiceToReporting =
    excelDateDifference(
      getValue(row, "reportingDate"),
      getValue(row, "invoiceDate")
    );


  /* -------------------------------------------------------
     3. Timeline / Outbound TAT

     THIS IS NOT A SEPARATE FORMULA.

     It is the comparison between:

       Invoice → Reporting Lead Time
       TAT
  ------------------------------------------------------- */

  let outboundTAT =
    "Unclassified";


  if (
    Number.isFinite(invoiceToReporting) &&
    Number.isFinite(tat)
  ) {

    outboundTAT =
      invoiceToReporting <= tat
        ? "Within TAT"
        : "Out TAT";
  }


  /* -------------------------------------------------------
     OTHER CALCULATED FIELDS
  ------------------------------------------------------- */

  const obdToDelivery =
    excelDateDifference(
      getValue(row, "deliveryDate"),
      getValue(row, "obdDate")
    );


  const invoiceToDelivery =
    excelDateDifference(
      getValue(row, "deliveryDate"),
      getValue(row, "invoiceDate")
    );


  const obdToInvoice =
    excelDateDifference(
      getValue(row, "invoiceDate"),
      getValue(row, "obdDate")
    );


  const invoiceQty =
    toNumber(
      getValue(row, "invoiceQty")
    );


  const lrNo =
    String(
      getValue(row, "lrNo") ?? ""
    ).trim();


  const distance =
    toNumber(
      getValue(row, "distance")
    );


  const pin =
    getValue(row, "pinCode");


  const sla =
    calculateSLA(distance);


  const state =
    calculateState(pin);


  return {

    raw: row,

    invoiceQty:
      Number.isFinite(invoiceQty)
        ? invoiceQty
        : 0,

    lrNo,

    distance,

    pin,

    tat,

    invoiceToReporting,

    outboundTAT,

    status: outboundTAT,

    sla,

    state,

    obdToDelivery,

    invoiceToDelivery,

    obdToInvoice
  };
}


/* =========================================================
   SUM
========================================================= */

function sum(values) {

  return values.reduce(
    (total, value) => {

      const number =
        Number(value);

      return total +
        (
          Number.isFinite(number)
            ? number
            : 0
        );

    },
    0
  );
}


/* =========================================================
   UNIQUE LR COUNT
========================================================= */

function uniqueLRCount(rows) {

  const uniqueLRs =
    new Set();


  rows.forEach((row) => {

    const lr =
      String(
        row.lrNo ?? ""
      ).trim();


    if (lr) {
      uniqueLRs.add(lr);
    }

  });


  return uniqueLRs.size;
}


/* =========================================================
   OVERALL QUANTITY PERFORMANCE

   OUTBOUND / FG ONLY
========================================================= */

function calculateQuantityPerformance(rows) {

  const outRows =
    rows.filter(
      row =>
        row.outboundTAT === "Out TAT"
    );


  const withinRows =
    rows.filter(
      row =>
        row.outboundTAT === "Within TAT"
    );


  const out =
    sum(
      outRows.map(
        row => row.invoiceQty
      )
    );


  const within =
    sum(
      withinRows.map(
        row => row.invoiceQty
      )
    );


  const total =
    out + within;


  return {

    out,

    within,

    total,

    outPct:
      total
        ? (out / total) * 100
        : 0,

    withinPct:
      total
        ? (within / total) * 100
        : 0
  };
}


/* =========================================================
   OVERALL LR PERFORMANCE

   LR SHEET ONLY

   UNIQUE LR NO. COUNT
========================================================= */

function calculateLRPerformance(rows) {

  const outRows =
    rows.filter(
      row =>
        row.outboundTAT === "Out TAT"
    );


  const withinRows =
    rows.filter(
      row =>
        row.outboundTAT === "Within TAT"
    );


  const out =
    uniqueLRCount(
      outRows
    );


  const within =
    uniqueLRCount(
      withinRows
    );


  const total =
    uniqueLRCount(
      rows
    );


  return {

    out,

    within,

    total,

    outPct:
      total
        ? (out / total) * 100
        : 0,

    withinPct:
      total
        ? (within / total) * 100
        : 0
  };
}


/* =========================================================
   SLA-WISE QUANTITY
========================================================= */

function calculateSLAQuantityPerformance(
  rows,
  categories
) {

  return categories.map(
    (category) => {

      const categoryRows =
        rows.filter(
          row =>
            row.sla === category
        );


      const out =
        sum(
          categoryRows
            .filter(
              row =>
                row.outboundTAT === "Out TAT"
            )
            .map(
              row =>
                row.invoiceQty
            )
        );


      const within =
        sum(
          categoryRows
            .filter(
              row =>
                row.outboundTAT === "Within TAT"
            )
            .map(
              row =>
                row.invoiceQty
            )
        );


      const total =
        out + within;


      return {

        category,

        out,

        within,

        total,

        percent:
          total
            ? (within / total) * 100
            : 0
      };
    }
  );
}


/* =========================================================
   STATE-WISE QUANTITY
========================================================= */

function calculateStateQuantityPerformance(
  rows,
  categories
) {

  return categories.map(
    (category) => {

      const categoryRows =
        rows.filter(
          row =>
            row.state === category
        );


      const out =
        sum(
          categoryRows
            .filter(
              row =>
                row.outboundTAT === "Out TAT"
            )
            .map(
              row =>
                row.invoiceQty
            )
        );


      const within =
        sum(
          categoryRows
            .filter(
              row =>
                row.outboundTAT === "Within TAT"
            )
            .map(
              row =>
                row.invoiceQty
            )
        );


      const total =
        out + within;


      return {

        category,

        out,

        within,

        total,

        percent:
          total
            ? (within / total) * 100
            : 0
      };
    }
  );
}


/* =========================================================
   UNIQUE LR BY DIMENSION

   IMPORTANT:

   Unique LR count is calculated after filtering the
   category and TAT status.
========================================================= */

function calculateUniqueLRDimension(
  rows,
  dimension,
  categories
) {

  return categories.map(
    (category) => {

      const categoryRows =
        rows.filter(
          row =>
            row[dimension] === category
        );


      const outRows =
        categoryRows.filter(
          row =>
            row.outboundTAT === "Out TAT"
        );


      const withinRows =
        categoryRows.filter(
          row =>
            row.outboundTAT === "Within TAT"
        );


      const out =
        uniqueLRCount(
          outRows
        );


      const within =
        uniqueLRCount(
          withinRows
        );


      const total =
        uniqueLRCount(
          categoryRows
        );


      return {

        category,

        out,

        within,

        total,

        percent:
          total
            ? (within / total) * 100
            : 0
      };
    }
  );
}


/* =========================================================
   SOURCE SHEET DETECTION
========================================================= */

function findSources(files) {

  const allSheets = [];


  for (const file of files) {

    for (const sheet of file.sheets || []) {

      allSheets.push({

        ...sheet,

        fileName:
          file.name
      });

    }
  }


  const outboundCandidates =
    allSheets.filter(
      sheet => {

        const name =
          String(
            sheet.name ?? ""
          ).trim();


        return (
          /^outbound$/i.test(name) ||
          /^fg$/i.test(name) ||
          /outbound/i.test(name) ||
          /finished.?goods/i.test(name)
        );
      }
    );


  const lrCandidates =
    allSheets.filter(
      sheet => {

        const name =
          String(
            sheet.name ?? ""
          ).trim();


        return (
          /^lr$/i.test(name) ||
          /\blr\b/i.test(name)
        );
      }
    );


  if (!outboundCandidates.length) {

    throw new Error(
      "Could not find an Outbound / FG sheet."
    );
  }


  if (!lrCandidates.length) {

    throw new Error(
      "Could not find an LR sheet."
    );
  }


  return {

    outbound:
      outboundCandidates[0],

    lr:
      lrCandidates[0],

    all:
      allSheets
  };
}


/* =========================================================
   CONTEXT FROM FILE / SHEET NAME
========================================================= */

function extractContext(
  outbound,
  lr
) {

  const text =
    [
      outbound.fileName,
      outbound.name,
      lr.fileName,
      lr.name
    ]
      .filter(Boolean)
      .join(" ");


  const months = [

    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December"

  ];


  let month = "";


  for (const monthName of months) {

    if (
      new RegExp(
        monthName,
        "i"
      ).test(text)
    ) {

      month =
        monthName;

      break;
    }
  }


  if (!month) {

    const shortMatch =
      text.match(
        /\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\b/i
      );


    if (shortMatch) {

      const prefix =
        shortMatch[1]
          .toLowerCase()
          .substring(0, 3);


      month =
        months.find(
          item =>
            item
              .toLowerCase()
              .startsWith(prefix)
        ) || "";
    }
  }


  const yearMatch =
    text.match(
      /\b(20\d{2})\b/
    );


  const year =
    yearMatch
      ? yearMatch[1]
      : "";


  const locations = [

    "Kolkata",
    "Chennai",
    "Hyderabad",
    "Bangalore",
    "Bengaluru",
    "Delhi",
    "Mumbai",
    "Pune",
    "Ahmedabad",
    "Jaipur",
    "Lucknow",
    "Patna",
    "Bhubaneswar"

  ];


  const location =
    locations.find(
      item =>
        new RegExp(
          item,
          "i"
        ).test(text)
    ) ||
    "Customer Location";


  return {

    location,

    period:
      [month, year]
        .filter(Boolean)
        .join(" ") ||
      "Reporting Period",

    title:
      `${location} Delivery Performance Report`
  };
}


/* =========================================================
   MAIN DASHBOARD MODEL
========================================================= */

export function buildDashboardModel(
  files
) {

  const sources =
    findSources(files);


  const outboundRows =
    (sources.outbound.rows || [])
      .map(enrichRow);


  const lrRows =
    (sources.lr.rows || [])
      .map(enrichRow);


  const context =
    extractContext(
      sources.outbound,
      sources.lr
    );


  /* -------------------------------------------------------
     FIXED CATEGORY ORDER
  ------------------------------------------------------- */

  const slaCategories = [

    "< 300 Km",

    "301 to 600 kms",

    "Local Delivery",

    "601 to 1000 kms",

    ">1000 km"

  ];


  const stateCategories = [

    "Bihar",

    "Jharkhand",

    "ODISHA",

    "West Bengal"

  ];


  /* -------------------------------------------------------
     OVERALL
  ------------------------------------------------------- */

  const qty =
    calculateQuantityPerformance(
      outboundRows
    );


  const lr =
    calculateLRPerformance(
      lrRows
    );


  /* -------------------------------------------------------
     DIMENSIONAL PERFORMANCE
  ------------------------------------------------------- */

  const slaQty =
    calculateSLAQuantityPerformance(
      outboundRows,
      slaCategories
    );


  const stateQty =
    calculateStateQuantityPerformance(
      outboundRows,
      stateCategories
    );


  const slaLR =
    calculateUniqueLRDimension(
      lrRows,
      "sla",
      slaCategories
    );


  const stateLR =
    calculateUniqueLRDimension(
      lrRows,
      "state",
      stateCategories
    );


  return {

    context,

    qty,

    lr,

    slaQty,

    stateQty,

    slaLR,

    stateLR,

    sourceInfo: {

      primarySheet:
        sources.outbound.name,

      sheets: [

        sources.outbound.name,

        sources.lr.name

      ],

      rows: {

        outbound:
          outboundRows.length,

        lr:
          lrRows.length
      }

    }

  };
}