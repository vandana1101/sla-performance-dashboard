import React from "react";

import {
  BarChart3,
  Package,
  Truck,
  MapPin,
  Ruler,
  Clock3,
  Pencil,
} from "lucide-react";

import {
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  ResponsiveContainer,
  LabelList,
} from "recharts";

/* =========================================================
   FORMATTERS
========================================================= */

const nf = new Intl.NumberFormat("en-IN");

const pct = (value) =>
  `${Number(value || 0).toFixed(2)}%`;

/* =========================================================
   KPI CARD
========================================================= */

function Kpi({
  label,
  value,
  sub,
  icon: Icon,
  accent,
}) {
  return (
    <div className={`kpi-card ${accent || ""}`}>

      <div className="kpi-icon">
        <Icon size={19} />
      </div>

      <div className="kpi-label">
        {label}
      </div>

      <div className="kpi-value">
        {value}
      </div>

      <div className="kpi-sub">
        {sub}
      </div>

    </div>
  );
}

/* =========================================================
   STANDARD PERFORMANCE TABLE

   Used for:
   - SLA / Quantity
   - SLA / LR
   - State / Quantity
   - State / LR
========================================================= */

function PerformanceTable({
  title,
  icon: Icon,
  rows,
  measureLabel,
}) {

  const outTotal = (rows || []).reduce(
    (sum, row) =>
      sum + Number(row.out || 0),
    0
  );

  const withinTotal = (rows || []).reduce(
    (sum, row) =>
      sum + Number(row.within || 0),
    0
  );

  const grandTotal =
    outTotal + withinTotal;

  const overallPct =
    grandTotal > 0
      ? (withinTotal / grandTotal) * 100
      : 0;

  return (
    <section className="panel">

      <div className="panel-head">

        <div className="panel-title">

          <div className="panel-icon">
            <Icon size={17} />
          </div>

          <h3>
            {title}
          </h3>

        </div>

        <span className="measure-badge">
          {measureLabel}
        </span>

      </div>

      <div className="table-wrap">

        <table>

          <thead>

            <tr>

              <th>
                CATEGORY
              </th>

              <th>
                OUT TAT
              </th>

              <th>
                WITHIN TAT
              </th>

              <th>
                TOTAL
              </th>

              <th>
                DELIVERY TAT %
              </th>

            </tr>

          </thead>

          <tbody>

            {(rows || []).map(
              (row) => (

                <tr
                  key={row.category}
                >

                  <td>
                    <strong>
                      {row.category}
                    </strong>
                  </td>

                  <td>
                    {nf.format(
                      row.out || 0
                    )}
                  </td>

                  <td>
                    {nf.format(
                      row.within || 0
                    )}
                  </td>

                  <td>
                    {nf.format(
                      row.total || 0
                    )}
                  </td>

                  <td>

                    <span
                      className={`percent ${
                        row.percent >= 80
                          ? "good"
                          : row.percent >= 60
                          ? "mid"
                          : "bad"
                      }`}
                    >
                      {pct(row.percent)}
                    </span>

                  </td>

                </tr>

              )
            )}

            <tr className="total-row">

              <td>
                Grand Total
              </td>

              <td>
                {nf.format(outTotal)}
              </td>

              <td>
                {nf.format(withinTotal)}
              </td>

              <td>
                {nf.format(grandTotal)}
              </td>

              <td>
                {pct(overallPct)}
              </td>

            </tr>

          </tbody>

        </table>

      </div>

    </section>
  );
}

/* =========================================================
   TAT-WISE PERFORMANCE TABLE

   Rows:
   - Out TAT
   - Within TAT
   - Grand Total

   No TAT 1 / TAT 2.
========================================================= */

function TatPerformanceTable({
  title,
  measureLabel,
  outValue,
  withinValue,
}) {

  const out =
    Number(outValue || 0);

  const within =
    Number(withinValue || 0);

  const total =
    out + within;

  const outPct =
    total > 0
      ? (out / total) * 100
      : 0;

  const withinPct =
    total > 0
      ? (within / total) * 100
      : 0;

  return (
    <section className="panel">

      {/* CARD HEADER */}

      <div className="panel-head">

        <div className="panel-title">

          <div className="panel-icon">
            <Clock3 size={17} />
          </div>

          <h3>
            {title}
          </h3>

        </div>

        <span className="measure-badge">
          {measureLabel}
        </span>

      </div>

      {/* TABLE */}

      <div className="table-wrap">

        <table>

          <thead>

            <tr>

              <th>
                CATEGORY
              </th>

              <th>
                {measureLabel}
              </th>

              <th>
                DELIVERY TAT %
              </th>

            </tr>

          </thead>

          <tbody>

            {/* OUT TAT */}

            <tr>

              <td>
                <strong>
                  Out TAT
                </strong>
              </td>

              <td>
                {nf.format(out)}
              </td>

              <td>

                <span className="percent bad">
                  {outPct.toFixed(2)}%
                </span>

              </td>

            </tr>

            {/* WITHIN TAT */}

            <tr>

              <td>
                <strong>
                  Within TAT
                </strong>
              </td>

              <td>
                {nf.format(within)}
              </td>

              <td>

                <span className="percent good">
                  {withinPct.toFixed(2)}%
                </span>

              </td>

            </tr>

            {/* GRAND TOTAL */}

            <tr className="total-row">

              <td>
                Grand Total
              </td>

              <td>
                {nf.format(total)}
              </td>

              <td>
                100.00%
              </td>

            </tr>

          </tbody>

        </table>

      </div>

    </section>
  );
}

/* =========================================================
   MAIN DASHBOARD
========================================================= */

export default function DashboardPage({
  model,
  onEditData,
}) {

  /* =======================================================
     EMPTY STATE
  ======================================================= */

  if (!model) {
    return null;
  }

  /* =======================================================
     ERROR
  ======================================================= */

  if (model.error) {

    return (
      <div className="error-box">
        {model.error}
      </div>
    );

  }

  /* =======================================================
     MODEL DATA
  ======================================================= */

  const {
    context,
    qty,
    lr,
    slaQty,
    stateQty,
    slaLR,
    stateLR,
    sourceInfo,
  } = model;

  /* =======================================================
     OVERALL TAT VALUES

     Quantity:
       Outbound / FG
       → Invoice Qty

     LR:
       LR sheet
       → Unique LR No.
  ======================================================= */

  const outQty =
    Number(qty?.out || 0);

  const withinQty =
    Number(qty?.within || 0);

  const outLR =
    Number(lr?.out || 0);

  const withinLR =
    Number(lr?.within || 0);

  /* =======================================================
     SLA GRAPH DATA
  ======================================================= */

  const chartData =
    (slaQty || []).map(
      (row) => ({

        name:
          String(row.category)
            .replace(
              "Local Delivery",
              "Local"
            )
            .replace(
              "< 300 Km",
              "< 300"
            )
            .replace(
              "301 to 600 kms",
              "301–600"
            )
            .replace(
              "601 to 1000 kms",
              "601–1000"
            ),

        "Within TAT":
          Number(
            row.within || 0
          ),

        "Out TAT":
          Number(
            row.out || 0
          ),

      })
    );

  /* =======================================================
     DASHBOARD
  ======================================================= */

  return (

    <main className="dashboard">

      {/* ===================================================
          DASHBOARD HERO
      =================================================== */}

      <div className="dashboard-hero">

        <div>

          <div className="eyebrow">
            DELIVERY PERFORMANCE REPORT
          </div>

          <h1>
            {context.title}
          </h1>

          <div className="source-line">

            <span>
              Calculated from{" "}
              <strong>
                {sourceInfo.primarySheet}
              </strong>
            </span>

            <span>
              •
            </span>

            <span>
              {context.location}
            </span>

            <span>
              •
            </span>

            <span>
              {context.period}
            </span>

          </div>

          {/* =================================================
              EDIT DATA BUTTON
          ================================================= */}

          {onEditData && (
            <button
              type="button"
              className="dashboard-edit-data-btn"
              onClick={onEditData}
            >
              <Pencil size={15} />
              Edit Data
            </button>
          )}

        </div>

        {/* SOURCE INFORMATION */}

        <div className="context-card">

          <div className="context-label">
            DATA SOURCES
          </div>

          <div className="source-tags">

            {sourceInfo.sheets.map(
              (sheet) => (

                <span
                  key={sheet}
                  className="source-tag"
                >
                  {sheet}
                </span>

              )
            )}

          </div>

          <div className="context-small">
            Runtime calculations · Format 1
          </div>

        </div>

      </div>

      {/* ===================================================
          KPI CARDS
      =================================================== */}

      <div className="kpi-grid">

        <Kpi
          label="Invoice Quantity"
          value={nf.format(
            qty?.total || 0
          )}
          sub={`${pct(
            qty?.withinPct || 0
          )} within TAT`}
          icon={Package}
          accent="blue"
        />

        <Kpi
          label="Within TAT Qty"
          value={nf.format(
            withinQty
          )}
          sub={`${pct(
            qty?.withinPct || 0
          )} of invoice qty`}
          icon={Clock3}
          accent="green"
        />

        <Kpi
          label="Out TAT Qty"
          value={nf.format(
            outQty
          )}
          sub={`${pct(
            qty?.outPct || 0
          )} of invoice qty`}
          icon={Clock3}
          accent="orange"
        />

        <Kpi
          label="Unique LR Count"
          value={nf.format(
            lr?.total || 0
          )}
          sub={`${pct(
            lr?.withinPct || 0
          )} within TAT`}
          icon={Truck}
          accent="purple"
        />

      </div>

      {/* ===================================================
          SLA PERFORMANCE GRAPH
      =================================================== */}

      <section className="panel chart-panel">

        <div className="panel-head">

          <div className="panel-title">

            <div className="panel-icon">
              <BarChart3 size={17} />
            </div>

            <h3>
              SLA performance snapshot
            </h3>

          </div>

          <span className="measure-badge">
            Invoice Qty
          </span>

        </div>

        <div className="chart-box">

          <ResponsiveContainer
            width="100%"
            height={300}
          >

            <BarChart
              data={chartData}
              margin={{
                top: 28,
                right: 20,
                left: 0,
                bottom: 5,
              }}
              barGap={5}
            >

              <CartesianGrid
                strokeDasharray="3 3"
                vertical={false}
              />

              <XAxis
                dataKey="name"
                tick={{
                  fontSize: 11,
                }}
              />

              <YAxis
                tick={{
                  fontSize: 10,
                }}
              />

              <Tooltip
                formatter={(value) =>
                  nf.format(value)
                }
              />

              {/* WITHIN TAT */}

              <Bar
                dataKey="Within TAT"
                fill="#16a34a"
                radius={[
                  5,
                  5,
                  0,
                  0,
                ]}
              >

                <LabelList
                  dataKey="Within TAT"
                  position="top"
                  formatter={(value) =>
                    nf.format(value)
                  }
                  style={{
                    fontSize: 10,
                    fontWeight: 700,
                    fill: "#16834b",
                  }}
                />

              </Bar>

              {/* OUT TAT */}

              <Bar
                dataKey="Out TAT"
                fill="#f97316"
                radius={[
                  5,
                  5,
                  0,
                  0,
                ]}
              >

                <LabelList
                  dataKey="Out TAT"
                  position="top"
                  formatter={(value) =>
                    nf.format(value)
                  }
                  style={{
                    fontSize: 10,
                    fontWeight: 700,
                    fill: "#ea7a25",
                  }}
                />

              </Bar>

            </BarChart>

          </ResponsiveContainer>

        </div>

      </section>

      {/* ===================================================
          01 — TAT-WISE PERFORMANCE
      =================================================== */}

      <div className="section-heading">

        <span>
          01
        </span>

        <div>

          <h2>
            TAT-wise performance
          </h2>

          <p>
            Delivery TAT performance shown in
            quantity and LR-wise views.
          </p>

        </div>

      </div>

      {/* ===================================================
          TAT TABLES — SIDE BY SIDE
      =================================================== */}

      <div className="two-col">

        {/* TAT / QUANTITY */}

        <TatPerformanceTable
          title="SLA / Quantity"
          measureLabel="Invoice Qty"
          outValue={outQty}
          withinValue={withinQty}
        />

        {/* TAT / LR */}

        <TatPerformanceTable
          title="SLA / LR"
          measureLabel="Unique LR"
          outValue={outLR}
          withinValue={withinLR}
        />

      </div>

      {/* ===================================================
          02 — SLA-WISE PERFORMANCE
      =================================================== */}

      <div className="section-heading">

        <span>
          02
        </span>

        <div>

          <h2>
            SLA-wise performance
          </h2>

          <p>
            Delivery TAT split by distance
            category.
          </p>

        </div>

      </div>

      <div className="two-col">

        {/* SLA / QUANTITY */}

        <PerformanceTable
          title="SLA / Quantity"
          icon={Ruler}
          rows={slaQty || []}
          measureLabel="Invoice Qty"
        />

        {/* SLA / LR */}

        <PerformanceTable
          title="SLA / LR"
          icon={Ruler}
          rows={slaLR || []}
          measureLabel="Unique LR"
        />

      </div>

      {/* ===================================================
          03 — STATE-WISE PERFORMANCE
      =================================================== */}

      <div className="section-heading">

        <span>
          03
        </span>

        <div>

          <h2>
            State-wise performance
          </h2>

          <p>
            Delivery TAT split using state
            derived from PIN code.
          </p>

        </div>

      </div>

      <div className="two-col">

        {/* STATE / QUANTITY */}

        <PerformanceTable
          title="State / Quantity"
          icon={MapPin}
          rows={stateQty || []}
          measureLabel="Invoice Qty"
        />

        {/* STATE / LR */}

        <PerformanceTable
          title="State / LR"
          icon={MapPin}
          rows={stateLR || []}
          measureLabel="Unique LR"
        />

      </div>

      {/* ===================================================
          FOOTER
      =================================================== */}

      <footer className="dashboard-footer">

        <div>

          <Clock3 size={16} />

          <span>
            All calculated fields are generated
            at runtime from uploaded raw data.
          </span>

        </div>

        <div>

          Source sheets:{" "}
          {sourceInfo.sheets.join(" · ")}

        </div>

      </footer>

    </main>
  );
}