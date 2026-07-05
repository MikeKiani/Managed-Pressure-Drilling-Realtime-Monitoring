# MPD Real-Time Monitoring — Power BI Analysis, Connection & Cleaning Code

This document explains the four monitoring charts in your `MPD.pbix`, the issues
found in the current build, and gives you ready-to-paste **Power Query (M)** code
to connect + clean the rig data, plus corrected **DAX** measures so every chart
aggregates correctly.

**Data model found in your file:** a single table named `MPD`, keyed by a
30-minute time bucket column `TimeBucket30min`, with these channels:

| Column | Meaning | Unit |
|---|---|---|
| `HoleDepth_m` | Total measured depth drilled | m |
| `BitDepth_m` | Current bit position | m |
| `ROP_m_per_hr` | Rate of Penetration | m/hr |
| `MudDensity_kg_m3` | Static mud weight (in) | kg/m³ |
| `PWDECD_kg_m3` | Equivalent Circulating Density (Pressure While Drilling) | kg/m³ |
| `FlowIn_m3min` | Mud pumped downhole | m³/min |
| `PumpOut_m3min` | Return / flow-out | m³/min |
| `StandpipePressure_kPa` | Standpipe pressure (SPP) | kPa |
| `ActualBackPressure_kPa` | Applied surface backpressure | kPa |
| `TargetBackPressure_kPa` | Backpressure setpoint | kPa |
| `FlowlineTemperature_C` | Return-line mud temperature | °C |
| `ShakerGas_pct` | Gas at the shale shakers | % |
| `Net Gain/Loss (m3)` | Pit / trip-tank volume balance (measure) | m³ |

---

## 1) Hole Depth, Bit Depth & ROP

**What it shows.** The operational state of the well and drilling efficiency.
- **Hole Depth** = how deep the hole is in total.
- **Bit Depth** = where the bit is right now. When `BitDepth ≈ HoleDepth` the bit
  is *on bottom* and actually making hole; when `BitDepth < HoleDepth` the bit is
  *off bottom* (tripping, reaming, or circulating).
- **ROP** = how fast new hole is being made. ROP is only meaningful *on bottom*.

**Why it matters for MPD.** It tells the monitoring engineer, at a glance, whether
the rig is drilling, tripping, or circulating — which changes how every other
pressure and flow signal should be interpreted.

**Chart type in your file:** line + clustered-column combo. ROP is the column,
Bit Depth is the line (secondary axis).

**⚠ Issues in the current build**
- **Hole Depth is not on the chart** even though the title says so. Add
  `HoleDepth_m` as a second line so `HoleDepth` vs `BitDepth` shows on/off-bottom.
- The X-axis uses the auto **Date Hierarchy → Year** level, which collapses all
  points. Switch the axis to the continuous `TimeBucket30min` datetime
  (see *Visual fixes*, section 6).

---

## 2) PWD ECD & Mud Density with Flow In / Pump Out

**What it shows.** The core MPD control picture: pressure regime vs hydraulics.
- **Mud Density** = the *static* mud weight going in.
- **PWD ECD** (Equivalent Circulating Density) = the *effective* density the
  formation actually feels while circulating (static weight + annular friction +
  cuttings load). ECD is normally **higher** than mud density while pumping.
- **Flow In vs Flow Out (Pump Out)** = the primary early kick/loss signal:
  - Flow Out **>** Flow In → possible **influx / kick**.
  - Flow Out **<** Flow In → possible **losses** to the formation.

**Why it matters for MPD.** The whole goal of Managed Pressure Drilling is to keep
bottom-hole pressure inside the narrow window between pore pressure and fracture
pressure. ECD is the number you steer; flow-in/out tells you whether the well is
balanced.

**Chart type in your file:** line + clustered-column combo. Flow In / Pump Out are
columns; Mud Density and PWD ECD are lines (secondary axis).

**⚠ Issue in the current build**
- The X-axis uses the auto **Date Hierarchy → Day** level, losing your 30-minute
  resolution. Switch to continuous `TimeBucket30min`.

---

## 3) Backpressure with Standpipe Pressure

**What it shows.** How well the MPD pressure-management hardware is behaving.
- **Standpipe Pressure (SPP)** = pressure feeding the drillstring; reflects pump
  rate and any downhole restriction. A sudden SPP change can flag a washout,
  plugged bit, or pack-off.
- **Actual vs Target Backpressure** = the applied surface backpressure vs its
  setpoint. In MPD, a backpressure pump + choke hold pressure on the closed
  annulus — especially during connections when the rig pumps stop (ECD drops to
  zero, so backpressure compensates). Actual tracking Target closely = healthy
  control; a persistent gap = a control problem.

**Why it matters for MPD.** This is the "is the choke/backpressure system doing its
job?" chart.

**Chart type in your file:** line + stacked-column combo. SPP is the column;
Actual + Target Backpressure are lines (secondary axis).

**⚠ Issue in the current build**
- Values are aggregated with **SUM** (`Sum(StandpipePressure_kPa)`,
  `Sum(ActualBackPressure_kPa)`, `Sum(TargetBackPressure_kPa)`). **Summing pressure
  over a bucket is physically wrong** — the bar height just grows with how many
  samples landed in the bucket. Use **AVERAGE** (measures provided below).

---

## 4) Net Gain/Loss with Flowline Temperature & Shaker Gas

**What it shows.** The primary **kick-detection / well-control** dashboard.
- **Net Gain/Loss (m³)** = pit / trip-tank volume balance. **Positive = gain**
  (possible influx), **negative = loss** (mud going into the formation). This is
  the single most important kick indicator.
- **Flowline Temperature (°C)** = return mud temperature. A sudden shift can
  accompany an influx (formation fluid at a different temperature) or a change in
  circulation.
- **Shaker Gas (%)** = gas detected at the shakers. Rising gas is an **early**
  warning of formation gas coming up the annulus.

**Why it matters for MPD.** Together these three give the earliest, clearest signal
that the well is taking a kick — the reason real-time monitoring exists.

**Chart type in your file:** line + clustered-column combo. Net Gain/Loss is the
column; Flowline Temp and Shaker Gas are lines (secondary axis).

**⚠ Issues in the current build**
- Flowline Temp and Shaker Gas use **SUM** (`Sum(FlowlineTemperature_C)`,
  `Sum(ShakerGas_pct)`) — again meaningless. Use **AVERAGE**.
- For gas and gain, a 30-minute **average can hide a short kick**. Use **MAX**
  within the bucket (or a 1-minute bucket) so transient peaks are not smoothed
  away.

---

## 5) Connection + Cleaning code (Power Query / M)

Paste this into **Home → Transform data → New Source → Blank Query → Advanced
Editor**, name the query `MPD`. It handles the messy realities of rig-sensor data:
absent-value sentinels (`-999.25` etc.), out-of-range spikes, null timestamps,
duplicate timestamps, sorting, and it rebuilds the `TimeBucket30min` (plus a finer
`TimeBucket1min`) column.

```m
// =====================================================================
//  MPD — Real-Time Monitoring source query
//  Connects to the rig data, cleans it, and builds time buckets.
// =====================================================================
let
    // ---------- 0. Settings you edit ----------
    // "Absent value" sentinels emitted by rig acquisition / WITSML / LAS.
    SentinelValues = { -999.25, -999, -9999, -9999.25, 9999, 999.25 },

    // Physical valid ranges. Anything outside becomes null (a gap), not a spike.
    Bounds = [
        HoleDepth_m            = {0,    12000},
        BitDepth_m             = {0,    12000},
        ROP_m_per_hr           = {0,    500},
        MudDensity_kg_m3       = {700,  2600},
        PWDECD_kg_m3           = {700,  2600},
        FlowIn_m3min           = {0,    12},
        PumpOut_m3min          = {0,    12},
        StandpipePressure_kPa  = {0,    60000},
        ActualBackPressure_kPa = {0,    30000},
        TargetBackPressure_kPa = {0,    30000},
        FlowlineTemperature_C  = {-10,  150},
        ShakerGas_pct          = {0,    100},
        NetGainLoss_m3         = {-100, 100}
    ],
    NumericCols = Record.FieldNames(Bounds),

    // ---------- 1. SOURCE ----------
    // OPTION A — SQL Server / Azure SQL staging table (near-real-time).
    // Only pulls the last 2 days so the model stays light; adjust as needed.
    Source = Sql.Database(
        "YOUR_SERVER", "YOUR_DATABASE",
        [Query =
            "SELECT TimeStamp, HoleDepth_m, BitDepth_m, ROP_m_per_hr,
                    MudDensity_kg_m3, PWDECD_kg_m3, FlowIn_m3min, PumpOut_m3min,
                    StandpipePressure_kPa, ActualBackPressure_kPa, TargetBackPressure_kPa,
                    FlowlineTemperature_C, ShakerGas_pct, NetGainLoss_m3
             FROM dbo.MPD_Realtime
             WHERE TimeStamp >= DATEADD(day, -2, SYSUTCDATETIME())"]
    ),

    // OPTION B — Folder of CSV exports dropped by the acquisition PC.
    // Comment out OPTION A above and use this instead:
    // Source =
    //     let
    //         Files   = Folder.Files("C:\RigData\MPD_Exports"),
    //         CsvOnly = Table.SelectRows(Files, each Text.EndsWith([Name], ".csv")),
    //         Loaded  = Table.AddColumn(CsvOnly, "Data",
    //                     each Csv.Document([Content],
    //                          [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv])),
    //         Promoted = Table.AddColumn(Loaded, "T",
    //                     each Table.PromoteHeaders([Data], [PromoteAllScalars=true])),
    //         Combined = Table.Combine(Promoted[T])
    //     in Combined,

    // ---------- 2. Type the columns ----------
    Typed = Table.TransformColumnTypes(
        Source,
        { {"TimeStamp", type datetime} }
        & List.Transform(NumericCols, each {_, type number})
    ),

    // ---------- 3. Replace sentinel/absent values with null ----------
    NoSentinels = List.Accumulate(
        SentinelValues, Typed,
        (state, sv) =>
            Table.ReplaceValue(state, sv, null, Replacer.ReplaceValue, NumericCols)
    ),

    // ---------- 4. Null-out physically impossible values ----------
    Clamped = Table.TransformColumns(
        NoSentinels,
        List.Transform(
            NumericCols,
            (c) =>
                let b = Record.Field(Bounds, c)
                in { c,
                     (v) => if v = null then null
                            else if v < b{0} or v > b{1} then null
                            else v,
                     type nullable number }
        )
    ),

    // ---------- 5. Drop rows with no timestamp, sort, de-duplicate ----------
    ValidTime = Table.SelectRows(Clamped, each [TimeStamp] <> null),
    Sorted    = Table.Sort(ValidTime, {{"TimeStamp", Order.Ascending}}),
    Deduped   = Table.Distinct(Sorted, {"TimeStamp"}),

    // ---------- 6. Useful row-level derived columns ----------
    // Flow balance: positive => flow-out exceeds flow-in => possible influx.
    WithFlowDelta = Table.AddColumn(Deduped, "FlowDelta_m3min",
        each try [PumpOut_m3min] - [FlowIn_m3min] otherwise null, type nullable number),
    // On-bottom flag: bit within 0.5 m of hole bottom.
    WithOnBottom = Table.AddColumn(WithFlowDelta, "OnBottom",
        each try (Number.Abs([HoleDepth_m] - [BitDepth_m]) <= 0.5) otherwise null, type nullable logical),

    // ---------- 7. Rebuild time buckets ----------
    AddBucket30 = Table.AddColumn(WithOnBottom, "TimeBucket30min",
        each
            let t = [TimeStamp], m = Time.Minute(DateTime.Time(t))
            in #datetime(Date.Year(t), Date.Month(t), Date.Day(t),
                         Time.Hour(DateTime.Time(t)), m - Number.Mod(m, 30), 0),
        type datetime),
    AddBucket1 = Table.AddColumn(AddBucket30, "TimeBucket1min",
        each
            let t = [TimeStamp]
            in #datetime(Date.Year(t), Date.Month(t), Date.Day(t),
                         Time.Hour(DateTime.Time(t)), Time.Minute(DateTime.Time(t)), 0),
        type datetime),

    // ---------- 8. Buffer for performance ----------
    Result = Table.Buffer(AddBucket1)
in
    Result
```

**Notes**
- If your source column is called something else (e.g. `Timestamp`, `DateTime`,
  or the net-gain channel is `Net Gain/Loss`), rename it to match the names above
  in one `Table.RenameColumns` step right after `Source`, or edit the SQL `SELECT`.
- For a true real-time feed, WITSML/historian data usually lands in a SQL/OData
  staging table first — point OPTION A at that table and set an automatic refresh
  (Power BI Service scheduled refresh, or a Streaming/Push dataset for sub-minute).

---

## 6) Corrected DAX measures

Create these (Modeling → New measure). They replace the wrong SUM aggregations and
add a few operational helpers. If your net-gain column is named `Net Gain/Loss`,
adjust the column reference.

```dax
// ---- Graph 1: Hole Depth, Bit Depth & ROP ----
Avg Hole Depth (m) = AVERAGE ( MPD[HoleDepth_m] )
Avg Bit Depth (m)  = AVERAGE ( MPD[BitDepth_m] )
Avg ROP (m/hr)     = AVERAGE ( MPD[ROP_m_per_hr] )
On-Bottom % =
DIVIDE (
    CALCULATE ( COUNTROWS ( MPD ), MPD[OnBottom] = TRUE() ),
    COUNTROWS ( MPD )
)

// ---- Graph 2: PWD ECD & Mud Density with Flow In / Out ----
Avg Mud Density (kg/m3) = AVERAGE ( MPD[MudDensity_kg_m3] )
Avg PWD ECD (kg/m3)     = AVERAGE ( MPD[PWDECD_kg_m3] )
Avg Flow In (m3/min)    = AVERAGE ( MPD[FlowIn_m3min] )
Avg Flow Out (m3/min)   = AVERAGE ( MPD[PumpOut_m3min] )
ECD - MW Overbalance (kg/m3) = [Avg PWD ECD (kg/m3)] - [Avg Mud Density (kg/m3)]
Flow Delta out-in (m3/min)   = [Avg Flow Out (m3/min)] - [Avg Flow In (m3/min)]

// ---- Graph 3: Backpressure with Standpipe Pressure  (SUM -> AVERAGE) ----
Avg Standpipe Pressure (kPa)   = AVERAGE ( MPD[StandpipePressure_kPa] )
Avg Actual Backpressure (kPa)  = AVERAGE ( MPD[ActualBackPressure_kPa] )
Avg Target Backpressure (kPa)  = AVERAGE ( MPD[TargetBackPressure_kPa] )
Backpressure Deviation (kPa)   = [Avg Actual Backpressure (kPa)] - [Avg Target Backpressure (kPa)]

// ---- Graph 4: Net Gain/Loss with Flowline Temp & Shaker Gas ----
Avg Flowline Temp (C) = AVERAGE ( MPD[FlowlineTemperature_C] )
Avg Shaker Gas (%)    = AVERAGE ( MPD[ShakerGas_pct] )
Max Shaker Gas (%)    = MAX ( MPD[ShakerGas_pct] )       // keeps kick peaks in a bucket
Net Gain/Loss (m3)    = AVERAGE ( MPD[NetGainLoss_m3] )  // average balance in the bucket
Max Net Gain (m3)     = MAX ( MPD[NetGainLoss_m3] )
// If NetGainLoss is a *cumulative* trip-tank total, use the last value instead:
Net Gain/Loss (last) =
VAR MaxT = MAX ( MPD[TimeStamp] )
RETURN CALCULATE ( SUM ( MPD[NetGainLoss_m3] ), MPD[TimeStamp] = MaxT )

// ---- Simple kick warning flag (optional KPI card) ----
Kick Warning =
IF (
    [Max Shaker Gas (%)] > 5
        || [Max Net Gain (m3)] > 1
        || [Flow Delta out-in (m3/min)] > 0.1,
    "CHECK", "OK"
)
```

---

## 7) Visual fixes checklist

For each of the four charts, in the visual's **Fields** and **Format** panes:

1. **Graph 1 (Hole/Bit/ROP)** — add `Avg Hole Depth (m)` as a second line so
   hole vs bit shows on/off-bottom. Put ROP as the column. On the X-axis, expand
   `TimeBucket30min`, choose the datetime field directly (not the Year/Day
   hierarchy), and set **X-axis → Type = Continuous**.
2. **Graph 2 (ECD/MW & Flow)** — set the X-axis to the continuous `TimeBucket30min`
   (currently drilled to *Day*). Keep Flow In/Out as columns, ECD/MW as lines.
3. **Graph 3 (Backpressure & SPP)** — replace the SUM fields with the AVERAGE
   measures above (`Avg Standpipe Pressure`, `Avg Actual/Target Backpressure`).
   Optionally add `Backpressure Deviation (kPa)`. Continuous X-axis.
4. **Graph 4 (Gain/Loss, Temp, Gas)** — replace the SUM fields with
   `Avg Flowline Temp (C)` and `Max Shaker Gas (%)` (MAX preserves peaks); use
   `Max Net Gain (m3)` for the column. Continuous X-axis.
5. **All charts** — to catch short kicks, consider switching the axis from
   `TimeBucket30min` to `TimeBucket1min` on the gas / gain / flow charts; 30-minute
   averaging is fine for depth/ROP but can smooth away a real-time influx.
6. Turn off Power BI's automatic date hierarchy for this table
   (**File → Options → Data Load → Time intelligence → uncheck Auto date/time**)
   so the continuous datetime axis behaves.
