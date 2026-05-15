"use client";

import React, { useMemo, useState } from "react";
import * as XLSX from "xlsx";
import { Card, CardContent } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { SimpleSelect } from "@/components/ui/select";
import { Upload, Download, BarChart3, FileSpreadsheet } from "lucide-react";
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, LabelList } from "recharts";

const DAYS = ["月", "火", "水", "木", "金", "土", "日"];
const HOURS = ["9時台", "10時台", "11時台", "12時台", "13時台", "14時台", "15時台", "16時台", "17時台", "18時台", "19時台", "20時台"];
const CALL_FILTERS = ["全回合計", "1回目", "2回目", "3回目", "4回目", "5回目"];
const METRICS = ["発信数", "接続数", "接続率"];

type Detail = {
  rowNo: number;
  connected: boolean;
  callLabel: string;
  isFinalCall: boolean;
  day: string | null;
  hour: string | null;
};

type Matrix = Record<string, Record<string, number>>;

type SummaryRow = Record<string, number | string> & { 曜日: string; 合計: number };

function excelSerialToDate(serial: number) {
  const utcDays = Math.floor(serial - 25569);
  const utcValue = utcDays * 86400;
  const dateInfo = new Date(utcValue * 1000);
  const fractionalDay = serial - Math.floor(serial) + 0.0000001;
  const totalSeconds = Math.floor(86400 * fractionalDay);
  const seconds = totalSeconds % 60;
  const hours = Math.floor(totalSeconds / 3600);
  const minutes = Math.floor(totalSeconds / 60) % 60;
  return new Date(dateInfo.getFullYear(), dateInfo.getMonth(), dateInfo.getDate(), hours, minutes, seconds);
}

function parseDate(value: unknown) {
  if (!value) return null;
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  if (typeof value === "number") return excelSerialToDate(value);
  if (typeof value === "string") {
    const normalized = value.trim().replace(/年|月/g, "/").replace(/日/g, "").replace(/-/g, "/");
    const parsed = new Date(normalized);
    if (!Number.isNaN(parsed.getTime())) return parsed;
  }
  return null;
}

function isConnected(result: unknown) {
  const text = String(result ?? "");
  return (
    text.includes("通話") ||
    text.includes("301.解約") ||
    text.includes("6.逝去") ||
    text.includes("305.フォロー済") ||
    text.includes("306.変更") ||
    text.includes("8.認知症")
  );
}

function getDayLabel(date: Date) {
  return ["日", "月", "火", "水", "木", "金", "土"][date.getDay()];
}

function getHourLabel(date: Date) {
  const hour = date.getHours();
  if (hour >= 20) return "20時台";
  if (hour < 9) return null;
  return `${hour}時台`;
}

function emptyMatrix(): Matrix {
  const obj: Matrix = {};
  DAYS.forEach((d) => {
    obj[d] = {};
    HOURS.forEach((h) => (obj[d][h] = 0));
  });
  return obj;
}

function buildSummary(details: Detail[], callFilter: string) {
  const calls = emptyMatrix();
  const connects = emptyMatrix();

  details.forEach((row) => {
    if (callFilter === "全回合計" || row.callLabel === callFilter) {
      if (row.day && row.hour) calls[row.day][row.hour] += 1;
    }
    if (row.connected && row.isFinalCall && (callFilter === "全回合計" || row.callLabel === callFilter)) {
      if (row.day && row.hour) connects[row.day][row.hour] += 1;
    }
  });

  return { calls, connects };
}

function matrixToRows(matrix: Matrix): SummaryRow[] {
  return DAYS.map((day) => {
    const row: SummaryRow = { 曜日: day, 合計: 0 };
    HOURS.forEach((h) => {
      row[h] = matrix[day][h] || 0;
      row.合計 += Number(row[h]);
    });
    return row;
  });
}

function rateRows(calls: Matrix, connects: Matrix): SummaryRow[] {
  return DAYS.map((day) => {
    const row: SummaryRow = { 曜日: day, 合計: 0 };
    let callTotal = 0;
    let connectTotal = 0;
    HOURS.forEach((h) => {
      const c = calls[day][h] || 0;
      const k = connects[day][h] || 0;
      row[h] = c ? k / c : 0;
      callTotal += c;
      connectTotal += k;
    });
    row.合計 = callTotal ? connectTotal / callTotal : 0;
    return row;
  });
}

function formatPct(value: number) {
  return `${(Number(value || 0) * 100).toFixed(1)}%`;
}

export default function OutboundAnalysisApp() {
  const [fileName, setFileName] = useState("");
  const [details, setDetails] = useState<Detail[]>([]);
  const [callFilter, setCallFilter] = useState("全回合計");
  const [metric, setMetric] = useState("接続率");
  const [error, setError] = useState("");

  const handleFile = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;
    setError("");
    setFileName(file.name);

    try {
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data, { type: "array", cellDates: true });
      const sheetName = workbook.SheetNames.includes("5月全体") ? "5月全体" : workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" }) as unknown[][];

      const generated: Detail[] = [];
      rows.slice(1, 30001).forEach((row, index) => {
        const result = row[1];
        const connected = isConnected(result);
        const callDates = [row[2], row[3], row[4], row[5], row[6]]
          .map(parseDate)
          .filter((v): v is Date => Boolean(v))
          .sort((a, b) => a.getTime() - b.getTime())
          .slice(0, 5);

        callDates.forEach((date, callIndex) => {
          generated.push({
            rowNo: index + 2,
            connected,
            callLabel: `${callIndex + 1}回目`,
            isFinalCall: callIndex === callDates.length - 1,
            day: getDayLabel(date),
            hour: getHourLabel(date),
          });
        });
      });

      setDetails(generated);
    } catch (e) {
      setError("ファイルを読み込めませんでした。Excel形式（.xlsx）またはCSV形式をご確認ください。");
      setDetails([]);
    }
  };

  const summary = useMemo(() => buildSummary(details, callFilter), [details, callFilter]);
  const callRows = useMemo(() => matrixToRows(summary.calls), [summary]);
  const connectRows = useMemo(() => matrixToRows(summary.connects), [summary]);
  const pctRows = useMemo(() => rateRows(summary.calls, summary.connects), [summary]);

  const totals = useMemo(() => {
    const totalCalls = callRows.reduce((sum, r) => sum + r.合計, 0);
    const totalConnects = connectRows.reduce((sum, r) => sum + r.合計, 0);
    return { totalCalls, totalConnects, rate: totalCalls ? totalConnects / totalCalls : 0 };
  }, [callRows, connectRows]);

  const chartData = useMemo(() => {
    const source = metric === "発信数" ? callRows : metric === "接続数" ? connectRows : pctRows;
    return source.map((r) => ({ 曜日: r.曜日, value: Number(r.合計) }));
  }, [metric, callRows, connectRows, pctRows]);

  const exportExcel = () => {
    const wb = XLSX.utils.book_new();
    const addSheet = (name: string, rows: SummaryRow[], isPct = false) => {
      const aoa: (string | number)[][] = [["曜日", ...HOURS, "合計"]];
      rows.forEach((r) => aoa.push([r.曜日, ...HOURS.map((h) => (isPct ? formatPct(Number(r[h])) : Number(r[h]))), isPct ? formatPct(r.合計) : r.合計]));
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), name);
    };
    addSheet("発信数", callRows);
    addSheet("接続数", connectRows);
    addSheet("接続率", pctRows, true);
    XLSX.writeFile(wb, `アウト分析_${callFilter}.xlsx`);
  };

  const renderTable = (title: string, rows: SummaryRow[], isPct = false) => (
    <Card>
      <CardContent>
        <h3 className="mb-3 text-lg font-bold">{title}</h3>
        <div className="overflow-auto">
          <table className="w-full min-w-[980px] border-collapse text-sm">
            <thead>
              <tr className="bg-slate-800 text-white">
                <th className="border px-2 py-2 text-left">曜日</th>
                {HOURS.map((h) => <th key={h} className="border px-2 py-2 text-right">{h}</th>)}
                <th className="border px-2 py-2 text-right">合計</th>
              </tr>
            </thead>
            <tbody>
              {rows.map((r) => (
                <tr key={r.曜日} className="odd:bg-white even:bg-slate-50">
                  <td className="border px-2 py-2 font-medium">{r.曜日}</td>
                  {HOURS.map((h) => <td key={h} className="border px-2 py-2 text-right">{isPct ? formatPct(Number(r[h])) : Number(r[h]).toLocaleString()}</td>)}
                  <td className="border bg-orange-50 px-2 py-2 text-right font-bold">{isPct ? formatPct(r.合計) : r.合計.toLocaleString()}</td>
                </tr>
              ))}
              <tr className="bg-orange-100 font-bold">
                <td className="border px-2 py-2">合計</td>
                {HOURS.map((h) => {
                  const c = callRows.reduce((s, r) => s + Number(r[h]), 0);
                  const k = connectRows.reduce((s, r) => s + Number(r[h]), 0);
                  const v = isPct ? (c ? k / c : 0) : rows.reduce((s, r) => s + Number(r[h]), 0);
                  return <td key={h} className="border px-2 py-2 text-right">{isPct ? formatPct(v) : v.toLocaleString()}</td>;
                })}
                <td className="border px-2 py-2 text-right">{isPct ? formatPct(totals.rate) : rows.reduce((s, r) => s + r.合計, 0).toLocaleString()}</td>
              </tr>
            </tbody>
          </table>
        </div>
      </CardContent>
    </Card>
  );

  return (
    <main className="min-h-screen bg-slate-100 p-6 text-slate-900">
      <div className="mx-auto max-w-7xl space-y-6">
        <section className="rounded-2xl bg-slate-900 p-6 text-white shadow-sm">
          <div className="flex flex-col gap-4 md:flex-row md:items-center md:justify-between">
            <div>
              <h1 className="text-2xl font-bold">アウトバウンド 曜日・時間帯別 接続分析アプリ</h1>
              <p className="mt-2 text-sm text-slate-300">B列=結果、C〜G列=発信日時。最大3万件まで想定。21時台は20時台へ加算します。</p>
            </div>
            <label className="inline-flex cursor-pointer items-center gap-2 rounded-xl bg-white px-4 py-3 font-bold text-slate-900 shadow-sm hover:bg-slate-100">
              <Upload size={18} />
              Excelをアップロード
              <input type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={handleFile} />
            </label>
          </div>
        </section>

        {error && <div className="rounded-xl bg-red-100 p-4 text-red-700">{error}</div>}

        <section className="grid gap-4 md:grid-cols-4">
          <Card><CardContent><div className="text-sm text-slate-500">読込ファイル</div><div className="mt-2 flex items-center gap-2 font-bold"><FileSpreadsheet size={18}/>{fileName || "未選択"}</div></CardContent></Card>
          <Card><CardContent><div className="text-sm text-slate-500">総発信数</div><div className="mt-2 text-2xl font-bold">{totals.totalCalls.toLocaleString()}</div></CardContent></Card>
          <Card><CardContent><div className="text-sm text-slate-500">総接続数</div><div className="mt-2 text-2xl font-bold">{totals.totalConnects.toLocaleString()}</div></CardContent></Card>
          <Card><CardContent><div className="text-sm text-slate-500">接続率</div><div className="mt-2 text-2xl font-bold">{formatPct(totals.rate)}</div></CardContent></Card>
        </section>

        <Card>
          <CardContent className="flex flex-col gap-4 md:flex-row md:items-end md:justify-between">
            <div className="flex flex-col gap-3 md:flex-row">
              <div>
                <div className="mb-1 text-sm font-bold">回数</div>
                <SimpleSelect value={callFilter} onChange={setCallFilter} options={CALL_FILTERS} className="w-40" />
              </div>
              <div>
                <div className="mb-1 text-sm font-bold">グラフ指標</div>
                <SimpleSelect value={metric} onChange={setMetric} options={METRICS} className="w-40" />
              </div>
            </div>
            <Button onClick={exportExcel} disabled={details.length === 0}><Download className="mr-2" size={18}/>Excel出力</Button>
          </CardContent>
        </Card>

        <Card>
          <CardContent>
            <div className="mb-3 flex items-center gap-2 text-lg font-bold"><BarChart3 size={20}/>{metric}：曜日別合計</div>
            <div className="h-72">
              <ResponsiveContainer width="100%" height="100%">
                <BarChart data={chartData} margin={{ top: 25, right: 25, left: 10, bottom: 5 }}>
                  <CartesianGrid strokeDasharray="3 3" />
                  <XAxis dataKey="曜日" />
                  <YAxis tickFormatter={(v) => metric === "接続率" ? `${Math.round(Number(v) * 100)}%` : Number(v).toLocaleString()} />
                  <Tooltip formatter={(v) => metric === "接続率" ? formatPct(Number(v)) : Number(v).toLocaleString()} />
                  <Bar dataKey="value" radius={[8, 8, 0, 0]}>
                    <LabelList dataKey="value" position="top" formatter={(v: number) => metric === "接続率" ? formatPct(v) : Number(v).toLocaleString()} />
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>
          </CardContent>
        </Card>

        {renderTable("曜日 × 時間帯：発信数", callRows)}
        {renderTable("曜日 × 時間帯：接続数", connectRows)}
        {renderTable("曜日 × 時間帯：接続率", pctRows, true)}
      </div>
    </main>
  );
}
