import React, { useState, useEffect } from 'react'
import {
  FiSearch, FiRefreshCw, FiChevronUp, FiChevronDown,
  FiImage, FiX, FiCalendar, FiPackage,
  FiHash, FiClock, FiZoomIn, FiDownload as FiDownloadIcon,
  FiChevronLeft, FiChevronRight, FiMapPin, FiTruck as FiTruckIcon,
  FiCheckCircle, FiAlertCircle,
} from 'react-icons/fi'
import * as XLSX from 'xlsx'
import { writeXlsx } from '../../../utils/excelDates'
import { FaTrain } from 'react-icons/fa'
import Navbar from '../../../components/layout/Navbar'
import Footer from '../../../components/layout/Footer'
import { useLazyGetWbinSurveyQuery } from '../../../store/api/ymsApi'
import { buildAssetUrl } from '../../../config/api'
import { StatCard, StatGrid, FilterCard } from '../../../components/ui/StatCard'
import { FilterBar, FilterField, FilterSearchBtn, FilterClearBtn, FilterExportBtn } from '../../../components/ui/FilterBar'

// ── Image helpers ─────────────────────────────────────────────────────────────
// SP returns paths like "stitching_outputs/WB_IN/20261118/CONT/left.jpg"
// We strip the "stitching_outputs/" prefix and proxy through /v1/container/img
function getWbinImageUrl(rawPath) {
  if (!rawPath) return null
  const rel = String(rawPath).replace(/^\/?stitching_outputs\//, '')
  return rel ? buildAssetUrl(`/v1/container/img?p=${rel}`) : null
}

const IMAGE_SIDES = [
  { key: 'Left1',  label: 'L',  side: 'left',         damage: false },
  { key: 'Right1', label: 'R',  side: 'right',        damage: false },
  { key: 'Back1',  label: 'B',  side: 'back',         damage: false },
  { key: 'Top1',   label: 'T',  side: 'top',          damage: false },
  { key: 'Left2',  label: 'LD', side: 'left damage',  damage: true  },
  { key: 'Right2', label: 'RD', side: 'right damage', damage: true  },
  { key: 'Top2',   label: 'TD', side: 'top damage',   damage: true  },
]

function formatDateParts(raw) {
  if (!raw) return null
  try {
    const d = new Date(String(raw).replace(' ', 'T'))
    if (isNaN(d.getTime())) return null
    const p = (n) => String(n).padStart(2, '0')
    return {
      date: `${p(d.getDate())}/${p(d.getMonth() + 1)}/${d.getFullYear()}`,
      time: `${p(d.getHours())}:${p(d.getMinutes())}`,
    }
  } catch { return null }
}
function formatDate(raw) {
  const parts = formatDateParts(raw)
  return parts ? `${parts.date} ${parts.time}` : ''
}

// ── Image probe cache ─────────────────────────────────────────────────────────
const imgProbeCache = new Map()
function useImageStatus(url) {
  const [status, setStatus] = useState(() => (!url ? 'missing' : imgProbeCache.get(url) || 'loading'))
  useEffect(() => {
    if (!url) { setStatus('missing'); return }
    const cached = imgProbeCache.get(url)
    if (cached) { setStatus(cached); return }
    setStatus('loading')
    let cancelled = false
    const img = new Image()
    img.onload  = () => { imgProbeCache.set(url, 'ok');      if (!cancelled) setStatus('ok') }
    img.onerror = () => { imgProbeCache.set(url, 'missing'); if (!cancelled) setStatus('missing') }
    img.src = url
    return () => { cancelled = true }
  }, [url])
  return status
}

// ── SafeImg ───────────────────────────────────────────────────────────────────
function SafeImg({ url, alt, className, style, fallback }) {
  const [err, setErr] = useState(false)
  useEffect(() => { setErr(false) }, [url])
  if (err) return fallback
  return <img src={url} alt={alt} className={className} style={style} onError={() => setErr(true)} />
}

// ── Lightbox ──────────────────────────────────────────────────────────────────
function ImageLightbox({ url, label, onClose }) {
  if (!url) return null
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/85 backdrop-blur-sm p-4" onClick={onClose}>
      <div className="relative max-w-5xl w-full bg-white rounded-2xl shadow-2xl overflow-hidden" onClick={e => e.stopPropagation()}>
        <div className="flex items-center justify-between bg-gradient-to-r from-[#0e4a78] via-[#0b3e66] to-[#072c4a] text-white px-5 py-3">
          <span className="font-semibold capitalize text-base flex items-center gap-2"><FiImage /> {label} Image</span>
          <button onClick={onClose} className="p-1.5 rounded-lg hover:bg-white/20 transition-colors"><FiX className="text-lg" /></button>
        </div>
        <div className="flex items-center justify-center bg-slate-50 p-4 overflow-auto" style={{ maxHeight: '85vh' }}>
          <SafeImg url={url} alt={label} className="max-w-full object-contain rounded-lg" style={{ maxHeight: '80vh' }}
            fallback={<div className="py-16 flex flex-col items-center gap-2 text-slate-400"><FiImage size={40} /><span className="text-sm font-semibold">Image not found</span></div>}
          />
        </div>
      </div>
    </div>
  )
}

// ── Sort icon ─────────────────────────────────────────────────────────────────
function SortIcon({ col, sortCol, sortDir }) {
  if (sortCol !== col) return <span className="opacity-30 ml-1"><FiChevronUp size={10} /></span>
  return sortDir === 'asc' ? <FiChevronUp className="ml-1 text-blue-200" size={11} /> : <FiChevronDown className="ml-1 text-blue-200" size={11} />
}

// ── Date cell ─────────────────────────────────────────────────────────────────
function DateCell({ raw }) {
  const parts = formatDateParts(raw)
  if (!parts) return <span className="text-slate-300 text-xs">—</span>
  return (
    <div>
      <p className="text-[11px] font-semibold text-slate-700">{parts.date}</p>
      <p className="text-[10px] text-slate-400">{parts.time}</p>
    </div>
  )
}

// ── Thumb button ──────────────────────────────────────────────────────────────
function ThumbBtn({ url, label, side, damage, onOpen }) {
  const status = useImageStatus(url)
  if (status !== 'ok') {
    return (
      <button disabled title={status === 'loading' ? `Checking ${side}…` : `No ${side} image`}
        className={`w-7 h-7 rounded-lg border border-slate-200 bg-slate-100 text-slate-300 opacity-50 cursor-not-allowed flex items-center justify-center text-[9px] font-bold ${status === 'loading' ? 'animate-pulse' : ''}`}>
        {label}
      </button>
    )
  }
  const cls = damage
    ? 'border-red-500/30 bg-red-500/10 text-red-600 hover:bg-red-600 hover:text-white'
    : 'border-[#0e4a78]/30 bg-[#0e4a78]/10 text-[#0e4a78] hover:bg-[#0e4a78] hover:text-white'
  return (
    <button onClick={() => onOpen(url, side)} title={`View ${side}`}
      className={`w-7 h-7 rounded-lg border flex items-center justify-center text-[9px] font-bold transition-all ${cls}`}>
      {label}
    </button>
  )
}

function ImagesCell({ row, onOpen }) {
  return (
    <div className="flex items-center gap-1 justify-center">
      {IMAGE_SIDES.map(({ key, label, side, damage }) => (
        <ThumbBtn key={key} url={getWbinImageUrl(row[key])} label={label} side={side} damage={damage} onOpen={onOpen} />
      ))}
    </div>
  )
}

// ── Image slot (detail modal) ─────────────────────────────────────────────────
function ImageSlot({ url, label, accent, onZoom }) {
  const status = useImageStatus(url)
  const ok = status === 'ok'
  return (
    <div className="bg-white rounded-2xl border border-slate-200 shadow-md overflow-hidden flex flex-col min-h-0">
      <div className="shrink-0 flex items-center justify-between gap-2 px-4 py-2.5 border-b border-slate-100 bg-slate-50/60">
        <span className="flex items-center gap-2">
          <span className="w-6 h-6 rounded-lg flex items-center justify-center" style={{ background: `${accent}12`, color: accent }}><FiImage size={12} /></span>
          <span className="text-[11px] font-black uppercase tracking-wider text-slate-600">{label}</span>
        </span>
        {ok && (
          <a href={url} download onClick={e => e.stopPropagation()}
            className="w-6 h-6 rounded-lg flex items-center justify-center text-slate-400 hover:text-[#0e4a78] hover:bg-[#0e4a78]/10 transition-colors" title="Download">
            <FiDownloadIcon size={12} />
          </a>
        )}
      </div>
      {ok ? (
        <button onClick={() => onZoom(url, label)}
          className="group relative flex-1 min-h-0 flex items-center justify-center overflow-hidden w-full"
          style={{ background: "repeating-conic-gradient(#f8fafc 0% 25%, #f1f5f9 0% 50%) 50% / 24px 24px" }}>
          <img src={url} alt={label} className="max-w-full max-h-full object-contain p-2" />
          <span className="absolute inset-0 bg-black/0 group-hover:bg-black/35 transition-colors flex items-center justify-center">
            <span className="w-14 h-14 rounded-full bg-white/90 opacity-0 group-hover:opacity-100 scale-90 group-hover:scale-100 transition-all flex items-center justify-center shadow-xl">
              <FiZoomIn className="text-[#0e4a78]" size={24} />
            </span>
          </span>
        </button>
      ) : (
        <div className="flex-1 min-h-0 flex flex-col items-center justify-center gap-2.5 text-slate-300 bg-slate-50/50 w-full">
          <span className={`w-16 h-16 rounded-2xl bg-slate-100 flex items-center justify-center ${status === 'loading' ? 'animate-pulse' : ''}`}><FiImage size={28} /></span>
          <span className="text-xs font-semibold text-slate-400">{status === 'loading' ? 'Checking…' : 'No image available'}</span>
        </div>
      )}
    </div>
  )
}

// ── Detail field ──────────────────────────────────────────────────────────────
function DetailField({ icon: Icon, label, children, accent = "#0e4a78" }) {
  return (
    <div className="flex items-start gap-2.5 py-1.5 border-b border-slate-100 last:border-b-0">
      <span className="flex-shrink-0 w-7 h-7 rounded-lg flex items-center justify-center mt-0.5" style={{ background: `${accent}12`, color: accent }}>
        {Icon && <Icon size={12} />}
      </span>
      <div className="min-w-0 flex-1">
        <p className="text-[9px] font-bold uppercase tracking-[0.1em] text-slate-400 leading-tight">{label}</p>
        <div className="text-[13px] font-semibold text-slate-800 break-words leading-tight">{children}</div>
      </div>
    </div>
  )
}

function DetailCard({ title, icon: Icon, accent, children }) {
  return (
    <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden flex flex-col min-h-0">
      <div className="shrink-0 flex items-center gap-2 px-4 py-2.5 border-b border-slate-100 bg-slate-50/60">
        <span className="w-6 h-6 rounded-lg flex items-center justify-center" style={{ background: `${accent}12`, color: accent }}><Icon size={12} /></span>
        <span className="text-[11px] font-black uppercase tracking-wider text-slate-600">{title}</span>
      </div>
      <div className="px-4 py-1 overflow-y-auto grid grid-cols-1 sm:grid-cols-2 sm:gap-x-4">{children}</div>
    </div>
  )
}

// ── Detail modal ──────────────────────────────────────────────────────────────
function DetailModal({ row, index, total, onClose, onPrev, onNext, onZoom }) {
  if (!row) return null
  const surveyed = row.IsSuveryCompleted === true || row.IsSuveryCompleted === 1
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/70 backdrop-blur-sm p-3 sm:p-6" onClick={onClose}>
      <div className="relative w-full max-w-[1600px] bg-white rounded-3xl shadow-2xl overflow-hidden flex flex-col ring-1 ring-black/5" style={{ height: '95vh' }} onClick={e => e.stopPropagation()}>
        {/* Header */}
        <div className="relative flex items-center justify-between text-white px-7 py-5 shrink-0 overflow-hidden"
          style={{ background: "linear-gradient(120deg,#0e4a78 0%,#0b3e66 45%,#072c4a 100%)" }}>
          <div className="absolute inset-0 opacity-[0.07] pointer-events-none"
            style={{ backgroundImage: "radial-gradient(circle at 1px 1px, white 1px, transparent 0)", backgroundSize: "18px 18px" }} />
          <div className="relative flex items-center gap-4">
            <div className="w-12 h-12 rounded-2xl bg-white/15 flex items-center justify-center shadow-inner"><FaTrain size={20} /></div>
            <div>
              <p className="font-black text-xl leading-tight tracking-tight">{row.ContNo || 'Record'}</p>
              <p className="text-xs text-white/60 font-medium mt-0.5">WB In Survey · {index + 1} of {total}</p>
            </div>
          </div>
          <div className="relative flex items-center gap-2">
            <button onClick={onPrev} disabled={index <= 0} className="w-10 h-10 rounded-xl flex items-center justify-center bg-white/10 hover:bg-white/20 disabled:opacity-25 disabled:cursor-not-allowed transition-all"><FiChevronLeft size={19} /></button>
            <button onClick={onNext} disabled={index >= total - 1} className="w-10 h-10 rounded-xl flex items-center justify-center bg-white/10 hover:bg-white/20 disabled:opacity-25 disabled:cursor-not-allowed transition-all"><FiChevronRight size={19} /></button>
            <div className="w-px h-7 bg-white/20 mx-2" />
            <button onClick={onClose} className="w-10 h-10 rounded-xl flex items-center justify-center bg-white/10 hover:bg-red-500/80 transition-all"><FiX size={19} /></button>
          </div>
        </div>

        {/* Body */}
        <div className="flex-1 min-h-0 flex flex-col bg-slate-100">
          {/* Details top */}
          <div className="shrink-0 border-b border-slate-200 bg-gradient-to-b from-slate-50 to-white">
            <div className="px-6 py-4">
              <div className="mb-3 flex flex-wrap gap-2.5">
                {row.ContNo && (
                  <span className="inline-flex items-center gap-2 px-4 py-2 rounded-xl bg-gradient-to-br from-[#1364a4] to-[#0a3b61] text-white font-black font-mono text-base shadow-lg shadow-[#0e4a78]/25">
                    <FiPackage size={15} /> {row.ContNo}
                  </span>
                )}
                <span className={`inline-flex items-center gap-1.5 px-3 py-2 rounded-xl text-xs font-bold border ${surveyed ? 'bg-emerald-50 text-emerald-700 border-emerald-200' : 'bg-amber-50 text-amber-700 border-amber-200'}`}>
                  {surveyed ? <FiCheckCircle size={13} /> : <FiAlertCircle size={13} />}
                  {surveyed ? 'Survey Done' : 'Survey Pending'}
                </span>
              </div>
              <div className="grid grid-cols-1 lg:grid-cols-2 gap-3">
                {/* Container details */}
                <DetailCard title="Container" icon={FiPackage} accent="#0e4a78">
                  <DetailField icon={FiPackage} label="Container No" accent="#0e4a78">{row.ContNo || <span className="text-slate-300 font-normal">—</span>}</DetailField>
                  <DetailField icon={FiHash} label="Size / Type" accent="#0e4a78">{`${row.ContainerSize || '—'} ${row.ContainerType || ''}`.trim() || '—'}</DetailField>
                  <DetailField icon={FiHash} label="Process" accent="#059669">{row.Process || <span className="text-slate-300 font-normal">—</span>}</DetailField>
                  <DetailField icon={FiHash} label="Shipping Line" accent="#7c3aed">{row.ShippingLine || <span className="text-slate-300 font-normal">—</span>}</DetailField>
                  <DetailField icon={FiHash} label="Booking No" accent="#64748b">{row.BookingNo || <span className="text-slate-300 font-normal">—</span>}</DetailField>
                  <DetailField icon={FiHash} label="Document No" accent="#64748b">{row.DocumentNo || <span className="text-slate-300 font-normal">—</span>}</DetailField>
                  <DetailField icon={FiHash} label="Status" accent="#d97706">{row.ContainerStatus || <span className="text-slate-300 font-normal">—</span>}</DetailField>
                  <DetailField icon={FiHash} label="Mode" accent="#0284c7">{row.Mode || <span className="text-slate-300 font-normal">—</span>}</DetailField>
                  <DetailField icon={FiMapPin} label="Location" accent="#d97706">{row.ContainerLocationName || <span className="text-slate-300 font-normal">—</span>}</DetailField>
                </DetailCard>
                {/* Gate / timing details */}
                <DetailCard title="Gate & Timing" icon={FiTruckIcon} accent="#0e4a78">
                  <DetailField icon={FiTruckIcon} label="Trailer No" accent="#0e4a78">{row.TrailerNo || <span className="text-slate-300 font-normal">—</span>}</DetailField>
                  <DetailField icon={FiHash} label="Gate Name" accent="#d97706">{row.GateName || <span className="text-slate-300 font-normal">—</span>}</DetailField>
                  <DetailField icon={FiClock} label="Survey Time" accent="#0284c7">{row.SurveyTime ? formatDate(row.SurveyTime) : <span className="text-slate-300 font-normal">—</span>}</DetailField>
                  <DetailField icon={FiClock} label="Detected Time" accent="#0284c7">{row.DetectedTime ? formatDate(row.DetectedTime) : <span className="text-slate-300 font-normal">—</span>}</DetailField>
                  <DetailField icon={FiClock} label="Gate In" accent="#059669">{row.GateInDate ? formatDate(row.GateInDate) : <span className="text-slate-300 font-normal">—</span>}</DetailField>
                  <DetailField icon={FiClock} label="Gate Out" accent="#d97706">{row.GateOutDate ? formatDate(row.GateOutDate) : <span className="text-slate-300 font-normal">—</span>}</DetailField>
                </DetailCard>
              </div>
            </div>
          </div>

          {/* Images — 7 slots */}
          <div className="flex-1 min-h-0 p-4 grid grid-cols-2 md:grid-cols-4 lg:grid-cols-7 gap-4 h-full">
            {[
              { key: 'Left1',  label: 'Left',         accent: '#0e4a78' },
              { key: 'Right1', label: 'Right',        accent: '#0e4a78' },
              { key: 'Back1',  label: 'Back',         accent: '#0e4a78' },
              { key: 'Top1',   label: 'Top',          accent: '#0e4a78' },
              { key: 'Left2',  label: 'Left Damage',  accent: '#dc2626' },
              { key: 'Right2', label: 'Right Damage', accent: '#dc2626' },
              { key: 'Top2',   label: 'Top Damage',   accent: '#dc2626' },
            ].map(({ key, label, accent }) => (
              <ImageSlot key={key} url={getWbinImageUrl(row[key])} label={label} accent={accent} onZoom={onZoom} />
            ))}
          </div>
        </div>
      </div>
    </div>
  )
}

// ── Today helper ──────────────────────────────────────────────────────────────
function todayStr() {
  const d = new Date()
  const p = (n) => String(n).padStart(2, '0')
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`
}

// ── Main component ────────────────────────────────────────────────────────────
export default function WbInSurvey() {
  const today = todayStr()
  const [fromDate,    setFromDate]    = useState(today)
  const [toDate,      setToDate]      = useState(today)
  const [containerNo, setContainerNo] = useState('')
  const [sortCol,     setSortCol]     = useState('SurveyTime')
  const [sortDir,     setSortDir]     = useState('desc')
  const [lightbox,    setLightbox]    = useState(null)   // { url, label }
  const [detailIdx,   setDetailIdx]   = useState(null)   // row index

  const [fetchWbin, { data, isFetching, isError }] = useLazyGetWbinSurveyQuery()

  // Load today on mount
  useEffect(() => {
    fetchWbin({ from_date: today, to_date: today })
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  const handleSearch = () => {
    fetchWbin({
      ...(fromDate    ? { from_date:    fromDate }    : {}),
      ...(toDate      ? { to_date:      toDate }      : {}),
      ...(containerNo.trim() ? { container_no: containerNo.trim() } : {}),
    })
  }

  const handleClear = () => {
    setFromDate(today); setToDate(today); setContainerNo('')
    fetchWbin({ from_date: today, to_date: today })
  }

  const toggleSort = (col) => {
    setSortCol(col)
    setSortDir(prev => (sortCol === col && prev === 'desc') ? 'asc' : 'desc')
  }

  const rows = (() => {
    const raw = Array.isArray(data?.data) ? [...data.data] : []
    return raw.sort((a, b) => {
      const av = String(a[sortCol] ?? ''), bv = String(b[sortCol] ?? '')
      return sortDir === 'asc' ? av.localeCompare(bv) : bv.localeCompare(av)
    })
  })()

  const surveyDone    = rows.filter(r => r.IsSuveryCompleted === true || r.IsSuveryCompleted === 1).length
  const surveyPending = rows.length - surveyDone

  const handleExport = () => {
    if (!rows.length) return
    const sheetData = rows.map((r, i) => ({
      '#':              i + 1,
      'Container No':   r.ContNo            || '',
      'Size':           r.ContainerSize     || '',
      'Type':           r.ContainerType     || '',
      'Process':        r.Process           || '',
      'Shipping Line':  r.ShippingLine      || '',
      'Booking No':     r.BookingNo         || '',
      'Document No':    r.DocumentNo        || '',
      'Status':         r.ContainerStatus   || '',
      'Mode':           r.Mode              || '',
      'Location':       r.ContainerLocationName || '',
      'Trailer No':     r.TrailerNo         || '',
      'Gate Name':      r.GateName          || '',
      'Survey Time':    r.SurveyTime        ? formatDate(r.SurveyTime)    : '',
      'Detected Time':  r.DetectedTime      ? formatDate(r.DetectedTime)  : '',
      'Survey Date':    r.SurveyDate        || '',
      'Gate In':        r.GateInDate        ? formatDate(r.GateInDate)    : '',
      'Gate Out':       r.GateOutDate       ? formatDate(r.GateOutDate)   : '',
      'Survey Done':    (r.IsSuveryCompleted === true || r.IsSuveryCompleted === 1) ? 'Yes' : 'No',
    }))
    const ws = XLSX.utils.json_to_sheet(sheetData)
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, 'WB In Survey')
    writeXlsx(wb, `WbInSurvey_${fromDate || 'today'}_${toDate || 'today'}.xlsx`)
  }

  const TH = ({ col, children }) => (
    <th onClick={() => toggleSort(col)}
      className="px-3 py-3 text-left text-xs font-semibold uppercase tracking-wider cursor-pointer select-none whitespace-nowrap hover:bg-white/10 transition-colors">
      <span className="flex items-center gap-0.5">{children}<SortIcon col={col} sortCol={sortCol} sortDir={sortDir} /></span>
    </th>
  )

  return (
    <div className="w-full min-h-screen relative overflow-hidden bg-cover bg-center" style={{ backgroundImage: "url('/Images/bgimageold.png')" }}>
      <div className="absolute inset-0 bg-white/70 backdrop-blur-[1px]" />
      <div className="relative z-10 flex flex-col min-h-screen">
        <Navbar />

        <main className="flex-1 px-4 sm:px-6 pb-10">

          {/* Header */}
          <header className="pt-6 pb-4">
            <p className="text-xs uppercase tracking-[0.2em] text-slate-500 font-semibold">Gate Management</p>
            <h1 className="text-2xl sm:text-3xl font-bold text-[#0e4a78] flex items-center gap-3 mt-1">
              <span className="w-10 h-10 rounded-2xl bg-[#0e4a78] flex items-center justify-center shadow-lg shadow-[#0e4a78]/30">
                <FaTrain className="text-white" size={18} />
              </span>
              WB In Survey
            </h1>
          </header>

          {/* Stat cards */}
          <StatGrid className="mb-5">
            <StatCard label="Total Records" value={rows.length} icon={FiPackage}   color="#0e4a78" loading={isFetching} />
            <StatCard label="Survey Done"   value={surveyDone}    icon={FiCheckCircle} color="#059669" loading={isFetching} />
            <StatCard label="Pending"       value={surveyPending} icon={FiAlertCircle} color="#d97706" loading={isFetching} />
          </StatGrid>

          {/* Filters */}
          <FilterCard className="mb-5">
            <FilterBar>
              <FilterField label="From Date">
                <div className="relative">
                  <FiCalendar className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={14} />
                  <input type="date" value={fromDate} onChange={e => setFromDate(e.target.value)}
                    className="pl-9 pr-3 py-2 text-sm border border-slate-200 rounded-xl bg-white focus:outline-none focus:ring-2 focus:ring-[#0e4a78]/30 w-full" />
                </div>
              </FilterField>
              <FilterField label="To Date">
                <div className="relative">
                  <FiCalendar className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={14} />
                  <input type="date" value={toDate} onChange={e => setToDate(e.target.value)}
                    className="pl-9 pr-3 py-2 text-sm border border-slate-200 rounded-xl bg-white focus:outline-none focus:ring-2 focus:ring-[#0e4a78]/30 w-full" />
                </div>
              </FilterField>
              <FilterField label="Container No">
                <div className="relative">
                  <FiSearch className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={14} />
                  <input type="text" value={containerNo} onChange={e => setContainerNo(e.target.value)}
                    onKeyDown={e => e.key === 'Enter' && handleSearch()}
                    placeholder="e.g. GESU1234567"
                    className="pl-9 pr-3 py-2 text-sm border border-slate-200 rounded-xl bg-white focus:outline-none focus:ring-2 focus:ring-[#0e4a78]/30 w-full" />
                </div>
              </FilterField>
              <div className="flex items-end gap-2">
                <FilterSearchBtn onClick={handleSearch} loading={isFetching} />
                <FilterClearBtn  onClick={handleClear} />
                <FilterExportBtn onClick={handleExport} disabled={!rows.length} />
              </div>
            </FilterBar>
          </FilterCard>

          {/* Table */}
          <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead>
                  <tr className="bg-gradient-to-r from-[#0e4a78] to-[#0b3e66] text-white">
                    <th className="px-3 py-3 text-left text-xs font-semibold uppercase tracking-wider w-10">#</th>
                    <TH col="ContNo">Container</TH>
                    <TH col="ContainerSize">Size/Type</TH>
                    <TH col="Process">Process</TH>
                    <TH col="ShippingLine">Shipping Line</TH>
                    <TH col="SurveyTime">Survey Time</TH>
                    <TH col="GateInDate">Gate In</TH>
                    <TH col="ContainerStatus">Status</TH>
                    <TH col="ContainerLocationName">Location</TH>
                    <TH col="IsSuveryCompleted">Surveyed</TH>
                    <th className="px-3 py-3 text-center text-xs font-semibold uppercase tracking-wider">Images</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {isFetching ? (
                    <tr><td colSpan={11} className="py-16 text-center text-slate-400 text-sm">Loading…</td></tr>
                  ) : isError ? (
                    <tr><td colSpan={11} className="py-16 text-center text-red-400 text-sm">Failed to load data</td></tr>
                  ) : rows.length === 0 ? (
                    <tr><td colSpan={11} className="py-16 text-center text-slate-400 text-sm">No records found</td></tr>
                  ) : rows.map((row, i) => {
                    const surveyed = row.IsSuveryCompleted === true || row.IsSuveryCompleted === 1
                    return (
                      <tr key={row.ContMasterID ?? i}
                        className="hover:bg-blue-50/40 transition-colors cursor-pointer"
                        onClick={() => setDetailIdx(i)}>
                        <td className="px-3 py-2.5 text-slate-400 text-xs">{i + 1}</td>
                        <td className="px-3 py-2.5 font-mono font-semibold text-[#0e4a78] text-xs whitespace-nowrap">{row.ContNo || '—'}</td>
                        <td className="px-3 py-2.5 text-slate-600 text-xs whitespace-nowrap">{`${row.ContainerSize || '—'} ${row.ContainerType || ''}`.trim()}</td>
                        <td className="px-3 py-2.5 text-slate-600 text-xs">{row.Process || '—'}</td>
                        <td className="px-3 py-2.5 text-slate-600 text-xs">{row.ShippingLine || '—'}</td>
                        <td className="px-3 py-2.5"><DateCell raw={row.SurveyTime} /></td>
                        <td className="px-3 py-2.5"><DateCell raw={row.GateInDate} /></td>
                        <td className="px-3 py-2.5 text-xs text-slate-600">{row.ContainerStatus || '—'}</td>
                        <td className="px-3 py-2.5 text-xs text-slate-600">{row.ContainerLocationName || '—'}</td>
                        <td className="px-3 py-2.5 text-center">
                          <span className={`inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-bold ${surveyed ? 'bg-emerald-100 text-emerald-700' : 'bg-amber-100 text-amber-700'}`}>
                            {surveyed ? <FiCheckCircle size={9} /> : <FiAlertCircle size={9} />}
                            {surveyed ? 'Yes' : 'No'}
                          </span>
                        </td>
                        <td className="px-3 py-2.5" onClick={e => e.stopPropagation()}>
                          <ImagesCell row={row} onOpen={(url, label) => setLightbox({ url, label })} />
                        </td>
                      </tr>
                    )
                  })}
                </tbody>
              </table>
            </div>
            {rows.length > 0 && (
              <div className="px-4 py-3 border-t border-slate-100 text-xs text-slate-400 flex items-center justify-between">
                <span>{rows.length} record{rows.length !== 1 ? 's' : ''}</span>
                <span>Click a row to view details</span>
              </div>
            )}
          </div>
        </main>

        <Footer />
      </div>

      {/* Lightbox */}
      {lightbox && <ImageLightbox url={lightbox.url} label={lightbox.label} onClose={() => setLightbox(null)} />}

      {/* Detail modal */}
      {detailIdx !== null && rows[detailIdx] && (
        <DetailModal
          row={rows[detailIdx]}
          index={detailIdx}
          total={rows.length}
          onClose={() => setDetailIdx(null)}
          onPrev={() => setDetailIdx(i => Math.max(0, i - 1))}
          onNext={() => setDetailIdx(i => Math.min(rows.length - 1, i + 1))}
          onZoom={(url, label) => setLightbox({ url, label })}
        />
      )}
    </div>
  )
}
