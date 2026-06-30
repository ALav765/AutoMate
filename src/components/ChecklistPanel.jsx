import { useState } from 'react'
import { CheckCircle2, AlertTriangle, XCircle, ChevronDown, ChevronUp } from 'lucide-react'

export default function ChecklistPanel({ results }) {
  const entries = Object.entries(results).sort(([a],[b]) => +a - +b)
  const passed  = entries.filter(([,r]) => r.passed).length

  return (
    <div className="bg-[#FBFAFE] border border-white/5 rounded-2xl overflow-hidden">

      {/* Summary */}
      <div className="flex items-center justify-between px-5 py-4 border-b border-white/5">
        <h3 className="text-sm font-semibold text-[#4B3F72]">Checklist Results</h3>
        <div className="flex items-center gap-3 text-xs">
          <span className="text-emerald-400 font-semibold">{passed} passing</span>
          {entries.length - passed > 0 && (
            <span className="text-red-400 font-semibold">
              {entries.length - passed} failing
            </span>
          )}
        </div>
      </div>

      {/* Individual check rows */}
      <div className="divide-y divide-[#EFEAFA]">
        {entries.map(([k, r]) => <CheckRow key={k} num={k} result={r} />)}
      </div>

    </div>
  )
}

function CheckRow({ num, result }) {
  const [open, setOpen] = useState(!result.passed)

  const isFloor = result.passed && result.diagnosis?.includes('FLOOR ALERT')
  const isFail  = !result.passed

  // Pick colour scheme based on status
  const scheme = isFloor ? 'amber' : isFail ? 'red' : 'green'
  const colors = {
    green: { bg: 'bg-emerald-500/5', icon: 'text-emerald-400', badge: 'bg-emerald-500/10 text-emerald-300' },
    amber: { bg: 'bg-amber-500/5',   icon: 'text-amber-400',   badge: 'bg-amber-500/10  text-amber-300'   },
    red:   { bg: 'bg-red-500/5',     icon: 'text-red-400',     badge: 'bg-red-500/10    text-red-300'     },
  }[scheme]

  const Icon  = isFloor ? AlertTriangle : isFail ? XCircle : CheckCircle2
  const badge = isFloor ? '⚠ Floor Alert' : isFail ? '✕ Fail' : '✓ Pass'

  return (
    <button
      onClick={() => setOpen(o => !o)}
      className={`w-full flex items-center gap-3 px-5 py-3 text-left transition
                  ${isFail || isFloor ? colors.bg : ''}`}
    >
      <Icon size={14} className={`${colors.icon} shrink-0`} />
      <span className="text-xs text-[#4B3F72]/30 font-mono w-5 shrink-0">#{num}</span>
      <span className="flex-1 text-sm text-[#4B3F72]/70">{result.label}</span>
      <span className={`text-xs px-2 py-0.5 rounded-full font-medium ${colors.badge}`}>
        {badge}
      </span>
      {open
        ? <ChevronUp size={13} className="text-[#4B3F72]/20 shrink-0" />
        : <ChevronDown size={13} className="text-[#4B3F72]/20 shrink-0" />
      }

      {/* Expanded detail */}
      {open && (
        <div className="w-full mt-2 ml-[44px] space-y-1" onClick={e => e.stopPropagation()}>
          {result.note      && <p className="text-xs text-[#4B3F72]/40">{result.note}</p>}
          {result.diagnosis && <p className={`text-xs ${colors.icon}`}>{result.diagnosis}</p>}
        </div>
      )}
    </button>
  )
}