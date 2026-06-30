import { Loader2, CheckCircle2, AlertTriangle } from 'lucide-react'

export default function RunStatus({ events, status }) {

  // Filter to only the events worth showing
  const visible = events.filter(e =>
    ['step','step_done','fix_applied','claude','error','check_results'].includes(e.type)
  )

  return (
    <div className="bg-white/[0.02] border border-white/5 rounded-2xl p-5">

      {/* Header */}
      <div className="flex items-center justify-between mb-4">
        <h3 className="text-sm font-semibold text-white">Pipeline Log</h3>
        {status === 'running' && (
          <span className="flex items-center gap-1.5 text-xs text-blue-400">
            <Loader2 size={11} className="animate-spin" /> Running
          </span>
        )}
        {status === 'done' && (
          <span className="flex items-center gap-1.5 text-xs text-emerald-400">
            <CheckCircle2 size={11} /> Complete
          </span>
        )}
        {status === 'error' && (
          <span className="flex items-center gap-1.5 text-xs text-red-400">
            <AlertTriangle size={11} /> Error
          </span>
        )}
      </div>

      {/* Event log */}
      <div className="space-y-1.5 font-mono text-xs max-h-48 overflow-y-auto">
        {visible.map((e, i) => {

          // Check results get special treatment
          if (e.type === 'check_results') {
            return (
              <div key={i} className="text-white/40 pl-4">
                <span className="text-emerald-400 font-semibold">{e.passed}</span>
                /{e.total} checks passing · loop {e.loop}
              </div>
            )
          }

          // Colour each event type differently
          const color = {
            step:        'text-white/40',
            step_done:   'text-white/60',
            fix_applied: 'text-amber-300/80',
            claude:      'text-purple-300/80',
            error:       'text-red-400',
          }[e.type] || 'text-white/30'

          const icon = {
            step:        '⟳',
            step_done:   '✓',
            fix_applied: '⚡',
            claude:      '◆',
            error:       '✗',
          }[e.type] || '·'

          return (
            <div key={i} className={`flex items-start gap-2 ${color}`}>
              <span className="shrink-0 mt-0.5">{icon}</span>
              <span>{e.message || e.step}</span>
            </div>
          )
        })}

        {!visible.length && (
          <p className="text-white/20">Waiting for pipeline to start…</p>
        )}
      </div>
    </div>
  )
}