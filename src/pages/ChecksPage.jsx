import { useState, useEffect } from 'react'
import { Plus, Loader2, CheckCircle2, AlertTriangle, Upload } from 'lucide-react'

export default function ChecksPage() {
  const [checks,  setChecks]  = useState([])
  const [loading, setLoading] = useState(true)
  const [modal,   setModal]   = useState(false)

  // useEffect runs after the component mounts — fetch checks from backend
  useEffect(() => {
    fetch('/api/checks')
      .then(r => r.json())
      .then(d => { setChecks(d.checks); setLoading(false) })
  }, [])   // empty array = run once on mount only

  return (
    <div className="p-8 max-w-3xl mx-auto">

      <div className="flex items-center justify-between mb-8">
        <div>
          <h2 className="text-2xl font-bold text-[#4B3F72]">Checklist</h2>
          <p className="text-sm text-[#4B3F72]/40 mt-1">
            {checks.length} checks · runs on every plan automatically
          </p>
        </div>
        <button
          onClick={() => setModal(true)}
          className="flex items-center gap-2 bg-[#EFEAFA] hover:bg-blue-500
                     text-[#4B3F72] text-sm font-semibold px-4 py-2.5 rounded-xl transition"
        >
          <Plus size={15} /> Add Check
        </button>
      </div>

      {loading ? (
        <div className="flex items-center gap-2 text-[#4B3F72]/30 text-sm">
          <Loader2 size={14} className="animate-spin" /> Loading…
        </div>
      ) : (
        <div className="space-y-2">
          {checks.map((c, i) => (
            <div key={i}
              className="flex items-start gap-4 bg-[#FBFAFE] border border-white/5 rounded-xl px-5 py-4">
              <span className="text-xs font-mono text-[#4B3F72]/20 mt-0.5 w-6 shrink-0">
                #{i+1}
              </span>
              <p className="text-sm text-[#4B3F72]/70">{c.label}</p>
            </div>
          ))}
        </div>
      )}

      {modal && (
        <AddCheckModal
          onClose={() => setModal(false)}
          onAdded={(c) => { setChecks(prev => [...prev, c]); setModal(false) }}
        />
      )}
    </div>
  )
}

function AddCheckModal({ onClose, onAdded }) {
  const [description, setDescription] = useState('')
  const [threshold,   setThreshold]   = useState('')
  const [failType,    setFailType]     = useState('business_alert')
  const [outputFile,  setOutputFile]   = useState(null)
  const [status,      setStatus]       = useState(null)
  const [jobId,       setJobId]        = useState(null)
  const [events,      setEvents]       = useState([])

  async function handleSubmit() {
    if (!description || !threshold || !outputFile) return
    setStatus('running')

    const form = new FormData()
    form.append('description', description)
    form.append('threshold',   threshold)
    form.append('fail_type',   failType)
    form.append('output_file', outputFile)

    const res = await fetch('/api/add-check', { method: 'POST', body: form })
    const { job_id } = await res.json()
    setJobId(job_id)

    const es = new EventSource(`/api/status/${job_id}`)
    es.onmessage = (e) => {
      const event = JSON.parse(e.data)
      setEvents(prev => [...prev, event])
      if (event.type === 'done')  { setStatus('done');  es.close(); onAdded({ label: description }) }
      if (event.type === 'error') { setStatus('error'); es.close() }
    }
  }

  return (
    <div className="fixed inset-0 bg-black/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
      <div className="bg-[#FBFAFE] border border-[#DCD3F0] rounded-2xl w-full max-w-lg shadow-2xl">

        {/* Header */}
        <div className="flex items-center justify-between px-6 py-5 border-b border-white/5">
          <h3 className="text-base font-semibold text-[#4B3F72]">Add New Check</h3>
          <button onClick={onClose} className="text-[#4B3F72]/30 hover:text-[#4B3F72] transition">✕</button>
        </div>

        <div className="px-6 py-6 space-y-5">
          {status === null && <>

            {/* Description */}
            <div>
              <label className="block text-xs font-semibold text-[#4B3F72]/40 uppercase tracking-wider mb-2">
                What should be checked?
              </label>
              <textarea
                rows={3}
                value={description}
                onChange={e => setDescription(e.target.value)}
                placeholder="e.g. HIT Liquid pump production should never exceed 500,000 units per month"
                className="w-full border-[#EFEAFA] border border-[#DCD3F0] rounded-lg px-4 py-3
                           text-sm text-[#4B3F72] placeholder:text-[#4B3F72]/20 resize-none
                           focus:outline-none focus:border-[#8B7FC7] transition"
              />
            </div>

            {/* Threshold */}
            <div>
              <label className="block text-xs font-semibold text-[#4B3F72]/40 uppercase tracking-wider mb-2">
                Threshold / rule
              </label>
              <input
                type="text"
                value={threshold}
                onChange={e => setThreshold(e.target.value)}
                placeholder="e.g. ≤ 500,000 units per month"
                className="w-full border-[#EFEAFA] border border-[#DCD3F0] rounded-lg px-4 py-2.5
                           text-sm text-[#4B3F72] placeholder:text-[#4B3F72]/20
                           focus:outline-none focus:border-[#8B7FC7] transition"
              />
            </div>

            {/* Fail type toggle */}
            <div>
              <label className="block text-xs font-semibold text-[#4B3F72]/40 uppercase tracking-wider mb-2">
                If it fails, this is a…
              </label>
              <div className="flex gap-3">
                {[
                  { val: 'business_alert', label: 'Business Alert', sub: 'flag only'    },
                  { val: 'calc_error',     label: 'Calc Error',     sub: 'auto-fix'     },
                ].map(opt => (
                  <button
                    key={opt.val}
                    onClick={() => setFailType(opt.val)}
                    className={`flex-1 text-left px-4 py-3 rounded-xl border text-sm transition
                      ${failType === opt.val
                        ? 'border-blue-500 bg-[#EFEAFA]/10 text-[#4B3F72]'
                        : 'border-[#DCD3F0] text-[#4B3F72]/40 hover:border-white/20'
                      }`}
                  >
                    <p className="font-medium">{opt.label}</p>
                    <p className="text-xs opacity-60 mt-0.5">{opt.sub}</p>
                  </button>
                ))}
              </div>
            </div>

            {/* Output file upload */}
            <div>
              <label className="block text-xs font-semibold text-[#4B3F72]/40 uppercase tracking-wider mb-2">
                Current output Excel
              </label>
              <label className="flex items-center gap-3 border border-dashed border-[#DCD3F0]
                                hover:border-white/20 rounded-xl px-4 py-3 cursor-pointer transition">
                <Upload size={14} className="text-[#4B3F72]/30 shrink-0" />
                <span className="text-sm text-[#4B3F72]/40 truncate">
                  {outputFile ? outputFile.name : 'Upload Integrated_Supply_Plan.xlsx'}
                </span>
                <input type="file" accept=".xlsx" className="hidden"
                  onChange={e => setOutputFile(e.target.files[0])} />
              </label>
            </div>

            <button
              onClick={handleSubmit}
              disabled={!description || !threshold || !outputFile}
              className="w-full bg-[#EFEAFA] hover:bg-blue-500 disabled:bg-white/10
                         disabled:text-[#4B3F72]/20 text-[#4B3F72] font-semibold py-3
                         rounded-xl text-sm transition"
            >
              Add Check & Run
            </button>
          </>}

          {/* Live log while running */}
          {status === 'running' && (
            <div className="space-y-1.5 font-mono text-xs">
              {events.map((e, i) => (
                <div key={i} className="flex items-start gap-2 text-[#4B3F72]/40">
                  <span>·</span>
                  <span>{e.message || e.step || e.type}</span>
                </div>
              ))}
              {!events.length && (
                <div className="flex items-center gap-2 text-[#4B3F72]/30">
                  <Loader2 size={12} className="animate-spin" /> Starting…
                </div>
              )}
            </div>
          )}

          {status === 'done' && (
            <div className="flex items-center gap-2 text-emerald-400 text-sm font-medium">
              <CheckCircle2 size={15} /> Check added and Excel updated.
            </div>
          )}

          {status === 'error' && (
            <div className="flex items-center gap-2 text-red-400 text-sm">
              <AlertTriangle size={14} /> Something went wrong.
            </div>
          )}
        </div>

        {status === 'done' && jobId && (
          <div className="px-6 pb-6">
            
            <a  href={`/api/download/${jobId}`}
              download
              className="block w-full text-center bg-teal-700/30 hover:bg-teal-700/50
                         border border-teal-500/20 text-teal-300 font-semibold
                         py-3 rounded-xl text-sm transition"
            >
              Download Updated Excel
            </a>
          </div>
        )}

      </div>
    </div>
  )
}